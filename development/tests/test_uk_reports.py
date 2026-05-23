from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.core.models import Exchange, Period, PeriodType, Ticker
from plugins.uk import reports


def _ticker() -> Ticker:
    return Ticker(
        exchange=Exchange.UK,
        code="ULVR",
        name="Unilever PLC",
        external_id="549300MKFYEKVRWML317",
    )


def test_uk_list_filings_searches_nsm_keywords(monkeypatch) -> None:
    calls = []

    def fake_search(company, lei, year, headline):
        calls.append((company, lei, year, headline))
        return [
            {
                "_id": headline,
                "headline": f"{headline} 2024",
                "document_date": "2024-12-31",
                "download_link": f"{headline}.pdf",
            }
        ]

    monkeypatch.setattr(reports, "search_filings", fake_search)

    df = reports._list_filings(_ticker(), Period(year=2024, type=PeriodType.ANNUAL))

    assert list(df["_id"]) == ["Annual Report", "Annual Financial Report"]
    assert calls[0] == ("Unilever PLC", "549300MKFYEKVRWML317", 2025, "Annual Report")


def test_uk_select_filing_prefers_direct_pdf() -> None:
    df = pd.DataFrame(
        [
            {
                "_id": "zip",
                "headline": "Unilever PLC Annual Financial Report",
                "download_link": "NSM/Portal/report.zip",
                "html_link": "NSM/Portal/report.html",
                "publication_date": "2024-03-14T10:00:00Z",
            },
            {
                "_id": "pdf",
                "headline": "Unilever PLC Annual Report on Form 20-F",
                "download_link": "NSM/Portal/report.pdf",
                "html_link": "",
                "publication_date": "2024-03-13T10:00:00Z",
            },
        ]
    )

    row = reports._select_filing(df, Period(year=2024, type=PeriodType.ANNUAL))

    assert row is not None
    assert row["_id"] == "pdf"


def test_uk_download_row_streams_direct_pdf(monkeypatch, tmp_path: Path) -> None:
    class _Client:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

    def fake_stream_to_file(_client, _url, dest, **_kwargs):
        Path(dest).write_bytes(b"%PDF\nbody")
        return 9

    monkeypatch.setattr(reports, "default_client", lambda **_kwargs: _Client())
    monkeypatch.setattr(reports, "stream_to_file", fake_stream_to_file)
    monkeypatch.setattr(reports, "nsm_rate", lambda: 99.0)

    source_url, source_format, n_bytes = reports._download_row_as_pdf(
        {"download_link": "NSM/Portal/report.pdf"}, tmp_path / "report.pdf"
    )

    assert source_url == "https://data.fca.org.uk/artefacts/NSM/Portal/report.pdf"
    assert source_format == "pdf"
    assert n_bytes == 9


def test_uk_download_row_renders_html_for_zip(monkeypatch, tmp_path: Path) -> None:
    calls = {}

    def fake_render(url, dest):
        calls["url"] = url
        Path(dest).write_bytes(b"%PDF\nbody")
        return 9

    monkeypatch.setattr(reports, "_render_url_to_pdf", fake_render)

    source_url, source_format, n_bytes = reports._download_row_as_pdf(
        {
            "download_link": "NSM/Portal/report.zip",
            "html_link": "NSM/Portal/reports/report.html",
        },
        tmp_path / "report.pdf",
    )

    assert calls["url"] == "https://data.fca.org.uk/artefacts/NSM/Portal/reports/report.html"
    assert source_url == calls["url"]
    assert source_format == "html"
    assert n_bytes == 9


def test_uk_download_reports_materializes_report_file(monkeypatch, tmp_path: Path) -> None:
    df = pd.DataFrame(
        [
            {
                "_id": "NI-000091833-0",
                "headline": "Unilever PLC Annual Report on Form 20-F 2023",
                "download_link": "NSM/Portal/report.pdf",
                "publication_date": "2024-03-14T11:52:00.000Z",
                "document_date": "2023-12-31",
                "type": "Annual Financial Report",
            }
        ]
    )
    monkeypatch.setattr(reports, "_list_filings", lambda *_args: df)
    monkeypatch.setattr(
        reports,
        "_download_row_as_pdf",
        lambda _row, dest: (dest.write_bytes(b"%PDF\nbody") and "https://nsm/pdf", "pdf", 9),
    )

    result = reports.download(_ticker(), Period(year=2024, type=PeriodType.ANNUAL), tmp_path)

    assert result[0].kind == "annual_report"
    assert result[0].accession_number == "NI-000091833-0"
    assert result[0].filing_date == "2024-03-14"
    assert Path(result[0].local_path).read_bytes().startswith(b"%PDF")

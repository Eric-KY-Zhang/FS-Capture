from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.core.models import Exchange, Period, PeriodType, Ticker
from plugins.jp import reports


def _ticker() -> Ticker:
    return Ticker(exchange=Exchange.JP, code="7203", name="トヨタ自動車株式会社", external_id="E02144")


def test_jp_list_filings_matches_ticker_and_annual_doc_type(monkeypatch) -> None:
    rows = [
        {
            "doc_id": "S100A",
            "doc_type_code": "120",
            "submit_date_time": "2024-06-25 15:00",
            "period_end": "2024-03-31",
            "edinet_code": "E02144",
            "sec_code": "7203",
            "filer_name": "トヨタ自動車株式会社",
            "title": "有価証券報告書",
        },
        {
            "doc_id": "S100B",
            "doc_type_code": "140",
            "submit_date_time": "2024-08-01 15:00",
            "edinet_code": "E02144",
            "sec_code": "7203",
        },
        {
            "doc_id": "S100C",
            "doc_type_code": "120",
            "submit_date_time": "2024-06-25 15:00",
            "edinet_code": "E99999",
            "sec_code": "9999",
        },
    ]
    monkeypatch.setattr(reports, "_scan_dates", lambda _period: ["2024-06-25"])
    monkeypatch.setattr(reports, "_list_documents", lambda _date: rows)

    df = reports._list_filings(_ticker(), Period(year=2024, type=PeriodType.ANNUAL))

    assert list(df["doc_id"]) == ["S100A"]


def test_jp_select_filing_prefers_doc_type_120() -> None:
    df = pd.DataFrame(
        [
            {"doc_id": "S100AMEND", "doc_type_code": "130", "submit_date_time": "2024-06-26"},
            {"doc_id": "S100MAIN", "doc_type_code": "120", "submit_date_time": "2024-06-25"},
        ]
    )

    row = reports._select_filing(df, Period(year=2024, type=PeriodType.ANNUAL))

    assert row is not None
    assert row["doc_id"] == "S100MAIN"


def test_jp_download_reports_materializes_report_file(monkeypatch, tmp_path: Path) -> None:
    df = pd.DataFrame(
        [
            {
                "doc_id": "S100TEST",
                "doc_type_code": "120",
                "submit_date_time": "2024-06-25 15:00",
                "period_end": "2024-03-31",
                "title": "有価証券報告書",
            }
        ]
    )
    monkeypatch.setattr(reports, "_list_filings", lambda *_args: df)
    monkeypatch.setattr(
        reports,
        "_download_doc_as_pdf",
        lambda _doc_id, dest: (dest.write_bytes(b"%PDF\nbody") and "https://edinet/pdf", "pdf", 9),
    )

    result = reports.download(_ticker(), Period(year=2024, type=PeriodType.ANNUAL), tmp_path)

    assert result[0].kind == "annual_report"
    assert result[0].accession_number == "S100TEST"
    assert Path(result[0].local_path).read_bytes().startswith(b"%PDF")


def test_jp_list_documents_uses_api_branch_when_key_configured(monkeypatch) -> None:
    calls = []

    monkeypatch.setattr(reports, "_edinet_api_key", lambda: "secret")
    monkeypatch.setattr("plugins.jp.edinet_api.list_documents", lambda date, api_key: calls.append((date, api_key)) or [])

    assert reports._list_documents("2024-06-25") == []
    assert calls == [("2024-06-25", "secret")]


def test_jp_list_documents_uses_web_branch_without_key(monkeypatch) -> None:
    calls = []

    monkeypatch.setattr(reports, "_edinet_api_key", lambda: "")
    monkeypatch.setattr("plugins.jp.edinet_web.list_documents", lambda date: calls.append(date) or [])

    assert reports._list_documents("2024-06-25") == []
    assert calls == ["2024-06-25"]

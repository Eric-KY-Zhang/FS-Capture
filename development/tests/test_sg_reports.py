from __future__ import annotations

from pathlib import Path

from app.core.models import Exchange, Period, PeriodType, Ticker
from plugins.sg import reports


def _ticker(name: str = "DBS GROUP HOLDINGS LTD", code: str = "D05") -> Ticker:
    return Ticker(exchange=Exchange.SG, code=code, name=name, external_id="SG1L01001701")


def test_download_annual_report(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr(
        reports,
        "_financial_reports",
        lambda: [
            {
                "id": "annual-2024",
                "documentDate": 1735516800000,
                "securityName": "DBS GROUP HOLDINGS LTD",
                "companyName": "DBS GROUP HOLDINGS LTD",
                "title": "Annual Report",
                "url": "https://links.sgx.com/annual",
            }
        ],
    )
    monkeypatch.setattr(
        reports,
        "_download_page_as_pdf",
        lambda _row, dest, _kind: (dest.write_bytes(b"%PDF\nbody") and "https://sgx/pdf", "pdf", 9),
    )

    result = reports.download(_ticker(), Period(year=2024, type=PeriodType.ANNUAL), tmp_path)

    assert len(result) == 1
    assert result[0].kind == "annual_report"
    assert result[0].accession_number == "annual-2024"
    assert Path(result[0].local_path).name == "SG_D05_DBS GROUP HOLDINGS LTD_2024_年报.pdf"


def test_download_h1_interim(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr(
        reports,
        "search_announcements",
        lambda *_args, **_kwargs: [
            {
                "id": "h1-2024",
                "submission_date": "20240801",
                "category_name": "Financial Statements",
                "title": "Financial Statements and Related Announcement::Half Yearly Results",
                "url": "https://links.sgx.com/h1",
            }
        ],
    )
    monkeypatch.setattr(
        reports,
        "_download_page_as_pdf",
        lambda _row, dest, _kind: (dest.write_bytes(b"%PDF\nbody") and "https://sgx/h1.pdf", "pdf", 9),
    )

    result = reports.download(
        _ticker("UNITED OVERSEAS BANK LIMITED", "U11"),
        Period(year=2024, type=PeriodType.Q2),
        tmp_path,
    )

    assert result[0].kind == "interim_report"
    assert result[0].filing_date == "2024-08-01"
    assert Path(result[0].local_path).name.endswith("_2024_半年报.pdf")


def test_download_ipo_prospectus(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr(
        reports,
        "_ipo_prospectuses",
        lambda: [
            {
                "id": "3407",
                "closing_date": 1733443200000,
                "name": "LION-CM EM ASIA INDEX ETF",
                "url": "https://links.sgx.com/ipo",
                "status": "Completed",
            }
        ],
    )
    monkeypatch.setattr(
        reports,
        "_download_page_as_pdf",
        lambda _row, dest, _kind: (dest.write_bytes(b"%PDF\nbody") and "https://sgx/ipo.pdf", "pdf", 9),
    )

    result = reports.download(
        _ticker("LION-CM EM ASIA INDEX ETF", "3407"),
        Period(year=2024, type=PeriodType.IPO_PROSPECTUS),
        tmp_path,
    )

    assert result[0].kind == "ipo_prospectus"
    assert result[0].accession_number == "3407"
    assert "招股书" in Path(result[0].local_path).name


def test_download_returns_empty_when_no_sg_filing(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr(reports, "_financial_reports", lambda: [])

    assert reports.download(_ticker(), Period(year=2024, type=PeriodType.ANNUAL), tmp_path) == []


def test_choose_pdf_link_prefers_main_documents() -> None:
    assert reports._choose_pdf_link(
        "annual_report",
        [
            "https://links.sgx.com/Letter%20to%20Shareholders.pdf",
            "https://links.sgx.com/DBS%20Annual%20Report%202024.pdf",
        ],
    ).endswith("Report%202024.pdf")
    assert reports._choose_pdf_link(
        "interim_report",
        [
            "https://links.sgx.com/2Q24%20CEO%20presentation.pdf",
            "https://links.sgx.com/2Q24%20Condensed%20Interim%20Financial%20Statement.pdf",
        ],
    ).endswith("Statement.pdf")

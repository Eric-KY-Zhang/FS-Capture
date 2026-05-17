"""Unit tests for the TW plugin's pure logic (no network)."""

from __future__ import annotations

from pathlib import Path

import pytest

from app.core.models import Exchange, Period, PeriodType, Ticker
from plugins.tw import TWShare
from plugins.tw import reports as tw_reports
from plugins.tw.reports import (
    _annual_score,
    _build_pdf_url,
    _extract_pdf_links,
    _quarter_score,
    _select_annual,
    _select_quarter,
    roc_year,
)


def test_roc_year_converts_western_year() -> None:
    assert roc_year(2024) == 113
    assert roc_year(2025) == 114
    assert roc_year(1912) == 1


def test_roc_year_rejects_pre_republic_year() -> None:
    with pytest.raises(ValueError):
        roc_year(1900)


def test_extract_pdf_links_handles_mtype_F_filename_grammar() -> None:
    html = """
    readfile2("F","2330","2024_2330_20250603F04.pdf");
    readfile2("F","2330","2024_2330_20250603FE4.pdf");
    readfile2("F","2330","2025_2330_20250603F01.pdf");
    // dup
    readfile2("F","2330","2024_2330_20250603F04.pdf");
    """
    rows = _extract_pdf_links(html, "F")
    filenames = [r["filename"] for r in rows]
    assert filenames == [
        "2024_2330_20250603F04.pdf",
        "2024_2330_20250603FE4.pdf",
        "2025_2330_20250603F01.pdf",
    ]
    assert rows[0]["yyyy"] == "2024"
    assert rows[0]["typecode"] == "F04"
    assert rows[0]["filedate"] == "20250603"


def test_extract_pdf_links_handles_mtype_A_filename_grammar() -> None:
    html = """
    readfile2("A","2330","202401_2330_AI1.pdf");
    readfile2("A","2330","202402_2330_AI1.pdf");
    readfile2("A","2330","202403_2330_AIA.pdf");
    """
    rows = _extract_pdf_links(html, "A")
    assert [r["filename"] for r in rows] == [
        "202401_2330_AI1.pdf",
        "202402_2330_AI1.pdf",
        "202403_2330_AIA.pdf",
    ]
    assert rows[0]["yyyy"] == "2024"
    assert rows[0]["qq"] == "01"
    assert rows[0]["typecode"] == "AI1"


def test_select_annual_prefers_chinese_F04_over_english_FE4() -> None:
    rows = [
        {"filename": "2024_2330_20250603FE4.pdf", "yyyy": "2024", "typecode": "FE4"},
        {"filename": "2024_2330_20250603F04.pdf", "yyyy": "2024", "typecode": "F04"},
        {"filename": "2024_2330_20250603F01.pdf", "yyyy": "2024", "typecode": "F01"},
    ]
    chosen = _select_annual(rows, target_fy=2024)
    assert chosen is not None
    assert chosen["typecode"] == "F04"


def test_select_annual_returns_none_when_no_F04_match() -> None:
    rows = [
        {"filename": "2023_2330_20240604F04.pdf", "yyyy": "2023", "typecode": "F04"},
        {"filename": "2024_2330_20250603F01.pdf", "yyyy": "2024", "typecode": "F01"},
    ]
    # asking for FY 2024 annual, but only 2023 annual and 2024 meeting notice present
    assert _select_annual(rows, target_fy=2024) is None


def test_select_quarter_picks_correct_quarter_suffix() -> None:
    rows = [
        {"filename": "202401_2330_AI1.pdf", "yyyy": "2024", "qq": "01", "typecode": "AI1"},
        {"filename": "202402_2330_AI1.pdf", "yyyy": "2024", "qq": "02", "typecode": "AI1"},
        {"filename": "202403_2330_AI1.pdf", "yyyy": "2024", "qq": "03", "typecode": "AI1"},
    ]
    chosen = _select_quarter(rows, target_fy=2024, period_type=PeriodType.Q2)
    assert chosen is not None
    assert chosen["filename"] == "202402_2330_AI1.pdf"


def test_quarter_score_prefers_consolidated_chinese_over_parent_only() -> None:
    assert _quarter_score("AI1") < _quarter_score("AIA")
    assert _quarter_score("AIA") < _quarter_score("AE1")
    assert _quarter_score("XYZ") == 99


def test_annual_score_rejects_meeting_documents() -> None:
    assert _annual_score("F04") == 0
    assert _annual_score("FE4") == 1
    assert _annual_score("F01") == 99
    assert _annual_score("F13") == 99


def test_build_pdf_url_uses_step_9_with_mtype_letter_as_kind() -> None:
    url = _build_pdf_url("2330", "2024_2330_20250603F04.pdf", "F")
    assert "step=9" in url
    assert "kind=F" in url
    assert "co_id=2330" in url
    assert "filename=2024_2330_20250603F04.pdf" in url


def test_download_annual_returns_report_when_pdf_succeeds(monkeypatch, tmp_path: Path) -> None:
    """Wire up the happy path for annual: monkeypatch the listing fetch and
    PDF download; expect one ReportFile with kind=annual_report.
    """
    listing_html = """
    readfile2("F","2330","2024_2330_20250603F04.pdf");
    readfile2("F","2330","2024_2330_20250603FE4.pdf");
    """

    def fake_fetch(_co, _roc, _mtype):
        return listing_html

    monkeypatch.setattr(tw_reports, "_fetch_listing", fake_fetch)

    def fake_download(url: str, dest: Path) -> int:
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(b"%PDF-1.4\n%fake")
        return dest.stat().st_size

    monkeypatch.setattr(tw_reports, "_download_pdf", fake_download)

    ticker = Ticker(exchange=Exchange.TW, code="2330", name="台積電", external_id="2330")
    reports_out = TWShare().download_reports(
        ticker, Period(year=2024, type=PeriodType.ANNUAL), tmp_path
    )
    assert len(reports_out) == 1
    rep = reports_out[0]
    assert rep.kind == "annual_report"
    assert rep.title == "2024_2330_20250603F04.pdf"  # Chinese annual preferred
    assert "kind=F" in rep.source_url
    assert Path(rep.local_path).exists()


def test_download_quarter_picks_correct_quarter_pdf(monkeypatch, tmp_path: Path) -> None:
    listing_html = """
    readfile2("A","2330","202401_2330_AI1.pdf");
    readfile2("A","2330","202402_2330_AI1.pdf");
    readfile2("A","2330","202403_2330_AI1.pdf");
    """

    monkeypatch.setattr(tw_reports, "_fetch_listing", lambda *_a, **_k: listing_html)

    def fake_download(url: str, dest: Path) -> int:
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(b"%PDF-1.4\n%fake")
        return dest.stat().st_size

    monkeypatch.setattr(tw_reports, "_download_pdf", fake_download)

    ticker = Ticker(exchange=Exchange.TW, code="2330", name="台積電", external_id="2330")
    reports_out = TWShare().download_reports(
        ticker, Period(year=2024, type=PeriodType.Q2), tmp_path
    )
    assert len(reports_out) == 1
    rep = reports_out[0]
    assert rep.kind == "interim_report"
    assert rep.title == "202402_2330_AI1.pdf"


def test_download_pdf_retries_temporary_twse_pdf_failure(monkeypatch, tmp_path: Path) -> None:
    class FakeLimiter:
        def acquire_blocking(self) -> None:
            return None

    class FakeResponse:
        content = b"<html><a href='/pdf/202402_2330_AI1_123.pdf'>pdf</a></html>"

        def raise_for_status(self) -> None:
            return None

    class FakeClient:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def post(self, *_args, **_kwargs):
            return FakeResponse()

    attempts = {"count": 0}

    def fake_stream_to_file(_client, _url, dest, **_kwargs):
        attempts["count"] += 1
        if attempts["count"] == 1:
            raise RuntimeError("illegal status line")
        Path(dest).write_bytes(b"%PDF-1.4\n%fake")
        return Path(dest).stat().st_size

    monkeypatch.setattr(tw_reports, "default_client", lambda **_kwargs: FakeClient())
    monkeypatch.setattr(tw_reports, "stream_to_file", fake_stream_to_file)
    monkeypatch.setattr(tw_reports, "limiter", lambda *_args, **_kwargs: FakeLimiter())
    monkeypatch.setattr(tw_reports.time, "sleep", lambda _seconds: None)

    url = _build_pdf_url("2330", "202402_2330_AI1.pdf", "A")
    n_bytes = tw_reports._download_pdf(url, tmp_path / "tw-q2.pdf")

    assert attempts["count"] == 2
    assert n_bytes is not None
    assert n_bytes > 0


def test_download_reports_returns_empty_when_no_candidates(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr(tw_reports, "_fetch_listing", lambda *_a, **_k: "<html></html>")
    ticker = Ticker(exchange=Exchange.TW, code="9999", name="無此公司", external_id="9999")
    reports_out = TWShare().download_reports(
        ticker, Period(year=2024, type=PeriodType.ANNUAL), tmp_path
    )
    assert reports_out == []

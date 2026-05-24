from __future__ import annotations

import json

from app.core.models import Exchange, Period, PeriodType, ReportFile, Ticker
from app.core.sidecar import read_sidecar, sidecar_path, write_sidecar


def test_write_sidecar_contains_report_metadata(tmp_path) -> None:
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    cache_root = tmp_path / "cache"
    pdf_path = output_dir / "A_600519_贵州茅台_2024_年报.pdf"
    pdf_path.write_bytes(b"%PDF-1.7\nbody")
    ticker = Ticker(exchange=Exchange.A_SHARE, code="600519", name="贵州茅台")
    period = Period(year=2024, type=PeriodType.ANNUAL)
    report = ReportFile(
        ticker=ticker,
        period=period,
        kind="annual_report",
        local_path=str(pdf_path),
        source_url="https://example.com/report.pdf",
        title="2024 年度报告",
    )

    sidecar = write_sidecar(report, cache_root)
    meta = json.loads(sidecar.read_text(encoding="utf-8"))

    assert sidecar == cache_root / "sidecars" / "A" / "A_600519_贵州茅台_2024_年报.meta.json"
    assert not pdf_path.with_suffix(pdf_path.suffix + ".meta.json").exists()
    assert meta["exchange"] == "A"
    assert meta["ticker_code"] == "600519"
    assert meta["ticker_name"] == "贵州茅台"
    assert meta["period_year"] == 2024
    assert meta["period_type"] == "annual"
    assert meta["kind"] == "annual_report"
    assert meta["title"] == "2024 年度报告"
    assert meta["source_url"] == "https://example.com/report.pdf"
    assert meta["downloaded_at"]
    assert meta["file_size_bytes"] == len(b"%PDF-1.7\nbody")
    assert len(meta["sha256"]) == 64


def test_write_sidecar_replaces_existing_file_without_temp_leftovers(tmp_path) -> None:
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    cache_root = tmp_path / "cache"
    pdf_path = output_dir / "US_AAPL_Apple_2024_年报.pdf"
    pdf_path.write_bytes(b"%PDF-1.7\nfirst")
    ticker = Ticker(exchange=Exchange.US, code="AAPL", name="Apple")
    period = Period(year=2024, type=PeriodType.ANNUAL)
    report = ReportFile(
        ticker=ticker,
        period=period,
        kind="annual_report",
        local_path=str(pdf_path),
        source_url="https://example.com/first.pdf",
        title="First",
    )

    sidecar = write_sidecar(report, cache_root)
    first = json.loads(sidecar.read_text(encoding="utf-8"))

    pdf_path.write_bytes(b"%PDF-1.7\nsecond")
    report.source_url = "https://example.com/second.pdf"
    report.title = "Second"
    second_sidecar = write_sidecar(report, cache_root)
    second = json.loads(second_sidecar.read_text(encoding="utf-8"))

    assert second_sidecar == sidecar
    assert first["sha256"] != second["sha256"]
    assert second["source_url"] == "https://example.com/second.pdf"
    assert second["title"] == "Second"
    assert list(sidecar.parent.glob("*.tmp")) == []


def test_sidecar_path_prefers_accession_number(tmp_path) -> None:
    ticker = Ticker(exchange=Exchange.US, code="AAPL", name="Apple")
    period = Period(year=2024, type=PeriodType.ANNUAL)
    report = ReportFile(
        ticker=ticker,
        period=period,
        kind="annual_report",
        local_path=str(tmp_path / "AAPL.pdf"),
        source_url="https://example.com/aapl.pdf",
        accession_number="0000320193-24-000123",
    )

    assert sidecar_path(report, tmp_path) == (
        tmp_path / "sidecars" / "US" / "0000320193-24-000123.meta.json"
    )


def test_read_sidecar_returns_payload(tmp_path) -> None:
    path = tmp_path / "sidecars" / "HK" / "hk-00700.meta.json"
    path.parent.mkdir(parents=True)
    path.write_text('{"exchange": "HK", "ticker_code": "00700"}', encoding="utf-8")

    payload = read_sidecar("hk-00700", Exchange.HK, tmp_path)

    assert payload == {"exchange": "HK", "ticker_code": "00700"}
    assert read_sidecar("missing", Exchange.HK, tmp_path) is None

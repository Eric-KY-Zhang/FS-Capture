from __future__ import annotations

import json
from contextlib import suppress
from pathlib import Path

from app.core.models import Exchange, Period, PeriodType, ReportFile, Ticker
from app.core.sidecar import write_sidecar


def test_write_sidecar_contains_report_metadata() -> None:
    work_dir = Path.cwd() / "sidecar_test_dir"
    work_dir.mkdir(exist_ok=True)
    pdf_path = work_dir / "A_600519_2024_annual_annual_report.pdf"
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

    sidecar = write_sidecar(report)
    meta = json.loads(sidecar.read_text(encoding="utf-8"))

    assert sidecar.name == "A_600519_2024_annual_annual_report.pdf.meta.json"
    assert meta["exchange"] == "A"
    assert meta["ticker_code"] == "600519"
    assert meta["period_year"] == 2024
    assert meta["period_type"] == "annual"
    assert meta["kind"] == "annual_report"
    assert meta["file_size_bytes"] == len(b"%PDF-1.7\nbody")
    assert len(meta["sha256"]) == 64

    with suppress(FileNotFoundError, PermissionError):
        sidecar.unlink()
        pdf_path.unlink()
        work_dir.rmdir()

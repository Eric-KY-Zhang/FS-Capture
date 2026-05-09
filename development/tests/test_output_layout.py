from __future__ import annotations

from pathlib import Path

from app.core.models import Exchange, Period, PeriodType, Ticker
from app.core.output_paths import report_output_path


def test_report_path_is_flat_under_output_root() -> None:
    ticker = Ticker(exchange=Exchange.A_SHARE, code="600519")
    period = Period(year=2024, type=PeriodType.ANNUAL)

    path = report_output_path(Path("output"), ticker, period, "annual_report", ".pdf")

    assert path.parent == Path("output")
    assert path.name == "A_600519_2024_annual_annual_report.pdf"

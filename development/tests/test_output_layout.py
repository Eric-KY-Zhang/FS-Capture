from __future__ import annotations

import unittest
from pathlib import Path

from app.core.models import Exchange, Period, PeriodType, Ticker
from app.core.output_paths import report_output_path


class OutputLayoutTests(unittest.TestCase):
    def test_report_path_is_flat_under_output_root(self) -> None:
        ticker = Ticker(exchange=Exchange.A_SHARE, code="600519")
        period = Period(year=2024, type=PeriodType.ANNUAL)

        path = report_output_path(Path("output"), ticker, period, "annual_report", ".pdf")

        self.assertEqual(path.parent, Path("output"))
        self.assertEqual(path.name, "A_600519_2024_annual_annual_report.pdf")


if __name__ == "__main__":
    unittest.main()

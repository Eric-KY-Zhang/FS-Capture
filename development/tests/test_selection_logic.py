from __future__ import annotations

import unittest

import pandas as pd

from app.core.models import Period, PeriodType
from plugins.ashare.reports import _column_for as ashare_column_for
from plugins.kr.reports import _select_filing as select_kr_filing
from plugins.us.reports import _filter_table as filter_us_table


class AShareReportSelectionTests(unittest.TestCase):
    def test_beijing_stock_exchange_uses_cninfo_bj_column(self) -> None:
        self.assertEqual(ashare_column_for("430047"), "bj")
        self.assertEqual(ashare_column_for("430047.BJ"), "bj")


class USReportSelectionTests(unittest.TestCase):
    def test_non_calendar_fiscal_quarters_are_mapped_by_annual_cycle(self) -> None:
        table = {
            "form": ["10-K", "10-K", "10-Q", "10-Q", "10-Q"],
            "accessionNumber": ["a0", "a1", "q1", "q2", "q3"],
            "primaryDocument": ["a0.htm", "a1.htm", "q1.htm", "q2.htm", "q3.htm"],
            "filingDate": ["2022-10-28", "2023-11-03", "2023-02-03", "2023-05-05", "2023-08-04"],
            "reportDate": ["2022-09-24", "2023-09-30", "2022-12-31", "2023-04-01", "2023-07-01"],
        }

        q1 = filter_us_table(table, Period(year=2023, type=PeriodType.Q1))
        q2 = filter_us_table(table, Period(year=2023, type=PeriodType.Q2))
        q3 = filter_us_table(table, Period(year=2023, type=PeriodType.Q3))

        self.assertEqual(q1[0]["accessionNumber"], "q1")
        self.assertEqual(q2[0]["accessionNumber"], "q2")
        self.assertEqual(q3[0]["accessionNumber"], "q3")


class KRReportSelectionTests(unittest.TestCase):
    def test_quarterly_reports_choose_q1_and_q3_by_receipt_month(self) -> None:
        df = pd.DataFrame([
            {"rcept_no": "q1", "report_nm": "분기보고서 (2023.03)", "rcept_dt": "20230515"},
            {"rcept_no": "q3", "report_nm": "분기보고서 (2023.09)", "rcept_dt": "20231114"},
        ])

        q1 = select_kr_filing(df, Period(year=2023, type=PeriodType.Q1))
        q3 = select_kr_filing(df, Period(year=2023, type=PeriodType.Q3))

        self.assertEqual(q1["rcept_no"], "q1")
        self.assertEqual(q3["rcept_no"], "q3")


if __name__ == "__main__":
    unittest.main()

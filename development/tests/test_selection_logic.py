from __future__ import annotations

import unittest

import pandas as pd

from app.core.models import Period, PeriodType
from plugins.hk.financials import _row_to_lines as hk_row_to_lines
from plugins.hk.financials import _select_period_row as select_hk_period_row
from plugins.kr.reports import _select_filing as select_kr_filing
from plugins.us.reports import _filter_table as filter_us_table
from plugins.us.financials import _select_period_value as select_us_period_value


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


class HKFinancialParsingTests(unittest.TestCase):
    def test_long_table_rows_are_pivoted_to_financial_lines(self) -> None:
        df = pd.DataFrame([
            {"REPORT_DATE": "2024-12-31 00:00:00", "STD_ITEM_NAME": "营业额", "AMOUNT": 100.0},
            {"REPORT_DATE": "2024-12-31 00:00:00", "STD_ITEM_NAME": "股东权益", "AMOUNT": 60.0},
            {"REPORT_DATE": "2023-12-31 00:00:00", "STD_ITEM_NAME": "营业额", "AMOUNT": 80.0},
        ])

        row = select_hk_period_row(df, Period(year=2024, type=PeriodType.ANNUAL))
        lines = hk_row_to_lines(row)

        self.assertEqual(lines["营业额"], 100.0)
        self.assertEqual(lines["营业收入"], 100.0)
        self.assertEqual(lines["股东权益合计"], 60.0)


class USFinancialSelectionTests(unittest.TestCase):
    def test_companyfacts_prefers_target_accession_and_report_date(self) -> None:
        units = {
            "USD": [
                {"fy": 2023, "fp": "FY", "accn": "old", "end": "2022-09-24", "val": 365817},
                {"fy": 2023, "fp": "FY", "accn": "target", "end": "2023-09-30", "val": 383285, "frame": None},
                {"fy": 2023, "fp": "FY", "accn": "target", "end": "2022-09-24", "val": 394328},
            ]
        }

        value = select_us_period_value(
            units,
            Period(year=2023, type=PeriodType.ANNUAL),
            {"accessionNumber": "target", "reportDate": "2023-09-30"},
        )

        self.assertEqual(value, 383285)


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import datetime as dt

import pytest

from app.core.models import Exchange, Period, PeriodType, Ticker
from plugins.hk import reports as hk_reports
from plugins.hk.fiscal_year import fiscal_year_end_month
from plugins.hk.reports import _select_main as select_hk_main


def _row(
    title: str,
    filing_date: str,
    *,
    url: str,
    file_size: int = 2_000_000,
    code: str = "00700",
) -> dict:
    return {
        "url": url,
        "title": title,
        "headline": "Financial Statements/ESG Information - [Annual Report]",
        "doc_type": "Annual Report",
        "filing_date": dt.date.fromisoformat(filing_date),
        "stock_codes": (code.zfill(5),),
        "source_id": url.rsplit("/", 1)[-1],
        "file_size": file_size,
    }


@pytest.fixture(autouse=True)
def _mock_pdf_verify(monkeypatch: pytest.MonkeyPatch) -> None:
    verified = {
        "https://example.com/tencent-annual.pdf",
        "https://example.com/alibaba-annual.pdf",
        "https://example.com/hsbc-main.pdf",
        "https://example.com/aia-annual.pdf",
        "https://example.com/china-mobile-annual.pdf",
    }
    monkeypatch.setattr(
        hk_reports,
        "verify_pdf_year_and_kind",
        lambda url, _year, _kind: url in verified,
    )


def test_tencent_annual_prefers_main_report() -> None:
    ticker = Ticker(exchange=Exchange.HK, code="00700")
    rows = [
        _row("年報 2023", "2024-03-22", url="https://example.com/tencent-annual.pdf"),
        _row("ESG 報告 2023", "2024-04-15", url="https://example.com/tencent-esg.pdf"),
        _row(
            "補充公告：股東週年大會通告",
            "2024-05-01",
            url="https://example.com/tencent-supp.pdf",
            file_size=200_000,
        ),
    ]

    selected = select_hk_main(rows, ticker, Period(year=2023, type=PeriodType.ANNUAL))

    assert selected["url"] == "https://example.com/tencent-annual.pdf"


def test_alibaba_march_fiscal_year_accepts_same_year_filing_window() -> None:
    ticker = Ticker(exchange=Exchange.HK, code="09988")
    rows = [
        _row(
            "Annual Report 2024",
            "2024-07-25",
            url="https://example.com/alibaba-annual.pdf",
            code="09988",
        ),
        _row(
            "Annual Report 2024 Supplementary Information",
            "2024-08-01",
            url="https://example.com/alibaba-supp.pdf",
            code="09988",
        ),
        _row(
            "Annual Report 2023",
            "2023-07-21",
            url="https://example.com/alibaba-old.pdf",
            code="09988",
        ),
    ]

    selected = select_hk_main(rows, ticker, Period(year=2024, type=PeriodType.ANNUAL))

    assert selected["url"] == "https://example.com/alibaba-annual.pdf"


def test_hsbc_uses_pdf_verification_to_avoid_subsidiary_reports() -> None:
    ticker = Ticker(exchange=Exchange.HK, code="00005")
    rows = [
        _row(
            "HSBC Holdings plc Annual Report and Accounts 2023",
            "2024-02-21",
            url="https://example.com/hsbc-main.pdf",
            code="00005",
        ),
        _row(
            "HSBC Bank plc Annual Report 2023",
            "2024-02-21",
            url="https://example.com/hsbc-bank.pdf",
            code="00005",
        ),
        _row(
            "Pillar 3 Disclosures 2023",
            "2024-02-21",
            url="https://example.com/hsbc-pillar3.pdf",
            code="00005",
        ),
    ]

    selected = select_hk_main(rows, ticker, Period(year=2023, type=PeriodType.ANNUAL))

    assert selected["url"] == "https://example.com/hsbc-main.pdf"


def test_aia_filters_non_report_esg_candidate() -> None:
    ticker = Ticker(exchange=Exchange.HK, code="01299")
    rows = [
        _row("2023 ESG Report", "2024-03-15", url="https://example.com/aia-esg.pdf", code="01299"),
        _row(
            "2023 Annual Report",
            "2024-03-15",
            url="https://example.com/aia-annual.pdf",
            code="01299",
        ),
    ]

    selected = select_hk_main(rows, ticker, Period(year=2023, type=PeriodType.ANNUAL))

    assert selected["url"] == "https://example.com/aia-annual.pdf"


def test_china_mobile_accepts_mixed_chinese_english_title() -> None:
    ticker = Ticker(exchange=Exchange.HK, code="00941")
    rows = [
        _row(
            "2023 年報 Annual Report",
            "2024-03-21",
            url="https://example.com/china-mobile-annual.pdf",
            code="00941",
        ),
        _row(
            "2023 可持續發展報告 Sustainability Report",
            "2024-04-25",
            url="https://example.com/china-mobile-sustainability.pdf",
            code="00941",
        ),
    ]

    selected = select_hk_main(rows, ticker, Period(year=2023, type=PeriodType.ANNUAL))

    assert selected["url"] == "https://example.com/china-mobile-annual.pdf"


def test_fiscal_year_lookup_returns_default_december() -> None:
    assert fiscal_year_end_month("00700") == 12


def test_fiscal_year_lookup_known_march_year_end() -> None:
    assert fiscal_year_end_month("09988") == 3
    assert fiscal_year_end_month("00823") == 3


def test_fiscal_year_lookup_known_june_year_end() -> None:
    assert fiscal_year_end_month("00016") == 6


def test_fiscal_year_lookup_known_may_year_end() -> None:
    assert fiscal_year_end_month("09901") == 5

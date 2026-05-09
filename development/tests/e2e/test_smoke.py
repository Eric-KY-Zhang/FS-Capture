from __future__ import annotations

import os
from pathlib import Path

import pytest

from app.core.models import Exchange, Period, PeriodType
from plugins import get_plugin

pytestmark = [
    pytest.mark.e2e,
    pytest.mark.skipif(
        os.environ.get("FS_CAPTURE_RUN_E2E") != "1",
        reason="set FS_CAPTURE_RUN_E2E=1 to run network smoke tests",
    ),
]


def _assert_downloaded_reports(reports) -> None:
    assert reports
    for report in reports:
        path = Path(report.local_path)
        assert path.exists()
        assert path.stat().st_size > 100_000


def test_ashare_600519_2024() -> None:
    plugin = get_plugin(Exchange.A_SHARE)
    ticker = plugin.resolve_name("600519")
    assert ticker.name
    company = plugin.fetch_company(ticker)
    assert company.currency == "CNY"
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)


def test_hk_00700_2024() -> None:
    plugin = get_plugin(Exchange.HK)
    ticker = plugin.resolve_name("00700")
    assert ticker.name
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)


def test_us_aapl_2024() -> None:
    plugin = get_plugin(Exchange.US)
    ticker = plugin.resolve_name("AAPL")
    assert ticker.name
    company = plugin.fetch_company(ticker)
    assert company.currency == "USD"
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)


def test_kr_smoke_requires_user_dart_key() -> None:
    pytest.skip("KR e2e requires a user-provided DART API key")

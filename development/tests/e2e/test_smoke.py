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


def _assert_sg_sidecars(reports) -> None:
    from app.core.sidecar import sidecar_path, write_sidecar

    for report in reports:
        path = write_sidecar(report)
        assert path == sidecar_path(report)
        assert path.parent.name == "SG"
        assert path.exists()


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


def test_kr_005930_2024() -> None:
    """Samsung Electronics 2024 annual report via DART OpenAPI.

    Requires a valid DART API key in config.toml (or via Settings dialog).
    """
    plugin = get_plugin(Exchange.KR)
    ticker = plugin.resolve_name("005930")
    assert ticker.name
    company = plugin.fetch_company(ticker)
    assert company.currency == "KRW"
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)


def test_tw_2330_2024() -> None:
    """TSMC 2024 annual report via MOPS (公開資訊觀測站 / doc.twse.com.tw)."""
    plugin = get_plugin(Exchange.TW)
    ticker = plugin.resolve_name("2330")
    assert ticker.name
    company = plugin.fetch_company(ticker)
    assert company.currency == "TWD"
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)


def test_tw_2330_2024_q2_interim() -> None:
    """TSMC 2024 半年报 (Q2 financial statements) via MOPS mtype=A."""
    plugin = get_plugin(Exchange.TW)
    ticker = plugin.resolve_name("2330")
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.Q2), out)
    assert any(report.kind == "interim_report" for report in reports)
    _assert_downloaded_reports(reports)


def test_sg_d05_2024_annual() -> None:
    plugin = get_plugin(Exchange.SG)
    ticker = plugin.resolve_name("D05")
    assert ticker.name
    company = plugin.fetch_company(ticker)
    assert company.currency == "SGD"
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)
    _assert_sg_sidecars(reports)


def test_sg_u11_2024_annual_and_h1() -> None:
    plugin = get_plugin(Exchange.SG)
    ticker = plugin.resolve_name("U11")
    assert ticker.name
    out = Path.cwd() / "e2e_output"
    annual = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    h1 = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.Q2), out)
    assert any(report.kind == "annual_report" for report in annual)
    assert any(report.kind == "interim_report" for report in h1)
    _assert_downloaded_reports(annual)
    _assert_downloaded_reports(h1)
    _assert_sg_sidecars(annual + h1)


def test_sg_z74_2024_annual() -> None:
    plugin = get_plugin(Exchange.SG)
    ticker = plugin.resolve_name("Z74")
    assert ticker.name
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)
    _assert_sg_sidecars(reports)


def test_sg_3407_2024_ipo_prospectus() -> None:
    plugin = get_plugin(Exchange.SG)
    ticker = plugin.resolve_name("3407")
    assert ticker.name == "LION-CM EM ASIA INDEX ETF"
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(
        ticker, Period(year=2024, type=PeriodType.IPO_PROSPECTUS), out
    )
    assert any(report.kind == "ipo_prospectus" for report in reports)
    _assert_downloaded_reports(reports)
    _assert_sg_sidecars(reports)


def _jp_public_annual_smoke(code: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from plugins.jp import name_resolver
    from plugins.jp import reports as jp_reports

    monkeypatch.delenv("EDINET_API_KEY", raising=False)
    monkeypatch.delenv("EDINET_SUBSCRIPTION_KEY", raising=False)
    name_resolver.reset_edinet_client()
    monkeypatch.setattr(name_resolver, "_edinet_api_key", lambda: "")
    monkeypatch.setattr(jp_reports, "_edinet_api_key", lambda: "")
    monkeypatch.setattr(name_resolver, "cached_or_load", lambda _key, loader, *, expire: loader())

    plugin = get_plugin(Exchange.JP)
    ticker = plugin.resolve_name(code)
    assert ticker.name
    company = plugin.fetch_company(ticker)
    assert company.currency == "JPY"
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(ticker, Period(year=2024, type=PeriodType.ANNUAL), out)
    assert any(report.kind == "annual_report" for report in reports)
    _assert_downloaded_reports(reports)


def test_jp_7203_2024_public_no_key(monkeypatch: pytest.MonkeyPatch) -> None:
    _jp_public_annual_smoke("7203", monkeypatch)


def test_jp_6758_2024_public_no_key(monkeypatch: pytest.MonkeyPatch) -> None:
    _jp_public_annual_smoke("6758", monkeypatch)


def test_jp_9984_2024_public_no_key(monkeypatch: pytest.MonkeyPatch) -> None:
    _jp_public_annual_smoke("9984", monkeypatch)


@pytest.mark.skipif(
    os.environ.get("FS_CAPTURE_RUN_SLOW_E2E") != "1",
    reason="set FS_CAPTURE_RUN_SLOW_E2E=1 to run the long TW IPO sweep",
)
def test_tw_2330_ipo_prospectus() -> None:
    """TSMC IPO prospectus via MOPS mtype=B sweep."""
    plugin = get_plugin(Exchange.TW)
    ticker = plugin.resolve_name("2330")
    out = Path.cwd() / "e2e_output"
    reports = plugin.download_reports(
        ticker, Period(year=1994, type=PeriodType.IPO_PROSPECTUS), out
    )
    assert any(report.kind == "ipo_prospectus" for report in reports)
    _assert_downloaded_reports(reports)

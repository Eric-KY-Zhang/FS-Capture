from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.core.models import Exchange, Period, PeriodType, Ticker
from plugins.ashare import AShare
from plugins.hk import HKShare
from plugins.jp import JPShare
from plugins.kr import KRShare
from plugins.sg import SGShare
from plugins.uk import UKShare
from plugins.us import USShare


class _Client:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


def test_ashare_resolve_name_uses_cached_map_and_orgid(monkeypatch) -> None:
    from plugins.ashare import name_resolver

    monkeypatch.setattr(name_resolver, "_load_name_map", lambda: {"600519": "贵州茅台"})
    monkeypatch.setattr(name_resolver, "_fetch_orgid", lambda _code: "gssh0600519")

    ticker = AShare().resolve_name("600519.SH")

    assert ticker.exchange is Exchange.A_SHARE
    assert ticker.code == "600519"
    assert ticker.name == "贵州茅台"
    assert ticker.external_id == "gssh0600519"


def test_ashare_download_reports_selects_annual_filing(monkeypatch) -> None:
    from plugins.ashare import reports

    row = {
        "announcementTitle": "贵州茅台2024年度报告",
        "announcementTime": 1745500000000,
        "adjunctUrl": "finalpage/2025-04-01/report.PDF",
    }
    monkeypatch.setattr(reports, "default_client", lambda **_kwargs: _Client())
    monkeypatch.setattr(reports, "_query_announcements", lambda *_args: [row])
    monkeypatch.setattr(reports, "stream_to_file", lambda *_args, **_kwargs: 100_001)

    ticker = Ticker(exchange=Exchange.A_SHARE, code="600519", name="贵州茅台", external_id="org")
    period = Period(year=2024, type=PeriodType.ANNUAL)

    result = AShare().download_reports(ticker, period, Path("out"))

    assert result[0].kind == "annual_report"
    assert result[0].title == "贵州茅台2024年度报告"


def test_hk_resolve_name_uses_prefix_stock_id(monkeypatch) -> None:
    from plugins.hk import name_resolver

    monkeypatch.setattr(name_resolver, "_cached_result", lambda _code: None)
    monkeypatch.setattr(
        name_resolver,
        "_fetch_hkex_prefix",
        lambda _code: {"code": "00700", "name": "TENCENT", "stockId": "7609"},
    )
    monkeypatch.setattr(name_resolver, "_fetch_chinese_short_name", lambda _code: "腾讯控股")
    monkeypatch.setattr(name_resolver, "_store_cache", lambda *_args: None)

    ticker = HKShare().resolve_name("700")

    assert ticker.code == "00700"
    assert ticker.name == "腾讯控股"
    assert ticker.external_id == "7609"


def test_hk_download_reports_selects_best_candidate(monkeypatch) -> None:
    from plugins.hk import reports

    rows = [
        {
            "url": "https://example.com/esg.pdf",
            "title": "ESG Report 2024",
            "doc_type": "Annual Report",
            "filing_date": None,
            "stock_codes": ("00700",),
            "source_id": "esg",
        },
        {
            "url": "https://example.com/annual.pdf",
            "title": "Annual Report 2024",
            "doc_type": "Annual Report",
            "filing_date": None,
            "stock_codes": ("00700",),
            "source_id": "annual",
            "file_size": 2_000_000,
        },
    ]
    monkeypatch.setattr(reports, "default_client", lambda **_kwargs: _Client())
    monkeypatch.setattr(reports, "_search_rows", lambda *_args, **_kwargs: rows)
    monkeypatch.setattr(reports, "verify_pdf_year_and_kind", lambda *_args: True)
    monkeypatch.setattr(reports, "stream_to_file", lambda *_args, **_kwargs: 100_001)

    ticker = Ticker(exchange=Exchange.HK, code="00700", name="TENCENT", external_id="7609")
    result = HKShare().download_reports(
        ticker, Period(year=2024, type=PeriodType.ANNUAL), Path("out")
    )

    assert result[0].source_url == "https://example.com/annual.pdf"


def test_us_resolve_name_supports_class_suffix_variant(monkeypatch) -> None:
    from plugins.us import name_resolver

    monkeypatch.setattr(
        name_resolver,
        "_load_map",
        lambda: {"BRK-B": {"cik": 1067983, "name": "Berkshire Hathaway Inc."}},
    )

    ticker = USShare().resolve_name("BRK.B")

    assert ticker.code == "BRK-B"
    assert ticker.external_id == "0001067983"


def test_us_download_reports_uses_recent_filings(monkeypatch) -> None:
    from plugins.us import reports

    payload = {
        "filings": {
            "recent": {
                "form": ["10-K"],
                "accessionNumber": ["0000320193-24-000123"],
                "primaryDocument": ["aapl-20240928.htm"],
                "filingDate": ["2024-11-01"],
                "reportDate": ["2024-09-28"],
            }
        }
    }
    monkeypatch.setattr(reports, "default_client", lambda **_kwargs: _Client())
    monkeypatch.setattr(reports, "get_json", lambda *_args, **_kwargs: payload)
    monkeypatch.setattr(
        reports,
        "_download_primary_as_pdf",
        lambda *_args, **_kwargs: ("https://sec/report", "html", 100_001),
    )

    ticker = Ticker(exchange=Exchange.US, code="AAPL", name="Apple Inc.", external_id="0000320193")
    result = USShare().download_reports(
        ticker, Period(year=2024, type=PeriodType.ANNUAL), Path("out")
    )

    assert result[0].kind == "annual_report"
    assert result[0].form == "10-K"


def test_us_html_filings_are_rendered_from_downloaded_local_file(monkeypatch) -> None:
    from plugins.us import reports

    calls = {}

    def fake_stream_to_file(client, url, dest, **kwargs):
        calls["client"] = client
        calls["url"] = url
        calls["dest"] = Path(dest)
        calls["kwargs"] = kwargs
        return 28

    def fake_render_url_to_pdf(url, dest):
        calls["render_url"] = url
        return 100_001

    monkeypatch.setattr(reports, "stream_to_file", fake_stream_to_file)
    monkeypatch.setattr(reports, "_render_url_to_pdf", fake_render_url_to_pdf)

    client = object()
    row = {
        "accessionNumber": "0000320193-24-000123",
        "primaryDocument": "aapl-20240928.htm",
    }
    dest = Path("US_AAPL_Apple Inc_2024_年报.pdf")

    source_url, source_format, n_bytes = reports._download_primary_as_pdf(
        client, "0000320193", row, dest
    )

    assert source_url.endswith("/aapl-20240928.htm")
    assert source_format == "html"
    assert n_bytes == 100_001
    assert calls["client"] is client
    assert calls["kwargs"]["source"] == "sec"
    assert calls["kwargs"]["read_timeout"] is None
    assert calls["render_url"].startswith("file:///")
    assert calls["dest"].name == "US_AAPL_Apple Inc_2024_年报.source.html"
    assert not calls["dest"].exists()


def test_kr_resolve_name_uses_corp_code_map(monkeypatch) -> None:
    from plugins.kr import name_resolver

    monkeypatch.setattr(
        name_resolver,
        "resolve_one",
        lambda _code: {"corp_code": "00126380", "corp_name": "삼성전자"},
    )

    ticker = KRShare().resolve_name("005930")

    assert ticker.code == "005930"
    assert ticker.name == "삼성전자"
    assert ticker.external_id == "00126380"


def test_kr_fetch_company_uses_dart_induty_code(monkeypatch) -> None:
    from plugins.kr import name_resolver

    class _Dart:
        def company(self, corp: str):
            assert corp == "00126380"
            return pd.DataFrame([{"corp_name": "삼성전자", "induty_code": "264"}])

    monkeypatch.setattr(name_resolver, "_dart", lambda: _Dart())

    ticker = Ticker(exchange=Exchange.KR, code="005930", name="삼성전자", external_id="00126380")
    company = KRShare().fetch_company(ticker)

    assert company.industry == "264"
    assert company.extra["induty_code"] == "264"


def test_kr_download_reports_selects_q3_by_month(monkeypatch) -> None:
    from plugins.kr import reports

    df = pd.DataFrame(
        [
            {"rcept_no": "q1", "report_nm": "분기보고서 (2024.03)", "rcept_dt": "20240515"},
            {"rcept_no": "q3", "report_nm": "분기보고서 (2024.09)", "rcept_dt": "20241114"},
        ]
    )
    monkeypatch.setattr(reports, "_list_filings", lambda *_args: df)
    monkeypatch.setattr(
        reports,
        "_download_rcept_as_pdf",
        lambda *_args: ("https://dart/report.pdf", "pdf", 100_001),
    )

    ticker = Ticker(exchange=Exchange.KR, code="005930", name="삼성전자", external_id="00126380")
    result = KRShare().download_reports(ticker, Period(year=2024, type=PeriodType.Q3), Path("out"))

    assert result[0].accession_number == "q3"
    assert result[0].kind == "q3_report"


def test_jp_plugin_is_registered() -> None:
    from plugins import get_plugin

    assert isinstance(get_plugin(Exchange.JP), JPShare)


def test_uk_plugin_is_registered() -> None:
    from plugins import get_plugin

    assert isinstance(get_plugin(Exchange.UK), UKShare)


def test_sg_plugin_is_registered() -> None:
    from plugins import get_plugin

    assert isinstance(get_plugin(Exchange.SG), SGShare)

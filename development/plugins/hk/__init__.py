from __future__ import annotations

from pathlib import Path

from app.core.models import (
    Company,
    Exchange,
    FinancialStatement,
    Period,
    ReportFile,
    Ticker,
)
from plugins.base import ExchangePlugin


class HKShare(ExchangePlugin):
    exchange = Exchange.HK

    def resolve_name(self, code: str) -> Ticker:
        from .name_resolver import resolve as _resolve
        return _resolve(code)

    def fetch_company(self, ticker: Ticker) -> Company:
        from .name_resolver import fetch_company as _fc
        return _fc(ticker)

    def download_reports(
        self, ticker: Ticker, period: Period, output_root: Path
    ) -> list[ReportFile]:
        from .reports import download as _download
        return _download(ticker, period, output_root)

    def fetch_financials(self, ticker: Ticker, period: Period) -> list[FinancialStatement]:
        from .financials import fetch as _fetch
        return _fetch(ticker, period)

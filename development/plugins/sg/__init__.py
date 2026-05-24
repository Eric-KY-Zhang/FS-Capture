from __future__ import annotations

from pathlib import Path

from app.core.models import Company, Exchange, Period, ReportFile, Ticker
from plugins.base import ExchangePlugin


class SGShare(ExchangePlugin):
    exchange = Exchange.SG

    def resolve_name(self, code: str) -> Ticker:
        from .name_resolver import resolve as _resolve

        return _resolve(code)

    def fetch_company(self, ticker: Ticker) -> Company:
        from .name_resolver import fetch_company as _fetch_company

        return _fetch_company(ticker)

    def download_reports(
        self, ticker: Ticker, period: Period, output_root: Path
    ) -> list[ReportFile]:
        from .reports import download as _download

        return _download(ticker, period, output_root)

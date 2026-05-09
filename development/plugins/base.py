from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from app.core.models import Company, Exchange, Period, ReportFile, Ticker


class ExchangePlugin(ABC):
    """Per-market data adapter. Subclass once per exchange.

    The orchestrator calls these in sequence per (ticker, period) task:
      1. resolve_name(code)        -> Ticker (with name + external_id)
      2. fetch_company(ticker)     -> Company metadata (industry, currency, ...)
      3. download_reports(t, p)    -> ReportFile[] streamed to disk
    """

    exchange: Exchange

    @abstractmethod
    def resolve_name(self, code: str) -> Ticker:
        """Convert a user-entered code into a Ticker with name + external_id.
        Must raise ValueError if the code cannot be resolved.
        """

    @abstractmethod
    def fetch_company(self, ticker: Ticker) -> Company:
        """Fetch listing date, industry, currency, etc."""

    @abstractmethod
    def download_reports(
        self,
        ticker: Ticker,
        period: Period,
        output_root: Path,
    ) -> list[ReportFile]:
        """Download annual / audit / quarterly reports for the given period.

        Files saved flat under output_root, with market/code/year/type in the filename.
        Returns a list of ReportFile records pointing to local paths.
        Empty list = no filings found for this period (not an error).
        """

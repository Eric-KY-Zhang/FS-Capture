from __future__ import annotations

from pathlib import Path

from app.core.models import Period, Ticker
from app.core.sidecar import iter_sidecars


def already_downloaded(
    ticker: Ticker,
    period: Period,
    cache_root: Path | None = None,
) -> bool:
    """True when sidecar cache contains this exchange/code/year/period tuple."""
    for payload in iter_sidecars(ticker.exchange, cache_root):
        if (
            payload.get("ticker_code") == ticker.code
            and payload.get("period_year") == period.year
            and payload.get("period_type") == period.type.value
        ):
            return True
    return False


def compute_incremental_pairs(
    tickers: list[Ticker],
    periods: list[Period],
    cache_root: Path | None = None,
) -> tuple[list[tuple[Ticker, Period]], int]:
    todo: list[tuple[Ticker, Period]] = []
    skipped = 0
    for ticker in tickers:
        for period in periods:
            if already_downloaded(ticker, period, cache_root):
                skipped += 1
            else:
                todo.append((ticker, period))
    return todo, skipped

from __future__ import annotations

import json

from app.core.incremental import already_downloaded, compute_incremental_pairs
from app.core.job import Job
from app.core.models import Exchange, Period, PeriodType, Ticker


def _period(year: int = 2024, period_type: PeriodType = PeriodType.ANNUAL) -> Period:
    return Period(year=year, type=period_type)


def _ticker(exchange: Exchange, code: str) -> Ticker:
    return Ticker(exchange=exchange, code=code, name=f"{code} Co")


def _write_sidecar(cache_root, exchange: Exchange, code: str, period: Period) -> None:
    path = cache_root / "sidecars" / exchange.value / f"{exchange.value}_{code}.meta.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "exchange": exchange.value,
        "ticker_code": code,
        "ticker_name": f"{code} Co",
        "period_year": period.year,
        "period_type": period.type.value,
        "kind": "annual_report",
        "source_url": "https://example.com/report.pdf",
    }
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def test_compute_incremental_pairs_skips_existing(tmp_path) -> None:
    period = _period()
    tickers = [
        _ticker(Exchange.A_SHARE, "600519"),
        _ticker(Exchange.A_SHARE, "000001"),
        _ticker(Exchange.A_SHARE, "300750"),
        _ticker(Exchange.A_SHARE, "600520"),
        _ticker(Exchange.A_SHARE, "601398"),
    ]
    for ticker in tickers[:3]:
        _write_sidecar(tmp_path, ticker.exchange, ticker.code, period)

    todo, skipped = compute_incremental_pairs(tickers, [period], tmp_path)

    assert skipped == 3
    assert [(ticker.code, period.year) for ticker, period in todo] == [
        ("600520", 2024),
        ("601398", 2024),
    ]


def test_compute_incremental_pairs_when_cache_empty(tmp_path) -> None:
    tickers = [_ticker(Exchange.HK, "00700"), _ticker(Exchange.HK, "09988")]
    periods = [_period(2024, PeriodType.ANNUAL), _period(2024, PeriodType.Q2)]

    todo, skipped = compute_incremental_pairs(tickers, periods, tmp_path)

    assert skipped == 0
    assert len(todo) == 4


def test_already_downloaded_matches_exchange_specific(tmp_path) -> None:
    period = _period()
    _write_sidecar(tmp_path, Exchange.HK, "00005", period)

    assert already_downloaded(_ticker(Exchange.HK, "00005"), period, tmp_path)
    assert not already_downloaded(_ticker(Exchange.US, "00005"), period, tmp_path)


def test_job_can_use_explicit_task_pairs_without_cartesian_product() -> None:
    annual = _period(2024, PeriodType.ANNUAL)
    interim = _period(2024, PeriodType.Q2)
    aapl = _ticker(Exchange.US, "AAPL")
    msft = _ticker(Exchange.US, "MSFT")
    pairs = [(aapl, annual), (msft, interim)]

    job = Job(
        tickers=[aapl, msft],
        periods=[annual, interim],
        output_dir="output",
        task_pairs=pairs,
    )

    assert job.task_count() == 2
    assert job.pairs() == pairs

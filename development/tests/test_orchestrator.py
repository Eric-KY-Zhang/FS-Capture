from __future__ import annotations

import threading
from pathlib import Path

import pytest

from app.core.cache import close_cache
from app.core.job import TaskResult, TaskStatus
from app.core.models import Company, Exchange, Period, PeriodType, Ticker
from app.core.orchestrator import OrchestratorSignals, _TaskRunnable


class _FakePlugin:
    exchange = Exchange.A_SHARE

    def __init__(self) -> None:
        self.download_called = False

    def resolve_name(self, code: str) -> Ticker:
        return Ticker(exchange=Exchange.A_SHARE, code=code, name="测试公司", external_id="org")

    def fetch_company(self, ticker: Ticker) -> Company:
        return Company(ticker=ticker)

    def download_reports(self, ticker: Ticker, period: Period, output_root: Path):
        self.download_called = True
        return []


def test_cancel_event_stops_task(monkeypatch: pytest.MonkeyPatch) -> None:
    import plugins

    plugin = _FakePlugin()
    monkeypatch.setattr(plugins, "get_plugin", lambda _exchange: plugin)
    cancel_event = threading.Event()
    cancel_event.set()
    result = TaskResult(
        ticker=Ticker(
            exchange=Exchange.A_SHARE,
            code="600519",
            name="贵州茅台",
            external_id="org",
        ),
        period=Period(year=2024, type=PeriodType.ANNUAL),
    )
    runnable = _TaskRunnable(
        result,
        OrchestratorSignals(),
        Path("output"),
        cancel_event,
    )

    runnable.run()

    assert result.status is TaskStatus.CANCELLED
    assert not plugin.download_called


def test_cache_close_idempotent() -> None:
    close_cache()
    close_cache()

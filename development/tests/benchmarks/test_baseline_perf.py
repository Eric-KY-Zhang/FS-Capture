"""v0.9 baseline performance benchmark.

Run manually:
    $env:FS_CAPTURE_BENCHMARK_RUNS="3"
    pytest -m benchmark --benchmark-only -s

The tests are marked as both benchmark and e2e, so normal
`pytest -m "not e2e"` runs do not execute network benchmarks.
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from statistics import median, pstdev

import httpx
import pytest

from app.core.cache import get_cache
from app.core.models import Exchange, Period, PeriodType
from app.core.settings import Settings, load_settings
from plugins import get_plugin

pytestmark = [
    pytest.mark.benchmark,
    pytest.mark.e2e,
    pytest.mark.skipif(
        os.environ.get("FS_CAPTURE_RUN_BENCHMARK") != "1",
        reason="set FS_CAPTURE_RUN_BENCHMARK=1 to run manual performance benchmarks",
    ),
]

RUNS = max(1, int(os.environ.get("FS_CAPTURE_BENCHMARK_RUNS", "3")))
ANNUAL_2024 = Period(year=2024, type=PeriodType.ANNUAL)

SINGLE_SCENARIOS = [
    (Exchange.A_SHARE, "600519"),
    (Exchange.HK, "00700"),
    (Exchange.US, "AAPL"),
    (Exchange.KR, "005930"),
    (Exchange.TW, "2330"),
    (Exchange.JP, "7203"),
    (Exchange.UK, "ULVR"),
    (Exchange.SG, "D05"),
]

BATCH_SCENARIOS = [
    (Exchange.A_SHARE, "600519"),
    (Exchange.A_SHARE, "000001"),
    (Exchange.A_SHARE, "300750"),
    (Exchange.HK, "00700"),
    (Exchange.HK, "09988"),
    (Exchange.HK, "00005"),
    (Exchange.US, "AAPL"),
    (Exchange.US, "MSFT"),
    (Exchange.US, "TSLA"),
    (Exchange.KR, "005930"),
    (Exchange.KR, "000660"),
    (Exchange.KR, "035420"),
    (Exchange.TW, "2330"),
    (Exchange.TW, "2317"),
    (Exchange.TW, "2454"),
    (Exchange.JP, "7203"),
    (Exchange.JP, "6758"),
    (Exchange.JP, "9984"),
    (Exchange.UK, "ULVR"),
    (Exchange.UK, "HSBA"),
    (Exchange.UK, "AZN"),
    (Exchange.SG, "D05"),
    (Exchange.SG, "U11"),
    (Exchange.SG, "Z74"),
]


def _enabled_scenarios(scenarios):
    raw = os.environ.get("FS_CAPTURE_BENCHMARK_MARKETS", "").strip()
    has_edinet_key = bool(load_settings().edinet.api_key)
    if not raw:
        return [
            (exchange, code)
            for exchange, code in scenarios
            if exchange is not Exchange.JP or has_edinet_key
        ]
    allowed = {item.strip().upper() for item in raw.split(",") if item.strip()}
    return [
        (exchange, code)
        for exchange, code in scenarios
        if exchange.value in allowed and (exchange is not Exchange.JP or has_edinet_key)
    ]


def _duration_stats(durations: list[float]) -> tuple[float, float]:
    return median(durations), pstdev(durations) if len(durations) > 1 else 0.0


def _download_once(exchange: Exchange, code: str, output_root: Path) -> tuple[int, int]:
    plugin = get_plugin(exchange)
    ticker = plugin.resolve_name(code)
    reports = plugin.download_reports(ticker, ANNUAL_2024, output_root)
    total_bytes = sum(Path(report.local_path).stat().st_size for report in reports)
    return len(reports), total_bytes


def _print_result(name: str, durations: list[float], extra: str = "") -> None:
    med, std = _duration_stats(durations)
    suffix = f" | {extra}" if extra else ""
    print(f"[BENCH] {name}: median={med:.2f}s std={std:.2f}s runs={len(durations)}{suffix}")


def test_source_ui_bootstrap_to_main_view() -> None:
    script = """
import os
import time
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
start = time.perf_counter()
from PySide6.QtWidgets import QApplication
from app.core.settings import Settings
from app.ui.i18n import LanguageManager
from app.ui.main_view import MainView
app = QApplication([])
LanguageManager.instance().set_language("zh")
view = MainView(Settings())
view.show()
app.processEvents()
print(f"{time.perf_counter() - start:.6f}")
"""
    env = os.environ.copy()
    env.setdefault("QT_QPA_PLATFORM", "offscreen")
    durations = []
    for _ in range(RUNS):
        proc = subprocess.run(
            [sys.executable, "-c", script],
            cwd=Path.cwd(),
            env=env,
            check=True,
            capture_output=True,
            text=True,
            timeout=60,
        )
        durations.append(float(proc.stdout.strip().splitlines()[-1]))
    _print_result("source-ui-bootstrap", durations, "QApplication + MainView offscreen")
    assert median(durations) > 0


@pytest.mark.parametrize(("exchange", "ticker_code"), _enabled_scenarios(SINGLE_SCENARIOS))
def test_single_ticker_2024_annual(tmp_path: Path, exchange: Exchange, ticker_code: str) -> None:
    durations = []
    report_counts = []
    byte_counts = []
    for run_index in range(RUNS):
        output_root = tmp_path / exchange.value / str(run_index)
        start = time.monotonic()
        report_count, total_bytes = _download_once(exchange, ticker_code, output_root)
        durations.append(time.monotonic() - start)
        report_counts.append(report_count)
        byte_counts.append(total_bytes)

    _print_result(
        f"{exchange.value} {ticker_code} annual-2024",
        durations,
        f"reports={report_counts} bytes={byte_counts}",
    )
    assert all(count > 0 for count in report_counts)
    assert all(size > 100_000 for size in byte_counts)


def test_batch_24_annual_downloads(tmp_path: Path) -> None:
    scenarios = _enabled_scenarios(BATCH_SCENARIOS)
    max_workers = Settings().concurrency.max_workers
    start = time.monotonic()
    results = []
    failures = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {
            executor.submit(
                _download_once,
                exchange,
                code,
                tmp_path / "batch" / exchange.value / code,
            ): (exchange, code)
            for exchange, code in scenarios
        }
        for future in as_completed(future_map):
            exchange, code = future_map[future]
            try:
                report_count, total_bytes = future.result()
                results.append((exchange.value, code, report_count, total_bytes))
            except Exception as exc:  # noqa: BLE001
                failures.append((exchange.value, code, f"{type(exc).__name__}: {exc}"))

    elapsed = time.monotonic() - start
    total_reports = sum(item[2] for item in results)
    total_bytes = sum(item[3] for item in results)
    zero_report_tasks = [(exchange, code) for exchange, code, count, _bytes in results if count == 0]
    print(
        "[BENCH] batch-annual-2024: "
        f"tasks={len(scenarios)} workers={max_workers} elapsed={elapsed:.2f}s "
        f"reports={total_reports} bytes={total_bytes} "
        f"zero_reports={zero_report_tasks} failures={failures}"
    )
    assert not failures
    assert len(results) == len(scenarios)
    assert total_reports > 0


def test_name_resolver_cold_warm_cache() -> None:
    from plugins.sg import sgxnet_web

    get_cache().delete("sg:ticker:D05:v1")
    sgxnet_web.reset_sgxnet_cache()
    plugin = get_plugin(Exchange.SG)

    start = time.monotonic()
    cold = plugin.resolve_name("D05")
    cold_duration = time.monotonic() - start

    start = time.monotonic()
    warm = plugin.resolve_name("D05")
    warm_duration = time.monotonic() - start

    print(
        "[BENCH] name-resolver SG D05: "
        f"cold={cold_duration * 1000:.1f}ms warm={warm_duration * 1000:.1f}ms "
        f"speedup={cold_duration / max(warm_duration, 0.000001):.1f}x"
    )
    assert cold.code == warm.code == "D05"
    assert warm_duration <= cold_duration


def test_dart_http2_probe() -> None:
    url = "https://opendart.fss.or.kr/"
    results = []
    for http2 in (False, True):
        durations = []
        versions = []
        for _ in range(RUNS):
            start = time.monotonic()
            with httpx.Client(http2=http2, timeout=20.0, follow_redirects=True) as client:
                resp = client.get(url)
                resp.raise_for_status()
                versions.append(resp.http_version)
            durations.append(time.monotonic() - start)
        med, std = _duration_stats(durations)
        results.append((http2, med, std, sorted(set(versions))))
    print(f"[BENCH] dart-http-probe: {results}")
    assert len(results) == 2

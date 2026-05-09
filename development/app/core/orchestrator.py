from __future__ import annotations

import traceback
import threading
from pathlib import Path

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal, Slot

from .job import Job, TaskResult, TaskStatus
from .models import Exchange, Ticker
from .settings import Settings


class OrchestratorSignals(QObject):
    job_started = Signal(int)                       # total tasks
    task_started = Signal(object)                   # TaskResult
    task_progress = Signal(object, str)             # TaskResult, status text
    task_finished = Signal(object)                  # TaskResult (mutated in place)
    job_finished = Signal(object)                   # Job
    log = Signal(str, str)                          # level, message


class _Cancelled(Exception):
    """Raised internally when the user cancels a running job."""


class _TaskRunnable(QRunnable):
    def __init__(
        self,
        result: TaskResult,
        signals: OrchestratorSignals,
        output_root: Path,
        cancel_event: threading.Event,
    ) -> None:
        super().__init__()
        self.result = result
        self.signals = signals
        self.output_root = output_root
        self.cancel_event = cancel_event

    def _check_cancel(self) -> None:
        if self.cancel_event.is_set():
            raise _Cancelled()

    @Slot()
    def run(self) -> None:
        from plugins import get_plugin

        r = self.result
        plugin = get_plugin(r.ticker.exchange)

        try:
            self._check_cancel()
            self.signals.task_started.emit(r)

            if not r.ticker.name or not r.ticker.external_id:
                self._check_cancel()
                r.status = TaskStatus.RESOLVING
                self.signals.task_progress.emit(r, "识别公司名称")
                resolved = plugin.resolve_name(r.ticker.code)
                r.ticker.name = resolved.name
                r.ticker.external_id = resolved.external_id

            self._check_cancel()
            try:
                r.company = plugin.fetch_company(r.ticker)
            except Exception as exc:  # noqa: BLE001
                self.signals.log.emit(
                    "warning",
                    f"{r.ticker.code} 公司基础资料获取失败，继续执行任务：{exc}",
                )

            self._check_cancel()
            r.status = TaskStatus.DOWNLOADING
            self.signals.task_progress.emit(r, f"下载{r.label()}")
            r.reports = plugin.download_reports(r.ticker, r.period, self.output_root)
            if r.reports:
                from app.core.sidecar import write_sidecar

                for report in r.reports:
                    try:
                        write_sidecar(report)
                    except Exception as exc:  # noqa: BLE001
                        self.signals.log.emit(
                            "warning",
                            f"{report.local_path} 元数据 sidecar 写入失败：{exc}",
                        )
            if not r.reports:
                self.signals.log.emit("warning", f"{r.ticker.code} 未找到{r.label()}文件")

            r.status = TaskStatus.DONE
        except _Cancelled:
            r.status = TaskStatus.CANCELLED
            r.error = "用户取消"
        except Exception as exc:  # noqa: BLE001
            r.status = TaskStatus.FAILED
            r.error = f"{type(exc).__name__}: {exc}"
            self.signals.log.emit("error", traceback.format_exc())
        finally:
            self.signals.task_finished.emit(r)


class Orchestrator(QObject):
    """Coordinates ticker name resolution and report downloads."""

    def __init__(self, settings: Settings, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.settings = settings
        self.signals = OrchestratorSignals()
        self.pool = QThreadPool.globalInstance()
        self.pool.setMaxThreadCount(settings.concurrency.max_workers)
        self._cancel_event = threading.Event()

    # ---- Synchronous name resolution (called from UI for confirm step) ----

    def resolve_name(self, exchange: Exchange, code: str) -> Ticker:
        from plugins import get_plugin
        return get_plugin(exchange).resolve_name(code)

    def request_cancel(self) -> None:
        self._cancel_event.set()
        self.signals.log.emit("info", "用户请求取消当前任务")

    # ---- Async job execution ----

    def submit_job(self, job: Job) -> None:
        self._cancel_event.clear()
        output_root = Path(job.output_dir)
        output_root.mkdir(parents=True, exist_ok=True)

        tasks: list[TaskResult] = [
            TaskResult(ticker=t, period=p)
            for t in job.tickers
            for p in job.periods
        ]
        job.results = tasks
        if not tasks:
            self.signals.job_started.emit(0)
            self.signals.job_finished.emit(job)
            return

        task_ids = {id(r) for r in tasks}

        self.signals.job_started.emit(len(tasks))

        state = {"n": 0, "total": len(tasks)}

        def _on_finished(r: TaskResult) -> None:
            # Only react to tasks belonging to *this* job
            if id(r) not in task_ids:
                return
            state["n"] += 1
            if state["n"] >= state["total"]:
                # Disconnect this slot before emitting to avoid double-firing on subsequent jobs
                try:
                    self.signals.task_finished.disconnect(_on_finished)
                except (RuntimeError, TypeError):
                    pass
                self.signals.job_finished.emit(job)

        self.signals.task_finished.connect(_on_finished)

        for r in tasks:
            self.pool.start(_TaskRunnable(r, self.signals, output_root, self._cancel_event))

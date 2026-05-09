from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from .models import Company, Period, ReportFile, Ticker


class JobMode(str, Enum):
    REPORT_DOWNLOAD = "report_download"


class TaskStatus(str, Enum):
    PENDING = "pending"
    RESOLVING = "resolving"
    DOWNLOADING = "downloading"
    DONE = "done"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass
class TaskResult:
    ticker: Ticker
    period: Period
    status: TaskStatus = TaskStatus.PENDING
    company: Company | None = None
    reports: list[ReportFile] = field(default_factory=list)
    error: str | None = None

    def label(self) -> str:
        return self.period.label()


@dataclass
class Job:
    tickers: list[Ticker]
    periods: list[Period]
    output_dir: str
    mode: JobMode = JobMode.REPORT_DOWNLOAD
    # Job-level results aggregated as workers finish.
    results: list[TaskResult] = field(default_factory=list)

    def task_count(self) -> int:
        return len(self.tickers) * len(self.periods)

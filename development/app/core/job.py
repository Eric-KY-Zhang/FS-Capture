from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional

from .models import Company, FinancialStatement, Period, ReportFile, Ticker


class TaskStatus(str, Enum):
    PENDING = "pending"
    RESOLVING = "resolving"
    DOWNLOADING = "downloading"
    SCRAPING = "scraping"
    DONE = "done"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass
class TaskResult:
    ticker: Ticker
    period: Period
    status: TaskStatus = TaskStatus.PENDING
    company: Optional[Company] = None
    reports: list[ReportFile] = field(default_factory=list)
    statements: list[FinancialStatement] = field(default_factory=list)
    error: Optional[str] = None


@dataclass
class Job:
    tickers: list[Ticker]
    periods: list[Period]
    output_dir: str
    # Job-level results aggregated as workers finish.
    results: list[TaskResult] = field(default_factory=list)

    def task_count(self) -> int:
        return len(self.tickers) * len(self.periods)

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.core.job import TaskResult, TaskStatus
from app.core.models import Exchange, Period, PeriodType
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager
from app.ui.styles.palette import exchange_accent


class _TaskRow(QFrame):
    def __init__(self, task: TaskResult, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.task = task
        self._status_key: str | None = "PD_WAITING"
        self._status_text = ui_strings.PD_WAITING
        self._detail_key: str | None = None
        self._detail_count: int | None = None
        self._detail_text = ""
        self.setObjectName("ProgressItemRow")
        h = QHBoxLayout(self)
        h.setContentsMargins(12, 8, 12, 8)
        h.setSpacing(12)

        # Color stripe
        accent = QLabel()
        accent.setFixedSize(3, 18)
        accent.setStyleSheet(
            f"background:{exchange_accent(task.ticker.exchange)}; border-radius:2px;"
        )
        h.addWidget(accent)

        self.code_label = QLabel(_ticker_label(task))
        self.code_label.setStyleSheet("font-weight: 600;")
        self.code_label.setMinimumWidth(120)

        self.name_label = QLabel(task.ticker.name or "")
        self.name_label.setStyleSheet("color: #475569;")
        self.name_label.setMinimumWidth(130)

        self.period_label = QLabel(_period_label(task.period))
        self.period_label.setStyleSheet("color: #64748B;")
        self.period_label.setMinimumWidth(80)

        self.status_label = QLabel(ui_strings.PD_WAITING)
        self.status_label.setObjectName("StatusPill")
        self.status_label.setProperty("state", "pending")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setMinimumWidth(90)

        self.detail_label = QLabel("")
        self.detail_label.setStyleSheet("color: #94A3B8; font-size: 12px;")
        self.detail_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        h.addWidget(self.code_label)
        h.addWidget(self.name_label)
        h.addWidget(self.period_label)
        h.addWidget(self.status_label)
        h.addWidget(self.detail_label, 1)

    def update_status(
        self,
        text: str,
        state: str,
        detail: str = "",
        *,
        status_key: str | None = None,
        detail_key: str | None = None,
        detail_count: int | None = None,
    ) -> None:
        self._status_key = status_key
        self._status_text = text
        self._detail_key = detail_key
        self._detail_count = detail_count
        self._detail_text = detail
        self.status_label.setText(text)
        self.status_label.setProperty("state", state)
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)
        self.detail_label.setText(detail)
        if self.task.ticker.name:
            self.name_label.setText(self.task.ticker.name)

    def retranslate(self) -> None:
        self.code_label.setText(_ticker_label(self.task))
        self.period_label.setText(_period_label(self.task.period))
        if self._status_key is not None:
            self.status_label.setText(getattr(ui_strings, self._status_key))
        else:
            self.status_label.setText(self._status_text)
        if self._detail_key == "PD_FILE_COUNT_FORMAT" and self._detail_count is not None:
            self.detail_label.setText(
                ui_strings.PD_FILE_COUNT_FORMAT.format(count=self._detail_count)
            )
        elif self._detail_key == "PD_NO_FILE":
            self.detail_label.setText(ui_strings.PD_NO_FILE)
        else:
            self.detail_label.setText(self._detail_text)


class ProgressDock(QFrame):
    """Slides in / displays per-task progress while a job runs."""

    cancel_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("ProgressDock")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.setVisible(False)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QWidget()
        header.setObjectName("CardHeader")
        h = QHBoxLayout(header)
        h.setContentsMargins(20, 16, 20, 8)
        h.setSpacing(12)

        self.title_label = QLabel(ui_strings.PD_TITLE)
        self.title_label.setObjectName("CardTitle")
        h.addWidget(self.title_label)
        self.summary_label = QLabel("")
        self.summary_label.setStyleSheet("color: #475569;")
        h.addWidget(self.summary_label)
        h.addStretch(1)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(220)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        h.addWidget(self.progress_bar)
        self.cancel_btn = QPushButton(ui_strings.COMMON_CANCEL)
        self.cancel_btn.setCursor(Qt.PointingHandCursor)
        self.cancel_btn.clicked.connect(self.cancel_requested.emit)
        h.addWidget(self.cancel_btn)
        outer.addWidget(header)

        body = QWidget()
        body.setObjectName("CardBody")
        b = QVBoxLayout(body)
        b.setContentsMargins(20, 0, 20, 16)
        b.setSpacing(8)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._list_host = QWidget()
        self._list_layout = QVBoxLayout(self._list_host)
        self._list_layout.setContentsMargins(0, 0, 0, 0)
        self._list_layout.setSpacing(6)
        self._list_layout.addStretch(1)
        self._scroll.setWidget(self._list_host)

        b.addWidget(self._scroll, 1)
        outer.addWidget(body, 1)

        self._rows: dict[int, _TaskRow] = {}
        self._completed = 0
        self._total = 0
        LanguageManager.instance().language_changed.connect(self._retranslate)

    # ---- API ---------------------------------------------------------------

    def reset(self, total: int) -> None:
        for r in list(self._rows.values()):
            self._list_layout.removeWidget(r)
            r.deleteLater()
        self._rows.clear()
        self._total = total
        self._completed = 0
        self.progress_bar.setRange(0, max(1, total))
        self.progress_bar.setValue(0)
        self.summary_label.setText(f"0 / {total}")
        self.cancel_btn.setEnabled(total > 0)
        self.setVisible(True)

    def on_task_started(self, task: TaskResult) -> None:
        if id(task) in self._rows:
            return
        row = _TaskRow(task)
        self._list_layout.insertWidget(self._list_layout.count() - 1, row)
        self._rows[id(task)] = row
        row.update_status(ui_strings.PD_STARTED, "pending", "", status_key="PD_STARTED")

    def on_task_progress(self, task: TaskResult, status_text: str) -> None:
        row = self._rows.get(id(task))
        if not row:
            self.on_task_started(task)
            row = self._rows.get(id(task))
        if row:
            row.update_status(
                ui_strings.PD_RUNNING,
                "pending",
                status_text,
                status_key="PD_RUNNING",
            )

    def on_task_finished(self, task: TaskResult) -> None:
        row = self._rows.get(id(task))
        if not row:
            self.on_task_started(task)
            row = self._rows.get(id(task))
        if row:
            if task.status is TaskStatus.DONE:
                if task.reports:
                    report_count = len(task.reports)
                    row.update_status(
                        ui_strings.PD_DONE,
                        "ok",
                        ui_strings.PD_FILE_COUNT_FORMAT.format(count=report_count),
                        status_key="PD_DONE",
                        detail_key="PD_FILE_COUNT_FORMAT",
                        detail_count=report_count,
                    )
                else:
                    row.update_status(
                        ui_strings.PD_DONE,
                        "ok",
                        ui_strings.PD_NO_FILE,
                        status_key="PD_DONE",
                        detail_key="PD_NO_FILE",
                    )
            elif task.status is TaskStatus.FAILED:
                row.update_status(
                    ui_strings.PD_FAILED,
                    "error",
                    task.error or "",
                    status_key="PD_FAILED",
                )
            elif task.status is TaskStatus.CANCELLED:
                row.update_status(
                    ui_strings.PD_CANCELLED,
                    "pending",
                    task.error or "",
                    status_key="PD_CANCELLED",
                )
            else:
                row.update_status(task.status.value, "pending", task.error or "")

        self._completed += 1
        self.progress_bar.setValue(self._completed)
        self.summary_label.setText(f"{self._completed} / {self._total}")
        if self._completed >= self._total:
            self.cancel_btn.setEnabled(False)

    def _retranslate(self, _lang: str = "") -> None:
        self.title_label.setText(ui_strings.PD_TITLE)
        self.cancel_btn.setText(ui_strings.COMMON_CANCEL)
        for row in self._rows.values():
            row.retranslate()


def _ticker_label(task: TaskResult) -> str:
    return f"{_exchange_name_for(task.ticker.exchange)} · {task.ticker.code}"


def _exchange_name_for(exchange: Exchange) -> str:
    return {
        Exchange.A_SHARE: ui_strings.ES_NAME_A_SHARE,
        Exchange.HK: ui_strings.ES_NAME_HK,
        Exchange.US: ui_strings.ES_NAME_US,
        Exchange.KR: ui_strings.ES_NAME_KR,
        Exchange.TW: ui_strings.ES_NAME_TW,
        Exchange.JP: ui_strings.ES_NAME_JP,
        Exchange.UK: ui_strings.ES_NAME_UK,
    }[exchange]


def _period_label(period: Period) -> str:
    return ui_strings.PD_PERIOD_LABEL_FORMAT.format(
        year=period.year,
        period_type=_period_type_name(period.type),
    )


def _period_type_name(period_type: PeriodType) -> str:
    return {
        PeriodType.ANNUAL: ui_strings.PS_ANNUAL,
        PeriodType.Q1: ui_strings.PS_Q1,
        PeriodType.Q2: ui_strings.PS_Q2,
        PeriodType.Q3: ui_strings.PS_Q3,
        PeriodType.IPO_PROSPECTUS: ui_strings.PS_IPO,
    }[period_type]

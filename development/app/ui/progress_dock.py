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
from app.ui import strings as ui_strings
from app.ui.styles.palette import exchange_accent


class _TaskRow(QFrame):
    def __init__(self, task: TaskResult, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.task = task
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

        self.code_label = QLabel(f"{task.ticker.exchange.display_name} · {task.ticker.code}")
        self.code_label.setStyleSheet("font-weight: 600;")
        self.code_label.setMinimumWidth(120)

        self.name_label = QLabel(task.ticker.name or "")
        self.name_label.setStyleSheet("color: #475569;")
        self.name_label.setMinimumWidth(130)

        self.period_label = QLabel(task.label())
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

    def update_status(self, text: str, state: str, detail: str = "") -> None:
        self.status_label.setText(text)
        self.status_label.setProperty("state", state)
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)
        self.detail_label.setText(detail)
        if self.task.ticker.name:
            self.name_label.setText(self.task.ticker.name)


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

        title = QLabel(ui_strings.PD_TITLE)
        title.setObjectName("CardTitle")
        h.addWidget(title)
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
        row.update_status(ui_strings.PD_STARTED, "pending", "")

    def on_task_progress(self, task: TaskResult, status_text: str) -> None:
        row = self._rows.get(id(task))
        if not row:
            self.on_task_started(task)
            row = self._rows.get(id(task))
        if row:
            row.update_status(ui_strings.PD_RUNNING, "pending", status_text)

    def on_task_finished(self, task: TaskResult) -> None:
        row = self._rows.get(id(task))
        if not row:
            self.on_task_started(task)
            row = self._rows.get(id(task))
        if row:
            if task.status is TaskStatus.DONE:
                details: list[str] = []
                if task.reports:
                    details.append(
                        ui_strings.PD_FILE_COUNT_FORMAT.format(count=len(task.reports))
                    )
                detail = " · ".join(details) if details else ui_strings.PD_NO_FILE
                row.update_status(ui_strings.PD_DONE, "ok", detail)
            elif task.status is TaskStatus.FAILED:
                row.update_status(ui_strings.PD_FAILED, "error", task.error or "")
            elif task.status is TaskStatus.CANCELLED:
                row.update_status(ui_strings.PD_CANCELLED, "pending", task.error or "")
            else:
                row.update_status(task.status.value, "pending", task.error or "")

        self._completed += 1
        self.progress_bar.setValue(self._completed)
        self.summary_label.setText(f"{self._completed} / {self._total}")
        if self._completed >= self._total:
            self.cancel_btn.setEnabled(False)

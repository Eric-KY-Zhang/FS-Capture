from __future__ import annotations

from typing import Optional

from loguru import logger
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
    QWidget,
)

from app.core.job import Job, TaskResult
from app.core.models import Exchange, Period, Ticker
from app.core.orchestrator import Orchestrator
from app.core.settings import Settings
from app.ui.exchange_panel import ExchangePanel
from app.ui.exchange_selector import ExchangeSelector
from app.ui.output_card import OutputCard
from app.ui.period_selector import PeriodSelector
from app.ui.progress_dock import ProgressDock
from app.ui.settings_dialog import SettingsDialog


class MainView(QWidget):
    """The main content area inside the window. Composites all sections."""

    def __init__(self, settings: Settings, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.settings = settings
        self.orchestrator = Orchestrator(settings, parent=self)

        # Wire orchestrator signals
        self.orchestrator.signals.job_started.connect(self._on_job_started)
        self.orchestrator.signals.task_started.connect(self._on_task_started)
        self.orchestrator.signals.task_progress.connect(self._on_task_progress)
        self.orchestrator.signals.task_finished.connect(self._on_task_finished)
        self.orchestrator.signals.job_finished.connect(self._on_job_finished)
        self.orchestrator.signals.log.connect(self._on_log)

        # Outer scroll
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._scroll = QScrollArea()
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(self._scroll, 1)

        content = QWidget()
        content.setObjectName("MainContent")
        self._scroll.setWidget(content)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(36, 24, 36, 36)
        layout.setSpacing(20)

        # ---- Hero / heading ---------------------------------------------------
        section = QLabel("FS CAPTURE · 一键抓取")
        section.setObjectName("SectionLabel")
        layout.addWidget(section)

        title = QLabel("上市公司官方披露 PDF 下载")
        title.setStyleSheet("font-size: 22px; font-weight: 700; color: #0F172A;")
        layout.addWidget(title)

        sub = QLabel("勾选交易所 → 录入股票代码 → 选择期间 → 下载披露文件 PDF")
        sub.setStyleSheet("color: #64748B; font-size: 13px;")
        layout.addWidget(sub)
        layout.addSpacing(4)

        # ---- Exchange selector ------------------------------------------------
        self.exchange_selector = ExchangeSelector()
        self.exchange_selector.selection_changed.connect(self._sync_exchange_panels)
        layout.addWidget(self.exchange_selector)

        # ---- Per-exchange panels ---------------------------------------------
        self._panels: dict[Exchange, ExchangePanel] = {}
        for ex in (Exchange.A_SHARE, Exchange.HK, Exchange.US, Exchange.KR):
            panel = ExchangePanel(ex)
            self._panels[ex] = panel
            layout.addWidget(panel)
            panel.setVisible(False)

        # ---- Period selector --------------------------------------------------
        self.period_selector = PeriodSelector()
        layout.addWidget(self.period_selector)

        # ---- Output card ------------------------------------------------------
        default_out = str(settings.output_path())
        self.output_card = OutputCard(default_out)
        layout.addWidget(self.output_card)

        # ---- Progress dock (initially hidden) ---------------------------------
        self.progress_dock = ProgressDock()
        layout.addWidget(self.progress_dock)

        # ---- Action row -------------------------------------------------------
        actions = QHBoxLayout()
        actions.setSpacing(12)

        self.settings_btn = QPushButton("⚙  设置")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.clicked.connect(self._open_settings)

        self.run_btn = QPushButton("▶  抓报告")
        self.run_btn.setProperty("variant", "primary")
        self.run_btn.setMinimumHeight(40)
        self.run_btn.setMinimumWidth(160)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(self._start_job)

        actions.addWidget(self.settings_btn)
        actions.addStretch(1)
        actions.addWidget(self.run_btn)
        layout.addLayout(actions)

        layout.addItem(QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Initial panel sync
        self._sync_exchange_panels()

    # ---- panel visibility --------------------------------------------------

    def _sync_exchange_panels(self) -> None:
        selected = set(self.exchange_selector.selected())
        for ex, panel in self._panels.items():
            visible = ex in selected
            panel.setVisible(visible)
            # Auto-add a first row when panel becomes visible and is empty
            if visible and not panel.resolved_tickers() and not panel._rows:  # type: ignore[attr-defined]
                panel.add_row()

    # ---- job actions -------------------------------------------------------

    def _start_job(self) -> None:
        # Collect resolved tickers across all visible panels
        tickers: list[Ticker] = []
        warnings: list[str] = []
        for ex in self.exchange_selector.selected():
            panel = self._panels[ex]
            if panel.has_pending():
                warnings.append(f"{ex.display_name} 中存在尚未确认的股票代码")
            tickers.extend(panel.resolved_tickers())

        if not tickers:
            QMessageBox.warning(self, "无法开始", "请先添加并确认至少一只股票")
            return
        if warnings:
            ret = QMessageBox.question(
                self, "继续？",
                "\n".join(warnings) + "\n\n仅继续抓取已确认的股票？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if ret != QMessageBox.Yes:
                return

        periods: list[Period] = self.period_selector.periods()
        if not periods:
            QMessageBox.warning(self, "无法开始", "请至少选择一种报告类型")
            return

        # Korea needs DART key
        kr_in_use = any(t.exchange == Exchange.KR for t in tickers)
        if kr_in_use and not self.settings.dart.api_key:
            ret = QMessageBox.question(
                self, "DART 密钥缺失",
                "韩股官方披露数据主要来自 DART。当前未配置 DART API 密钥，"
                "如继续使用韩股功能，请先在设置中填入密钥。\n\n是否现在打开设置？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if ret == QMessageBox.Yes:
                self._open_settings()
            return

        out_path = self.output_card.path()
        if not out_path:
            QMessageBox.warning(self, "无效路径", "请选择有效的输出文件夹")
            return

        job = Job(tickers=tickers, periods=periods, output_dir=out_path)
        self._set_action_buttons_enabled(False)
        self.progress_dock.reset(job.task_count())
        logger.info(
            f"提交报告下载任务：{len(tickers)} 只股票 × {len(periods)} 个期间 = "
            f"{job.task_count()} 个 task"
        )
        self.orchestrator.submit_job(job)

    def _set_action_buttons_enabled(self, enabled: bool) -> None:
        self.run_btn.setEnabled(enabled)
        self.settings_btn.setEnabled(enabled)

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self.settings, self)
        if dlg.exec() == dlg.DialogCode.Accepted:
            # concurrency may have changed
            self.orchestrator.pool.setMaxThreadCount(self.settings.concurrency.max_workers)

    # ---- orchestrator slots ------------------------------------------------

    def _on_job_started(self, total: int) -> None:
        logger.info(f"Job 开始，共 {total} 个 task")

    def _on_task_started(self, task: TaskResult) -> None:
        self.progress_dock.on_task_started(task)

    def _on_task_progress(self, task: TaskResult, status_text: str) -> None:
        self.progress_dock.on_task_progress(task, status_text)

    def _on_task_finished(self, task: TaskResult) -> None:
        self.progress_dock.on_task_finished(task)

    def _on_job_finished(self, job: Job) -> None:
        self._set_action_buttons_enabled(True)
        ok = sum(1 for r in job.results if r.status.value == "done")
        fail = sum(1 for r in job.results if r.status.value == "failed")
        report_count = sum(len(r.reports) for r in job.results)

        QMessageBox.information(
            self,
            "下载完成",
            f"总计 {len(job.results)} 个 task：成功 {ok}，失败 {fail}\n\n"
            f"已下载 {report_count} 个 PDF\n"
            f"输出位置：{job.output_dir}",
        )

    def _on_log(self, level: str, message: str) -> None:
        getattr(logger, level if hasattr(logger, level) else "info")(message)

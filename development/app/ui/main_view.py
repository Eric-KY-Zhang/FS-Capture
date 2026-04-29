from __future__ import annotations

from datetime import datetime
from pathlib import Path
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


def _workbook_path(job: Job) -> Path:
    first_code = job.tickers[0].code if job.tickers else "结果"
    years = sorted({p.year for p in job.periods})
    period_part = (
        f"{years[0]}-{years[-1]}" if len(years) > 1
        else str(years[0]) if years
        else "未选期间"
    )
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    base = Path(job.output_dir) / f"底稿_{first_code}_{period_part}_{ts}.xlsx"
    if not base.exists():
        return base
    for idx in range(2, 1000):
        candidate = base.with_stem(f"{base.stem}_{idx}")
        if not candidate.exists():
            return candidate
    raise FileExistsError(f"无法生成不重名的工作簿文件名：{base}")


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

        title = QLabel("上市公司年报 / 审计报告 / 财务数据")
        title.setStyleSheet("font-size: 22px; font-weight: 700; color: #0F172A;")
        layout.addWidget(title)

        sub = QLabel("勾选交易所 → 录入股票代码 → 选择期间 → 一键生成 PDF 与底稿")
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

        self.run_btn = QPushButton("▶  开始抓取")
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
            QMessageBox.warning(self, "无法开始", "请至少选择一种期间类型")
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
        self.run_btn.setEnabled(False)
        self.progress_dock.reset(job.task_count())
        logger.info(
            f"提交任务：{len(tickers)} 只股票 × {len(periods)} 个期间 = "
            f"{job.task_count()} 个 task"
        )
        self.orchestrator.submit_job(job)

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
        self.run_btn.setEnabled(True)
        ok = sum(1 for r in job.results if r.status.value == "done")
        fail = sum(1 for r in job.results if r.status.value == "failed")

        # Write workbook
        from app.exporters.excel_writer import write_workbook

        xlsx_path = _workbook_path(job)
        try:
            wr = write_workbook(job, xlsx_path)
            xlsx_msg = f"\n📊 工作簿：{wr.path}"
        except Exception as exc:
            logger.exception(f"workbook write failed: {exc}")
            xlsx_msg = f"\n⚠ 工作簿写入失败：{exc}"

        QMessageBox.information(
            self,
            "抓取完成",
            f"总计 {len(job.results)} 个 task：成功 {ok}，失败 {fail}\n\n"
            f"📁 输出位置：{job.output_dir}{xlsx_msg}",
        )

    def _on_log(self, level: str, message: str) -> None:
        getattr(logger, level if hasattr(logger, level) else "info")(message)

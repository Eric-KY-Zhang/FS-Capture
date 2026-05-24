from __future__ import annotations

from collections.abc import Callable

from loguru import logger
from PySide6.QtCore import QRunnable, Qt, QThreadPool, Slot
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QFrame,
    QGraphicsDropShadowEffect,
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
from app.ui import strings as ui_strings
from app.ui.exchange_panel import ExchangePanel
from app.ui.exchange_selector import MARKET_ORDER, ExchangeSelector
from app.ui.i18n import LanguageManager
from app.ui.output_card import OutputCard
from app.ui.period_selector import PeriodSelector
from app.ui.progress_dock import ProgressDock
from app.ui.settings_dialog import SettingsDialog


def _load_ashare_name_map() -> object:
    from plugins.ashare import name_resolver

    return name_resolver._load_name_map()  # noqa: SLF001


def _load_tw_name_map() -> object:
    from plugins.tw import name_resolver

    return name_resolver._load_map()  # noqa: SLF001


def _load_us_name_map() -> object:
    from plugins.us import name_resolver

    return name_resolver._load_map()  # noqa: SLF001


_PREWARM_LOADERS: dict[Exchange, Callable[[], object]] = {
    Exchange.A_SHARE: _load_ashare_name_map,
    Exchange.TW: _load_tw_name_map,
    Exchange.US: _load_us_name_map,
}


def _prewarm_name_resolver(exchange: Exchange) -> None:
    loader = _PREWARM_LOADERS.get(exchange)
    if loader is None:
        return
    try:
        loader()
    except Exception:  # noqa: BLE001
        pass


class _NameResolverPrewarmTask(QRunnable):
    def __init__(self, exchange: Exchange) -> None:
        super().__init__()
        self.exchange = exchange

    @Slot()
    def run(self) -> None:
        _prewarm_name_resolver(self.exchange)


class MainView(QWidget):
    """The main content area inside the window. Composites all sections."""

    def __init__(self, settings: Settings, parent: QWidget | None = None) -> None:
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
        layout.setContentsMargins(36, 22, 36, 24)
        layout.setSpacing(18)

        # ---- Hero / heading ---------------------------------------------------
        self.title_label = QLabel(ui_strings.MV_TITLE)
        self.title_label.setObjectName("HeroTitle")
        layout.addWidget(self.title_label)

        self.subtitle_label = QLabel(ui_strings.MV_SUBTITLE)
        self.subtitle_label.setObjectName("HeroSubtitle")
        layout.addWidget(self.subtitle_label)
        layout.addSpacing(4)

        # ---- Exchange selector ------------------------------------------------
        self.section_label = QLabel(ui_strings.MV_SECTION)
        self.section_label.setObjectName("SectionLabel")
        layout.addWidget(self.section_label)

        self.exchange_selector = ExchangeSelector()
        self.exchange_selector.selection_changed.connect(self._sync_exchange_panels)
        layout.addWidget(self.exchange_selector)

        # ---- Per-exchange panels ---------------------------------------------
        self._panels: dict[Exchange, ExchangePanel] = {}
        self._prewarm_started: set[Exchange] = set()
        for ex in MARKET_ORDER:
            panel = ExchangePanel(ex)
            self._panels[ex] = panel
            layout.addWidget(panel)
            panel.setVisible(False)

        # ---- Report and output configuration ---------------------------------
        config_grid = QWidget()
        config_grid.setObjectName("ConfigGrid")
        config_layout = QHBoxLayout(config_grid)
        config_layout.setContentsMargins(0, 0, 0, 0)
        config_layout.setSpacing(12)

        self.period_selector = PeriodSelector()
        default_out = str(settings.output_path())
        self.output_card = OutputCard(default_out)
        config_layout.addWidget(self.period_selector, 3)
        config_layout.addWidget(self.output_card, 2)
        layout.addWidget(config_grid)

        # ---- Progress dock (initially hidden) ---------------------------------
        self.progress_dock = ProgressDock()
        self.progress_dock.cancel_requested.connect(self.orchestrator.request_cancel)
        layout.addWidget(self.progress_dock)
        layout.addItem(QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # ---- Fixed action bar -------------------------------------------------
        self._action_bar = QFrame()
        self._action_bar.setObjectName("ActionBar")
        actions = QHBoxLayout(self._action_bar)
        actions.setContentsMargins(36, 12, 36, 12)
        actions.setSpacing(12)

        self.settings_btn = QPushButton(ui_strings.MV_SETTINGS_BUTTON)
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.clicked.connect(self._open_settings)

        self.run_btn = QPushButton(ui_strings.MV_RUN_BUTTON)
        self.run_btn.setProperty("variant", "primary")
        self.run_btn.setMinimumHeight(40)
        self.run_btn.setMinimumWidth(160)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(lambda: self._start_job(incremental=False))
        # Primary CTA indigo glow — matches HTML design spec
        # `boxShadow: 0 6px 18px rgba(99,102,241,.28)` (board-03-workbench.jsx:315).
        # Qt QSS doesn't support box-shadow, so we use QGraphicsDropShadowEffect.
        run_btn_shadow = QGraphicsDropShadowEffect(self.run_btn)
        run_btn_shadow.setBlurRadius(18)
        run_btn_shadow.setColor(QColor(99, 102, 241, 71))  # alpha 71/255 ≈ 0.28
        run_btn_shadow.setOffset(0, 6)
        self.run_btn.setGraphicsEffect(run_btn_shadow)

        self.incremental_btn = QPushButton(ui_strings.MV_INCREMENTAL_BUTTON)
        self.incremental_btn.setMinimumHeight(40)
        self.incremental_btn.setCursor(Qt.PointingHandCursor)
        self.incremental_btn.clicked.connect(lambda: self._start_job(incremental=True))

        actions.addWidget(self.settings_btn)
        actions.addStretch(1)
        actions.addWidget(self.incremental_btn)
        actions.addWidget(self.run_btn)
        outer.addWidget(self._action_bar)

        # Initial panel sync
        self._sync_exchange_panels()
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def _retranslate(self, _lang: str = "") -> None:
        self.section_label.setText(ui_strings.MV_SECTION)
        self.title_label.setText(ui_strings.MV_TITLE)
        self.subtitle_label.setText(ui_strings.MV_SUBTITLE)
        self.settings_btn.setText(ui_strings.MV_SETTINGS_BUTTON)
        self.run_btn.setText(ui_strings.MV_RUN_BUTTON)
        self.incremental_btn.setText(ui_strings.MV_INCREMENTAL_BUTTON)

    # ---- panel visibility --------------------------------------------------

    def _sync_exchange_panels(self) -> None:
        selected = set(self.exchange_selector.selected())
        for ex, panel in self._panels.items():
            visible = ex in selected
            panel.setVisible(visible)
            # Auto-add a first row when panel becomes visible and is empty
            if visible and not panel.resolved_tickers() and not panel._rows:  # type: ignore[attr-defined]
                panel.add_row()
            if visible and ex not in self._prewarm_started:
                self._prewarm_started.add(ex)
                QThreadPool.globalInstance().start(_NameResolverPrewarmTask(ex))

    # ---- job actions -------------------------------------------------------

    def _start_job(self, incremental: bool = False) -> None:
        # Collect resolved tickers across all visible panels
        tickers: list[Ticker] = []
        warnings: list[str] = []
        for ex in self.exchange_selector.selected():
            panel = self._panels[ex]
            if panel.has_pending():
                warnings.append(
                    f"{self._exchange_name_for(ex)}{ui_strings.MV_UNCONFIRMED_SUFFIX}"
                )
            tickers.extend(panel.resolved_tickers())

        if not tickers:
            QMessageBox.warning(
                self, ui_strings.MV_CANNOT_START_TITLE, ui_strings.MV_NO_TICKERS_BODY
            )
            return
        if warnings:
            ret = QMessageBox.question(
                self,
                ui_strings.MV_CONTINUE_TITLE,
                "\n".join(warnings) + ui_strings.MV_ONLY_CONFIRMED_BODY,
                QMessageBox.Yes | QMessageBox.No,
            )
            if ret != QMessageBox.Yes:
                return

        periods: list[Period] = self.period_selector.periods()
        if not periods:
            QMessageBox.warning(
                self, ui_strings.MV_CANNOT_START_TITLE, ui_strings.MV_NO_PERIOD_BODY
            )
            return

        # Korea: DART API key is optional; without it we fall back to public crawler.
        kr_in_use = any(t.exchange == Exchange.KR for t in tickers)
        if kr_in_use and not self.settings.dart.api_key:
            logger.info(
                ui_strings.MV_KR_NO_KEY_LOG
            )
        jp_in_use = any(t.exchange == Exchange.JP for t in tickers)
        if jp_in_use and not self.settings.edinet.api_key:
            logger.info(ui_strings.MV_JP_NO_KEY_LOG)

        out_path = self.output_card.path()
        if not out_path:
            QMessageBox.warning(
                self, ui_strings.MV_INVALID_PATH_TITLE, ui_strings.MV_INVALID_PATH_BODY
            )
            return

        task_pairs: list[tuple[Ticker, Period]] | None = None
        if incremental:
            from app.core.incremental import compute_incremental_pairs

            task_pairs, skipped = compute_incremental_pairs(tickers, periods)
            if skipped:
                logger.info(ui_strings.MV_INCREMENTAL_SKIPPED_FORMAT.format(count=skipped))
            if not task_pairs:
                QMessageBox.information(
                    self,
                    ui_strings.MV_INCREMENTAL_NONE_TITLE,
                    ui_strings.MV_INCREMENTAL_NONE_BODY,
                )
                return

        job = Job(
            tickers=tickers,
            periods=periods,
            output_dir=out_path,
            task_pairs=task_pairs,
        )
        self._set_action_buttons_enabled(False)
        self.progress_dock.reset(job.task_count())
        logger.info(
            ui_strings.MV_JOB_SUBMITTED_FORMAT.format(
                tickers=len(tickers), periods=len(periods), tasks=job.task_count()
            )
        )
        self.orchestrator.submit_job(job)

    def _set_action_buttons_enabled(self, enabled: bool) -> None:
        self.run_btn.setEnabled(enabled)
        self.incremental_btn.setEnabled(enabled)
        self.settings_btn.setEnabled(enabled)

    def _open_settings(self) -> None:
        dlg = SettingsDialog(self.settings, self)
        if dlg.exec() == dlg.DialogCode.Accepted:
            # concurrency may have changed
            self.orchestrator.pool.setMaxThreadCount(self.settings.concurrency.max_workers)

    # ---- orchestrator slots ------------------------------------------------

    def _on_job_started(self, total: int) -> None:
        logger.info(ui_strings.MV_JOB_STARTED_FORMAT.format(total=total))

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
            ui_strings.MV_DONE_TITLE,
            ui_strings.MV_DONE_BODY_FORMAT.format(
                total=len(job.results),
                ok=ok,
                fail=fail,
                reports=report_count,
                output_dir=job.output_dir,
            ),
        )

    def _on_log(self, level: str, message: str) -> None:
        getattr(logger, level if hasattr(logger, level) else "info")(message)

    @staticmethod
    def _exchange_name_for(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: ui_strings.ES_NAME_A_SHARE,
            Exchange.HK: ui_strings.ES_NAME_HK,
            Exchange.US: ui_strings.ES_NAME_US,
            Exchange.KR: ui_strings.ES_NAME_KR,
            Exchange.TW: ui_strings.ES_NAME_TW,
            Exchange.JP: ui_strings.ES_NAME_JP,
            Exchange.UK: ui_strings.ES_NAME_UK,
            Exchange.SG: ui_strings.ES_NAME_SG,
        }[exchange]

from __future__ import annotations

import os
import re

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLabel

from app.core.job import TaskResult
from app.core.models import Exchange, Period, PeriodType, Ticker
from app.core.settings import Settings
from app.ui.i18n import LanguageManager
from app.ui.main_view import MainView
from app.ui.progress_dock import ProgressDock
from app.ui.settings_dialog import SettingsDialog

_CJK_RE = re.compile(r"[\u3400-\u9fff]")


def _app() -> QApplication:
    return QApplication.instance() or QApplication([])


def _main_view() -> MainView:
    _app()
    LanguageManager.instance().set_language("zh")
    return MainView(Settings())


def test_run_button_text_flips_on_language_change() -> None:
    view = _main_view()

    assert view.run_btn.text() == "▶  抓报告"

    LanguageManager.instance().set_language("en")

    assert view.run_btn.text() == "▶  Download Reports"


def test_main_view_has_no_cjk_labels_after_switching_to_en() -> None:
    view = _main_view()

    LanguageManager.instance().set_language("en")

    offenders = [
        label.text() for label in view.findChildren(QLabel) if _CJK_RE.search(label.text())
    ]
    assert offenders == []


def test_ticker_row_preserves_input_when_language_changes() -> None:
    view = _main_view()
    row = view._panels[Exchange.A_SHARE]._rows[0]  # type: ignore[attr-defined]
    row.code_input.setText("600519")

    LanguageManager.instance().set_language("en")

    assert row.code_input.text() == "600519"
    assert row.confirm_btn.text() == "Confirm"
    assert row.status_pill.text() == "Pending"


def test_settings_dialog_retranslates_own_labels_and_options() -> None:
    _app()
    LanguageManager.instance().set_language("zh")
    dialog = SettingsDialog(Settings())

    LanguageManager.instance().set_language("en")

    assert dialog.windowTitle() == "Settings"
    assert dialog.language_label.text() == "UI language"
    assert dialog.sec_ua_label.text() == "SEC User-Agent"
    assert dialog.theme.itemText(dialog.theme.findData("light")) == "Light"


def test_progress_rows_retranslate_market_period_and_status() -> None:
    _app()
    LanguageManager.instance().set_language("zh")
    dock = ProgressDock()
    task = TaskResult(
        ticker=Ticker(exchange=Exchange.A_SHARE, code="600519", name="贵州茅台"),
        period=Period(year=2024, type=PeriodType.ANNUAL),
    )
    dock.reset(1)
    dock.on_task_started(task)
    row = next(iter(dock._rows.values()))  # type: ignore[attr-defined]

    assert row.code_label.text() == "A股 · 600519"
    assert row.status_label.text() == "已开始"

    LanguageManager.instance().set_language("en")

    assert row.code_label.text() == "A-Share · 600519"
    assert row.period_label.text() == "Annual report 2024"
    assert row.status_label.text() == "Started"

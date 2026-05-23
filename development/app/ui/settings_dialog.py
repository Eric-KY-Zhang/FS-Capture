from __future__ import annotations

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLineEdit,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from app.core.settings import Settings, invalidate_dart_client_cache, save_settings
from app.ui import strings as ui_strings


class SettingsDialog(QDialog):
    def __init__(self, settings: Settings, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle(ui_strings.SD_TITLE)
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(10)

        self.dart_key = QLineEdit(settings.dart.api_key)
        self.dart_key.setPlaceholderText(ui_strings.SD_DART_PLACEHOLDER)
        self.dart_key.setEchoMode(QLineEdit.Password)
        form.addRow(ui_strings.SD_DART_LABEL, self.dart_key)

        self.workers = QSpinBox()
        self.workers.setRange(1, 16)
        self.workers.setValue(settings.concurrency.max_workers)
        form.addRow(ui_strings.SD_WORKERS_LABEL, self.workers)

        self.theme = QComboBox()
        self.theme.addItems(["light", "dark"])
        self.theme.setCurrentText(settings.ui.theme)
        form.addRow(ui_strings.SD_THEME_LABEL, self.theme)

        self.sec_ua = QLineEdit(settings.sec.user_agent)
        form.addRow("SEC User-Agent", self.sec_ua)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText(ui_strings.SD_SAVE)
        buttons.button(QDialogButtonBox.Cancel).setText(ui_strings.COMMON_CANCEL)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _save(self) -> None:
        old_dart_key = self.settings.dart.api_key
        new_dart_key = self.dart_key.text().strip()
        self.settings.dart.api_key = new_dart_key
        self.settings.concurrency.max_workers = self.workers.value()
        self.settings.ui.theme = self.theme.currentText()
        self.settings.sec.user_agent = self.sec_ua.text().strip()
        save_settings(self.settings)
        if new_dart_key != old_dart_key:
            invalidate_dart_client_cache()
        self.accept()

from __future__ import annotations

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from app.core.settings import (
    Settings,
    invalidate_dart_client_cache,
    invalidate_edinet_client_cache,
    save_settings,
)
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager


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
        self.dart_label = QLabel(ui_strings.SD_DART_LABEL)
        form.addRow(self.dart_label, self.dart_key)

        self.edinet_key = QLineEdit(settings.edinet.api_key)
        self.edinet_key.setPlaceholderText(ui_strings.SD_EDINET_PLACEHOLDER)
        self.edinet_key.setEchoMode(QLineEdit.Password)
        self.edinet_label = QLabel(ui_strings.SD_EDINET_LABEL)
        form.addRow(self.edinet_label, self.edinet_key)

        self.workers = QSpinBox()
        self.workers.setRange(1, 16)
        self.workers.setValue(settings.concurrency.max_workers)
        self.workers_label = QLabel(ui_strings.SD_WORKERS_LABEL)
        form.addRow(self.workers_label, self.workers)

        self.theme = QComboBox()
        self.theme.addItem(ui_strings.SD_THEME_LIGHT, "light")
        self.theme.addItem(ui_strings.SD_THEME_DARK, "dark")
        theme_idx = self.theme.findData(settings.ui.theme)
        self.theme.setCurrentIndex(max(0, theme_idx))
        self.theme_label = QLabel(ui_strings.SD_THEME_LABEL)
        form.addRow(self.theme_label, self.theme)

        self.language = QComboBox()
        self.language.addItem(ui_strings.SD_LANGUAGE_ZH, "zh")
        self.language.addItem(ui_strings.SD_LANGUAGE_EN, "en")
        lang_idx = self.language.findData(settings.ui.language)
        self.language.setCurrentIndex(max(0, lang_idx))
        self.language_label = QLabel(ui_strings.SD_LANGUAGE_LABEL)
        form.addRow(self.language_label, self.language)

        self.sec_ua = QLineEdit(settings.sec.user_agent)
        self.sec_ua_label = QLabel(ui_strings.SD_SEC_UA_LABEL)
        form.addRow(self.sec_ua_label, self.sec_ua)

        layout.addLayout(form)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText(ui_strings.SD_SAVE)
        self.buttons.button(QDialogButtonBox.Cancel).setText(ui_strings.COMMON_CANCEL)
        self.buttons.accepted.connect(self._save)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def _save(self) -> None:
        old_dart_key = self.settings.dart.api_key
        new_dart_key = self.dart_key.text().strip()
        old_edinet_key = self.settings.edinet.api_key
        new_edinet_key = self.edinet_key.text().strip()
        old_language = self.settings.ui.language
        new_language = self.language.currentData()
        self.settings.dart.api_key = new_dart_key
        self.settings.edinet.api_key = new_edinet_key
        self.settings.concurrency.max_workers = self.workers.value()
        self.settings.ui.theme = self.theme.currentData()
        self.settings.ui.language = new_language
        self.settings.sec.user_agent = self.sec_ua.text().strip()
        save_settings(self.settings)
        if new_dart_key != old_dart_key:
            invalidate_dart_client_cache()
        if new_edinet_key != old_edinet_key:
            invalidate_edinet_client_cache()
        if new_language != old_language:
            LanguageManager.instance().set_language(new_language)
        self.accept()

    def _retranslate(self, _lang: str = "") -> None:
        self.setWindowTitle(ui_strings.SD_TITLE)
        self.dart_key.setPlaceholderText(ui_strings.SD_DART_PLACEHOLDER)
        self.dart_label.setText(ui_strings.SD_DART_LABEL)
        self.edinet_key.setPlaceholderText(ui_strings.SD_EDINET_PLACEHOLDER)
        self.edinet_label.setText(ui_strings.SD_EDINET_LABEL)
        self.workers_label.setText(ui_strings.SD_WORKERS_LABEL)
        self.theme_label.setText(ui_strings.SD_THEME_LABEL)
        light_idx = self.theme.findData("light")
        dark_idx = self.theme.findData("dark")
        if light_idx >= 0:
            self.theme.setItemText(light_idx, ui_strings.SD_THEME_LIGHT)
        if dark_idx >= 0:
            self.theme.setItemText(dark_idx, ui_strings.SD_THEME_DARK)
        self.language_label.setText(ui_strings.SD_LANGUAGE_LABEL)
        zh_idx = self.language.findData("zh")
        en_idx = self.language.findData("en")
        if zh_idx >= 0:
            self.language.setItemText(zh_idx, ui_strings.SD_LANGUAGE_ZH)
        if en_idx >= 0:
            self.language.setItemText(en_idx, ui_strings.SD_LANGUAGE_EN)
        self.sec_ua_label.setText(ui_strings.SD_SEC_UA_LABEL)
        self.buttons.button(QDialogButtonBox.Ok).setText(ui_strings.SD_SAVE)
        self.buttons.button(QDialogButtonBox.Cancel).setText(ui_strings.COMMON_CANCEL)

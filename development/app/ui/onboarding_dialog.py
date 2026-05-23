from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager


class OnboardingDialog(QDialog):
    """First-run guide for the PDF download workflow."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.open_settings_requested = False
        self.setWindowTitle(ui_strings.OB_WINDOW_TITLE)
        self.setModal(True)
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(14)

        self.title_label = QLabel(ui_strings.OB_TITLE)
        self.title_label.setObjectName("CardTitle")
        self.title_label.setWordWrap(True)
        layout.addWidget(self.title_label)

        self.hint_label = QLabel(ui_strings.OB_HINT)
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        self.dart_label = QLabel(ui_strings.OB_DART_BODY)
        self.dart_label.setOpenExternalLinks(True)
        self.dart_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.dart_label.setWordWrap(True)
        layout.addWidget(self.dart_label)

        actions = QHBoxLayout()
        actions.addStretch(1)
        self.later_btn = QPushButton(ui_strings.OB_LATER)
        self.later_btn.clicked.connect(self.accept)
        self.settings_btn = QPushButton(ui_strings.OB_SETTINGS)
        self.settings_btn.setProperty("variant", "primary")
        self.settings_btn.clicked.connect(self._accept_for_settings)
        actions.addWidget(self.later_btn)
        actions.addWidget(self.settings_btn)
        layout.addLayout(actions)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def _retranslate(self, _lang: str = "") -> None:
        self.setWindowTitle(ui_strings.OB_WINDOW_TITLE)
        self.title_label.setText(ui_strings.OB_TITLE)
        self.hint_label.setText(ui_strings.OB_HINT)
        self.dart_label.setText(ui_strings.OB_DART_BODY)
        self.later_btn.setText(ui_strings.OB_LATER)
        self.settings_btn.setText(ui_strings.OB_SETTINGS)

    def _accept_for_settings(self) -> None:
        self.open_settings_requested = True
        self.accept()

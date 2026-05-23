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

        title = QLabel(ui_strings.OB_TITLE)
        title.setObjectName("CardTitle")
        title.setWordWrap(True)
        layout.addWidget(title)

        hint = QLabel(ui_strings.OB_HINT)
        hint.setWordWrap(True)
        layout.addWidget(hint)

        dart = QLabel(ui_strings.OB_DART_BODY)
        dart.setOpenExternalLinks(True)
        dart.setTextInteractionFlags(Qt.TextBrowserInteraction)
        dart.setWordWrap(True)
        layout.addWidget(dart)

        actions = QHBoxLayout()
        actions.addStretch(1)
        later_btn = QPushButton(ui_strings.OB_LATER)
        later_btn.clicked.connect(self.accept)
        settings_btn = QPushButton(ui_strings.OB_SETTINGS)
        settings_btn.setProperty("variant", "primary")
        settings_btn.clicked.connect(self._accept_for_settings)
        actions.addWidget(later_btn)
        actions.addWidget(settings_btn)
        layout.addLayout(actions)

    def _accept_for_settings(self) -> None:
        self.open_settings_requested = True
        self.accept()

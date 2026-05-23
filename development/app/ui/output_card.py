from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager


class OutputCard(QFrame):
    path_changed = Signal(str)

    def __init__(self, initial_path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        self.setProperty("class", "Card")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QWidget()
        header.setObjectName("CardHeader")
        h = QHBoxLayout(header)
        h.setContentsMargins(20, 16, 20, 8)
        self.title_label = QLabel(ui_strings.OC_TITLE)
        self.title_label.setObjectName("CardTitle")
        h.addWidget(self.title_label)
        h.addStretch(1)
        outer.addWidget(header)

        body = QWidget()
        body.setObjectName("CardBody")
        b = QHBoxLayout(body)
        b.setContentsMargins(20, 0, 20, 16)
        b.setSpacing(10)

        self.path_input = QLineEdit(initial_path)
        self.path_input.textChanged.connect(self.path_changed.emit)
        self.browse_btn = QPushButton(ui_strings.OC_BROWSE)
        self.browse_btn.setCursor(Qt.PointingHandCursor)
        self.browse_btn.clicked.connect(self._browse)
        self.open_btn = QPushButton(ui_strings.OC_OPEN_DIR)
        self.open_btn.setCursor(Qt.PointingHandCursor)
        self.open_btn.clicked.connect(self._open_dir)

        b.addWidget(self.path_input, 1)
        b.addWidget(self.browse_btn)
        b.addWidget(self.open_btn)
        outer.addWidget(body)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def path(self) -> str:
        return self.path_input.text().strip()

    def _retranslate(self, _lang: str = "") -> None:
        self.title_label.setText(ui_strings.OC_TITLE)
        self.browse_btn.setText(ui_strings.OC_BROWSE)
        self.open_btn.setText(ui_strings.OC_OPEN_DIR)

    def _browse(self) -> None:
        d = QFileDialog.getExistingDirectory(
            self, ui_strings.OC_SELECT_DIR_TITLE, self.path() or str(Path.home())
        )
        if d:
            self.path_input.setText(d)

    def _open_dir(self) -> None:
        import os
        import subprocess
        import sys

        p = Path(self.path())
        p.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(str(p))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])

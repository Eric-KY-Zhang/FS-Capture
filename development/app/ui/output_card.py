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
        title = QLabel(ui_strings.OC_TITLE)
        title.setObjectName("CardTitle")
        h.addWidget(title)
        h.addStretch(1)
        outer.addWidget(header)

        body = QWidget()
        body.setObjectName("CardBody")
        b = QHBoxLayout(body)
        b.setContentsMargins(20, 0, 20, 16)
        b.setSpacing(10)

        self.path_input = QLineEdit(initial_path)
        self.path_input.textChanged.connect(self.path_changed.emit)
        browse = QPushButton(ui_strings.OC_BROWSE)
        browse.setCursor(Qt.PointingHandCursor)
        browse.clicked.connect(self._browse)
        open_btn = QPushButton(ui_strings.OC_OPEN_DIR)
        open_btn.setCursor(Qt.PointingHandCursor)
        open_btn.clicked.connect(self._open_dir)

        b.addWidget(self.path_input, 1)
        b.addWidget(browse)
        b.addWidget(open_btn)
        outer.addWidget(body)

    def path(self) -> str:
        return self.path_input.text().strip()

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

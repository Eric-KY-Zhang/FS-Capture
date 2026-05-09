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


class OnboardingDialog(QDialog):
    """First-run guide for the PDF download workflow."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.open_settings_requested = False
        self.setWindowTitle("欢迎使用 FS Capture")
        self.setModal(True)
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(14)

        title = QLabel("FS Capture 帮你一键下载 4 市场上市公司官方披露 PDF")
        title.setObjectName("CardTitle")
        title.setWordWrap(True)
        layout.addWidget(title)

        hint = QLabel("输入第一个股票代码试试。确认公司名称后，选择年份和报告类型即可开始下载。")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        dart = QLabel(
            "韩股需要 DART API Key。可在 "
            '<a href="https://opendart.fss.or.kr/">opendart.fss.or.kr</a> 注册后粘贴到设置。'
        )
        dart.setOpenExternalLinks(True)
        dart.setTextInteractionFlags(Qt.TextBrowserInteraction)
        dart.setWordWrap(True)
        layout.addWidget(dart)

        actions = QHBoxLayout()
        actions.addStretch(1)
        later_btn = QPushButton("稍后再说")
        later_btn.clicked.connect(self.accept)
        settings_btn = QPushButton("现在就配 DART")
        settings_btn.setProperty("variant", "primary")
        settings_btn.clicked.connect(self._accept_for_settings)
        actions.addWidget(later_btn)
        actions.addWidget(settings_btn)
        layout.addLayout(actions)

    def _accept_for_settings(self) -> None:
        self.open_settings_requested = True
        self.accept()

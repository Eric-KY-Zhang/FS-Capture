from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.core.models import Exchange
from app.ui.styles.palette import exchange_accent


class ExchangeChip(QPushButton):
    """A clickable card representing one exchange. Acts as a toggle."""

    def __init__(self, exchange: Exchange, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.exchange = exchange
        self.setObjectName("ExchangeChip")
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(72)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.toggled.connect(self._sync_active_property)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(12)

        accent = QLabel()
        accent.setFixedSize(6, 28)
        accent.setStyleSheet(
            f"background:{exchange_accent(exchange)}; border-radius:3px;"
        )

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        text_box.setContentsMargins(0, 0, 0, 0)
        name = QLabel(self._zh_name(exchange))
        name.setObjectName("ExchangeChipName")
        meta = QLabel(self._meta(exchange))
        meta.setObjectName("ExchangeChipMeta")
        text_box.addWidget(name)
        text_box.addWidget(meta)

        layout.addWidget(accent)
        layout.addLayout(text_box, 1)

        self.check_indicator = QLabel("○")
        self.check_indicator.setStyleSheet(
            "color: #CBD5E1; font-size: 18px; font-weight: 700;"
        )
        layout.addWidget(self.check_indicator)

    def _sync_active_property(self, checked: bool) -> None:
        self.setProperty("active", "true" if checked else "false")
        self.style().unpolish(self)
        self.style().polish(self)
        self.check_indicator.setText("●" if checked else "○")
        self.check_indicator.setStyleSheet(
            f"color: {exchange_accent(self.exchange) if checked else '#CBD5E1'};"
            f" font-size: 18px; font-weight: 700;"
        )

    @staticmethod
    def _zh_name(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: "A股",
            Exchange.HK: "港股",
            Exchange.US: "美股",
            Exchange.KR: "韩股",
        }[exchange]

    @staticmethod
    def _meta(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: "上交所 · 深交所 · 北交所",
            Exchange.HK: "香港交易所",
            Exchange.US: "NYSE · NASDAQ",
            Exchange.KR: "KOSPI · KOSDAQ",
        }[exchange]


class ExchangeSelector(QWidget):
    """Top row of 4 chip-style exchange toggles. Emits selection_changed when
    any chip is toggled."""

    selection_changed = Signal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        self.chips: dict[Exchange, ExchangeChip] = {}
        for ex in (Exchange.A_SHARE, Exchange.HK, Exchange.US, Exchange.KR):
            chip = ExchangeChip(ex, self)
            chip.toggled.connect(lambda *_: self.selection_changed.emit())
            self.chips[ex] = chip
            layout.addWidget(chip, 1)

        # Default: A-share enabled
        self.chips[Exchange.A_SHARE].setChecked(True)

    def selected(self) -> list[Exchange]:
        return [ex for ex, c in self.chips.items() if c.isChecked()]

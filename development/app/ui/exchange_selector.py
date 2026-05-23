from __future__ import annotations

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
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager
from app.ui.styles.palette import exchange_accent


class ExchangeChip(QPushButton):
    """A clickable card representing one exchange. Acts as a toggle."""

    def __init__(self, exchange: Exchange, parent: QWidget | None = None) -> None:
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
        accent.setStyleSheet(f"background:{exchange_accent(exchange)}; border-radius:3px;")

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        text_box.setContentsMargins(0, 0, 0, 0)
        self.name_label = QLabel(self._zh_name(exchange))
        self.name_label.setObjectName("ExchangeChipName")
        self.meta_label = QLabel(self._meta(exchange))
        self.meta_label.setObjectName("ExchangeChipMeta")
        text_box.addWidget(self.name_label)
        text_box.addWidget(self.meta_label)

        layout.addWidget(accent)
        layout.addLayout(text_box, 1)

        self.check_indicator = QLabel("○")
        self.check_indicator.setStyleSheet("color: #CBD5E1; font-size: 18px; font-weight: 700;")
        layout.addWidget(self.check_indicator)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def _sync_active_property(self, checked: bool) -> None:
        self.setProperty("active", "true" if checked else "false")
        self.style().unpolish(self)
        self.style().polish(self)
        self.check_indicator.setText("●" if checked else "○")
        self.check_indicator.setStyleSheet(
            f"color: {exchange_accent(self.exchange) if checked else '#CBD5E1'};"
            f" font-size: 18px; font-weight: 700;"
        )

    def _retranslate(self, _lang: str = "") -> None:
        self.name_label.setText(self._zh_name(self.exchange))
        self.meta_label.setText(self._meta(self.exchange))

    @staticmethod
    def _zh_name(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: ui_strings.ES_NAME_A_SHARE,
            Exchange.HK: ui_strings.ES_NAME_HK,
            Exchange.US: ui_strings.ES_NAME_US,
            Exchange.KR: ui_strings.ES_NAME_KR,
            Exchange.TW: ui_strings.ES_NAME_TW,
            Exchange.JP: ui_strings.ES_NAME_JP,
        }[exchange]

    @staticmethod
    def _meta(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: ui_strings.ES_META_A_SHARE,
            Exchange.HK: ui_strings.ES_META_HK,
            Exchange.US: ui_strings.ES_META_US,
            Exchange.KR: ui_strings.ES_META_KR,
            Exchange.TW: ui_strings.ES_META_TW,
            Exchange.JP: ui_strings.ES_META_JP,
        }[exchange]


class ExchangeSelector(QWidget):
    """Top row of 4 chip-style exchange toggles. Emits selection_changed when
    any chip is toggled."""

    selection_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        self.chips: dict[Exchange, ExchangeChip] = {}
        for ex in (
            Exchange.A_SHARE,
            Exchange.HK,
            Exchange.US,
            Exchange.KR,
            Exchange.TW,
            Exchange.JP,
        ):
            chip = ExchangeChip(ex, self)
            chip.toggled.connect(lambda *_: self.selection_changed.emit())
            self.chips[ex] = chip
            layout.addWidget(chip, 1)

        # Default: A-share enabled
        self.chips[Exchange.A_SHARE].setChecked(True)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def selected(self) -> list[Exchange]:
        return [ex for ex, c in self.chips.items() if c.isChecked()]

    def _retranslate(self, _lang: str = "") -> None:
        for chip in self.chips.values():
            chip._retranslate()

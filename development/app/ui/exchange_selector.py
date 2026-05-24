from __future__ import annotations

from PySide6.QtCore import QRectF, QSize, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.core.models import Exchange
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager
from app.ui.styles.palette import exchange_accent

MARKET_ORDER: tuple[Exchange, ...] = (
    Exchange.UK,
    Exchange.US,
    Exchange.A_SHARE,
    Exchange.HK,
    Exchange.TW,
    Exchange.KR,
    Exchange.JP,
    Exchange.SG,
)


class _MarketPinGlyph(QWidget):
    def __init__(self, exchange: Exchange, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.exchange = exchange
        self._active = False
        self.setFixedSize(28, 34)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)

    def set_active(self, active: bool) -> None:
        self._active = active
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802  # type: ignore[override]
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        accent = QColor(exchange_accent(self.exchange))
        shadow = QColor(15, 23, 42, 44 if self._active else 28)
        path = self._pin_path(offset_y=1.5)

        painter.save()
        painter.translate(0, 2)
        painter.fillPath(path, shadow)
        painter.restore()

        if self._active:
            ring = self._pin_path(offset_y=0, inset=-2.2)
            ring_color = QColor(accent)
            ring_color.setAlpha(54)
            painter.fillPath(ring, ring_color)

        painter.fillPath(path, accent)
        painter.setPen(QPen(QColor(255, 255, 255, 210), 1.2))
        painter.drawPath(path)

        dot = QRectF(10.6, 9.4, 6.8, 6.8)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(255, 255, 255))
        painter.drawEllipse(dot)

        route_dot = QColor("#38BDF8")
        route_dot.setAlpha(230 if self._active else 190)
        painter.setBrush(route_dot)
        painter.drawEllipse(QRectF(12.4, 11.2, 3.2, 3.2))

    @staticmethod
    def _pin_path(offset_y: float, inset: float = 0) -> QPainterPath:
        left = 4.5 + inset
        right = 23.5 - inset
        top = 2.5 + offset_y + inset
        bottom = 31.0 + offset_y - inset
        cx = 14.0
        mid = 18.0 + offset_y

        path = QPainterPath()
        path.moveTo(cx, bottom)
        path.cubicTo(right + 1.4, mid + 2.0, right + 1.0, top + 2.0, cx, top)
        path.cubicTo(left - 1.0, top + 2.0, left - 1.4, mid + 2.0, cx, bottom)
        path.closeSubpath()
        return path


class MarketPin(QWidget):
    """A compact market toggle using the Atlas map-pin mark."""

    toggled = Signal(bool)

    def __init__(self, exchange: Exchange, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.exchange = exchange
        self._checked = False
        self.setObjectName("MarketPin")
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setMinimumSize(QSize(90, 86))
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 8, 6, 8)
        layout.setSpacing(4)
        layout.setAlignment(Qt.AlignCenter)

        self.pin_glyph = _MarketPinGlyph(exchange, self)
        layout.addWidget(self.pin_glyph, 0, Qt.AlignHCenter)

        self.name_label = QLabel(self._name(exchange))
        self.name_label.setObjectName("MarketPinName")
        self.name_label.setAlignment(Qt.AlignCenter)
        self.name_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        self.meta_label = QLabel(self._meta(exchange))
        self.meta_label.setObjectName("MarketPinMeta")
        self.meta_label.setAlignment(Qt.AlignCenter)
        self.meta_label.setWordWrap(False)
        self.meta_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        layout.addWidget(self.name_label)
        layout.addWidget(self.meta_label)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def setChecked(self, checked: bool) -> None:  # noqa: N802
        if self._checked == checked:
            return
        self._checked = checked
        self._sync_active_property()
        self.toggled.emit(checked)

    def isChecked(self) -> bool:  # noqa: N802
        return self._checked

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802  # type: ignore[override]
        if event.button() == Qt.LeftButton and self.rect().contains(event.position().toPoint()):
            self.setChecked(not self._checked)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event) -> None:  # noqa: N802  # type: ignore[override]
        if event.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Space):
            self.setChecked(not self._checked)
            event.accept()
            return
        super().keyPressEvent(event)

    def _sync_active_property(self) -> None:
        self.setProperty("active", "true" if self._checked else "false")
        self.pin_glyph.set_active(self._checked)
        self.style().unpolish(self)
        self.style().polish(self)

    def _retranslate(self, _lang: str = "") -> None:
        self.name_label.setText(self._name(self.exchange))
        self.meta_label.setText(self._meta(self.exchange))

    @staticmethod
    def _name(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: ui_strings.MP_NAME_A_SHARE,
            Exchange.HK: ui_strings.MP_NAME_HK,
            Exchange.US: ui_strings.MP_NAME_US,
            Exchange.KR: ui_strings.MP_NAME_KR,
            Exchange.TW: ui_strings.MP_NAME_TW,
            Exchange.JP: ui_strings.MP_NAME_JP,
            Exchange.UK: ui_strings.MP_NAME_UK,
            Exchange.SG: ui_strings.MP_NAME_SG,
        }[exchange]

    @staticmethod
    def _meta(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: ui_strings.MP_META_A_SHARE,
            Exchange.HK: ui_strings.MP_META_HK,
            Exchange.US: ui_strings.MP_META_US,
            Exchange.KR: ui_strings.MP_META_KR,
            Exchange.TW: ui_strings.MP_META_TW,
            Exchange.JP: ui_strings.MP_META_JP,
            Exchange.UK: ui_strings.MP_META_UK,
            Exchange.SG: ui_strings.MP_META_SG,
        }[exchange]


class ExchangeSelector(QWidget):
    """Horizontal row of MarketPin toggles."""

    selection_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("MarketPinSelector")
        self.setAttribute(Qt.WA_StyledBackground, True)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        self.pins: dict[Exchange, MarketPin] = {}
        for exchange in MARKET_ORDER:
            pin = MarketPin(exchange, self)
            pin.toggled.connect(lambda *_: self.selection_changed.emit())
            self.pins[exchange] = pin
            layout.addWidget(pin, 1)

        self.pins[Exchange.A_SHARE].setChecked(True)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def selected(self) -> list[Exchange]:
        return [exchange for exchange, pin in self.pins.items() if pin.isChecked()]

    def _retranslate(self, _lang: str = "") -> None:
        for pin in self.pins.values():
            pin._retranslate()

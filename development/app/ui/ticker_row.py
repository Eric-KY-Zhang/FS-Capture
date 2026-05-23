from __future__ import annotations

from PySide6.QtCore import QObject, QRunnable, QSignalBlocker, Qt, QThreadPool, Signal, Slot
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSizePolicy,
    QWidget,
)

from app.core.models import Exchange, Ticker
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager


class _ResolveSignals(QObject):
    finished = Signal(object, object, str)  # ticker (or None), code, error_msg


class _ResolveRunnable(QRunnable):
    def __init__(self, exchange: Exchange, code: str, signals: _ResolveSignals) -> None:
        super().__init__()
        self.exchange = exchange
        self.code = code
        self.signals = signals

    @Slot()
    def run(self) -> None:
        try:
            from plugins import get_plugin

            plugin = get_plugin(self.exchange)
            ticker = plugin.resolve_name(self.code)
            self.signals.finished.emit(ticker, self.code, "")
        except Exception as exc:  # noqa: BLE001
            self.signals.finished.emit(None, self.code, f"{type(exc).__name__}: {exc}")


class TickerRow(QWidget):
    """A single row in an ExchangePanel. Holds a stock code, a confirm action,
    and the resolved company name (or an error pill).
    """

    removed = Signal(QWidget)  # emitted when user clicks ×
    resolved = Signal(QWidget)  # emitted when name resolved successfully (or cleared)

    def __init__(self, exchange: Exchange, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.exchange = exchange
        self._ticker: Ticker | None = None
        self._pill_key = "TR_PENDING"

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 4)
        layout.setSpacing(10)

        self.code_input = QLineEdit()
        self.code_input.setPlaceholderText(self._placeholder_for(exchange))
        self.code_input.setMaximumWidth(160)
        self.code_input.returnPressed.connect(self.confirm)
        self.code_input.textChanged.connect(self._on_text_changed)

        self.confirm_btn = QPushButton(ui_strings.TR_CONFIRM)
        self.confirm_btn.setProperty("variant", "primary")
        self.confirm_btn.setCursor(Qt.PointingHandCursor)
        self.confirm_btn.clicked.connect(self.confirm)

        self.status_pill = QLabel(ui_strings.TR_PENDING)
        self.status_pill.setObjectName("StatusPill")
        self.status_pill.setProperty("state", "pending")
        self.status_pill.setAlignment(Qt.AlignCenter)
        self.status_pill.setMinimumWidth(70)

        self.name_label = QLabel("")
        self.name_label.setStyleSheet("color: #475569; font-size: 13px;")
        self.name_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.name_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

        self.delete_btn = QPushButton("×")
        self.delete_btn.setProperty("variant", "danger-icon")
        self.delete_btn.setCursor(Qt.PointingHandCursor)
        self.delete_btn.clicked.connect(self._on_delete_clicked)
        self.delete_btn.setToolTip(ui_strings.TR_DELETE_TOOLTIP)

        layout.addWidget(self.code_input)
        layout.addWidget(self.confirm_btn)
        layout.addWidget(self.status_pill)
        layout.addWidget(self.name_label, 1)
        layout.addWidget(self.delete_btn)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    # ---- public API ---------------------------------------------------------

    @property
    def ticker(self) -> Ticker | None:
        return self._ticker

    def is_resolved(self) -> bool:
        return self._ticker is not None

    # ---- actions ------------------------------------------------------------

    def confirm(self) -> None:
        code = self.code_input.text().strip()
        if not code:
            self._set_state("error", "TR_ENTER_CODE", "")
            return
        self._set_state("resolving", "TR_RESOLVING", "")
        self.confirm_btn.setEnabled(False)
        self.code_input.setEnabled(False)

        signals = _ResolveSignals()
        signals.finished.connect(self._on_resolved)
        QThreadPool.globalInstance().start(_ResolveRunnable(self.exchange, code, signals))

    def _on_resolved(self, ticker: Ticker | None, code: str, error: str) -> None:
        self.confirm_btn.setEnabled(True)
        self.code_input.setEnabled(True)
        if ticker is not None:
            blocker = QSignalBlocker(self.code_input)
            self.code_input.setText(ticker.code)
            del blocker
            self._ticker = ticker
            self._set_state("ok", "TR_RESOLVED", ticker.name or "")
        else:
            self._ticker = None
            self._set_state("error", "TR_NOT_FOUND", error or ui_strings.TR_UNRESOLVABLE)
        self.resolved.emit(self)

    # ---- internals ----------------------------------------------------------

    def _set_state(self, state: str, pill_key: str, name_text: str) -> None:
        self._pill_key = pill_key
        self.status_pill.setProperty("state", state)
        self.status_pill.setText(getattr(ui_strings, pill_key))
        # Force QSS re-evaluation after dynamic property change
        self.status_pill.style().unpolish(self.status_pill)
        self.status_pill.style().polish(self.status_pill)
        self.name_label.setText(name_text)

    def _on_text_changed(self, _text: str) -> None:
        if self._ticker is not None:
            self._ticker = None
            self._set_state("pending", "TR_PENDING", "")
            self.resolved.emit(self)

    def _on_delete_clicked(self) -> None:
        self.removed.emit(self)

    def _retranslate(self, _lang: str = "") -> None:
        self.code_input.setPlaceholderText(self._placeholder_for(self.exchange))
        self.confirm_btn.setText(ui_strings.TR_CONFIRM)
        self.status_pill.setText(getattr(ui_strings, self._pill_key))
        self.delete_btn.setToolTip(ui_strings.TR_DELETE_TOOLTIP)

    @staticmethod
    def _placeholder_for(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: ui_strings.TR_PLACEHOLDER_A_SHARE,
            Exchange.HK: ui_strings.TR_PLACEHOLDER_HK,
            Exchange.US: ui_strings.TR_PLACEHOLDER_US,
            Exchange.KR: ui_strings.TR_PLACEHOLDER_KR,
            Exchange.TW: ui_strings.TR_PLACEHOLDER_TW,
        }[exchange]

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.core.models import Exchange, Ticker
from app.ui.batch_import_dialog import BatchImportDialog
from app.ui.styles.palette import exchange_accent
from app.ui.ticker_row import TickerRow


class ExchangePanel(QFrame):
    """Card-style section listing all rows for one exchange."""

    rows_changed = Signal()

    def __init__(self, exchange: Exchange, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.exchange = exchange
        self.setObjectName("Card")
        self.setProperty("class", "Card")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Header
        header = QWidget()
        header.setObjectName("CardHeader")
        h_layout = QHBoxLayout(header)
        h_layout.setContentsMargins(20, 16, 20, 12)
        h_layout.setSpacing(12)

        # Color stripe
        accent = QLabel()
        accent.setFixedSize(3, 22)
        accent.setStyleSheet(f"background:{exchange_accent(exchange)}; border-radius:2px;")
        h_layout.addWidget(accent)

        title = QLabel(self._title_for(exchange))
        title.setObjectName("CardTitle")
        h_layout.addWidget(title)

        sub = QLabel(self._subtitle_for(exchange))
        sub.setObjectName("CardSubtitle")
        sub.setStyleSheet("color: #94A3B8; font-size: 12px;")
        h_layout.addWidget(sub)
        h_layout.addStretch(1)

        self.batch_btn = QPushButton("＋ 批量添加")
        self.batch_btn.setProperty("variant", "ghost")
        self.batch_btn.setCursor(Qt.PointingHandCursor)
        self.batch_btn.clicked.connect(self._open_batch_import)
        h_layout.addWidget(self.batch_btn)

        self.add_btn = QPushButton("＋ 单只添加")
        self.add_btn.setProperty("variant", "ghost")
        self.add_btn.setCursor(Qt.PointingHandCursor)
        self.add_btn.clicked.connect(self.add_row)
        h_layout.addWidget(self.add_btn)

        outer.addWidget(header)

        # Body — vertical list of TickerRow
        self._rows_host = QWidget()
        self._rows_host.setObjectName("CardBody")
        self._rows_layout = QVBoxLayout(self._rows_host)
        self._rows_layout.setContentsMargins(20, 0, 20, 16)
        self._rows_layout.setSpacing(2)

        outer.addWidget(self._rows_host)

        # Empty state hint
        self._empty_label = QLabel("暂无股票，点击右上角添加或批量导入")
        self._empty_label.setStyleSheet("color: #94A3B8; font-size: 12px; padding: 12px 4px;")
        self._rows_layout.addWidget(self._empty_label)

        self._rows: list[TickerRow] = []

    # ---- API ---------------------------------------------------------------

    def add_row(self, code: str = "") -> TickerRow:
        if self._empty_label.isVisible():
            self._empty_label.hide()
        row = TickerRow(self.exchange, self._rows_host)
        if code:
            row.code_input.setText(code)
        row.removed.connect(self._on_row_removed)
        row.resolved.connect(lambda *_: self.rows_changed.emit())
        self._rows.append(row)
        self._rows_layout.addWidget(row)
        self.rows_changed.emit()
        return row

    def add_codes(self, codes: list[str], *, auto_confirm: bool = True) -> int:
        existing = {r.code_input.text().strip().upper() for r in self._rows}
        added = 0
        for code in codes:
            if code.upper() in existing:
                continue
            row = self.add_row(code)
            existing.add(code.upper())
            added += 1
            if auto_confirm:
                row.confirm()
        return added

    def resolved_tickers(self) -> list[Ticker]:
        out: list[Ticker] = []
        for r in self._rows:
            t = r.ticker
            if t is not None:
                out.append(t)
        return out

    def has_pending(self) -> bool:
        for r in self._rows:
            if r.code_input.text().strip() and not r.is_resolved():
                return True
        return False

    # ---- internals ---------------------------------------------------------

    def _on_row_removed(self, row: QWidget) -> None:
        if row in self._rows:
            self._rows.remove(row)
        self._rows_layout.removeWidget(row)
        row.deleteLater()
        if not self._rows:
            self._empty_label.show()
        self.rows_changed.emit()

    def _open_batch_import(self) -> None:
        dialog = BatchImportDialog(self.exchange, self)
        if dialog.exec() != dialog.DialogCode.Accepted:
            return

        codes = dialog.codes()
        if not codes:
            QMessageBox.warning(self, "没有可添加的代码", "未识别到有效股票代码")
            return

        added = self.add_codes(codes, auto_confirm=dialog.auto_confirm.isChecked())
        skipped = len(codes) - added
        if skipped:
            QMessageBox.information(
                self,
                "批量添加完成",
                f"已添加 {added} 只股票，跳过 {skipped} 个重复代码。",
            )

    @staticmethod
    def _title_for(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: "A股 · A-Share",
            Exchange.HK: "港股 · Hong Kong",
            Exchange.US: "美股 · United States",
            Exchange.KR: "韩股 · Korea",
            Exchange.TW: "台股 · Taiwan",
        }[exchange]

    @staticmethod
    def _subtitle_for(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: "巨潮资讯网 · 东方财富",
            Exchange.HK: "披露易 · 东方财富",
            Exchange.US: "SEC EDGAR",
            Exchange.KR: "DART 电子公示",
            Exchange.TW: "公開資訊觀測站 MOPS",
        }[exchange]

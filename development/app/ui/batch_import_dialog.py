from __future__ import annotations

import re

from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.core.models import Exchange

_CELL_SPLIT_RE = re.compile(r"[\t,，;；、]+")
_SPACE_SPLIT_RE = re.compile(r"\s+")

_A_SHARE_RE = re.compile(r"^(?:SH|SZ|BJ)?\d{6}(?:\.(?:SH|SZ|BJ))?$", re.IGNORECASE)
_HK_RE = re.compile(r"^(?:HK)?\d{1,5}(?:\.HK)?$", re.IGNORECASE)
_US_RE = re.compile(r"^[A-Z]{1,5}(?:[.-][A-Z]{1,2})?$", re.IGNORECASE)
_KR_RE = re.compile(r"^\d{1,6}(?:\.(?:KS|KQ))?$", re.IGNORECASE)
_TW_RE = re.compile(r"^(?:TW)?\d{4}(?:\.(?:TW|TWO))?$", re.IGNORECASE)

_PATTERNS = {
    Exchange.A_SHARE: _A_SHARE_RE,
    Exchange.HK: _HK_RE,
    Exchange.US: _US_RE,
    Exchange.KR: _KR_RE,
    Exchange.TW: _TW_RE,
}

_US_STOPWORDS = {
    "ADR",
    "AMEX",
    "CLASS",
    "CO",
    "CORP",
    "INC",
    "LLC",
    "LTD",
    "NASDAQ",
    "NYSE",
    "PLC",
    "THE",
}


def parse_ticker_codes(text: str, exchange: Exchange) -> list[str]:
    """Parse pasted ticker text from lines, CSV cells, or Excel columns."""
    out: list[str] = []
    seen: set[str] = set()
    for token in _candidate_tokens(text, exchange):
        code = _normalize_token(token, exchange)
        if not code or code in seen:
            continue
        seen.add(code)
        out.append(code)
    return out


def _candidate_tokens(text: str, exchange: Exchange):
    for line in text.splitlines():
        for cell in _CELL_SPLIT_RE.split(line):
            cell = cell.strip()
            if not cell:
                continue
            tokens = [t for t in _SPACE_SPLIT_RE.split(cell) if t]
            if exchange == Exchange.US and len(tokens) > 1:
                yield tokens[0]
            else:
                yield from tokens


def _normalize_token(token: str, exchange: Exchange) -> str:
    code = token.strip().strip("'\"`“”‘’()[]{}").upper()
    code = code.rstrip(".")
    if not code or code in _US_STOPWORDS:
        return ""
    if not _PATTERNS[exchange].match(code):
        return ""

    if exchange == Exchange.A_SHARE:
        code = re.sub(r"^(SH|SZ|BJ)", "", code)
        code = re.sub(r"\.(SH|SZ|BJ)$", "", code)
        return code
    if exchange == Exchange.HK:
        code = code.removeprefix("HK").removesuffix(".HK")
        return code.zfill(5)
    if exchange == Exchange.KR:
        code = re.sub(r"\.(KS|KQ)$", "", code)
        return code.zfill(6)
    if exchange == Exchange.TW:
        code = code.removeprefix("TW")
        code = re.sub(r"\.(TW|TWO)$", "", code)
        return code.zfill(4)
    return code


class BatchImportDialog(QDialog):
    """Dialog for pasting a batch of ticker codes."""

    def __init__(self, exchange: Exchange, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.exchange = exchange
        self.setWindowTitle(f"批量添加{exchange.display_name}股票")
        self.setModal(True)
        self.setMinimumSize(520, 380)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(12)

        title = QLabel("粘贴股票代码")
        title.setObjectName("CardTitle")
        layout.addWidget(title)

        hint = QLabel("支持从 Excel、网页或文本复制，多行、逗号和制表符会自动拆分。")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.editor = QTextEdit()
        self.editor.setPlaceholderText(self._placeholder_for(exchange))
        self.editor.setAcceptRichText(False)
        layout.addWidget(self.editor, 1)

        self.auto_confirm = QCheckBox("添加后自动确认公司名称")
        self.auto_confirm.setChecked(True)
        layout.addWidget(self.auto_confirm)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("添加")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def codes(self) -> list[str]:
        return parse_ticker_codes(self.editor.toPlainText(), self.exchange)

    @staticmethod
    def _placeholder_for(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: "600519\n000001\n300750",
            Exchange.HK: "00700\n09988\n00005",
            Exchange.US: "AAPL\nMSFT\nBRK.B",
            Exchange.KR: "005930\n000660\n035420",
            Exchange.TW: "2330\n2317\n2454",
        }[exchange]

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
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager

_CELL_SPLIT_RE = re.compile(r"[\t,，;；、]+")
_SPACE_SPLIT_RE = re.compile(r"\s+")

_A_SHARE_RE = re.compile(r"^(?:SH|SZ|BJ)?\d{6}(?:\.(?:SH|SZ|BJ))?$", re.IGNORECASE)
_HK_RE = re.compile(r"^(?:HK)?\d{1,5}(?:\.HK)?$", re.IGNORECASE)
_US_RE = re.compile(r"^[A-Z]{1,5}(?:[.-][A-Z]{1,2})?$", re.IGNORECASE)
_KR_RE = re.compile(r"^\d{1,6}(?:\.(?:KS|KQ))?$", re.IGNORECASE)
_TW_RE = re.compile(r"^(?:TW)?\d{4}(?:\.(?:TW|TWO))?$", re.IGNORECASE)
_JP_RE = re.compile(r"^(?:JP)?\d{4}(?:\.(?:T|JP))?$", re.IGNORECASE)

_PATTERNS = {
    Exchange.A_SHARE: _A_SHARE_RE,
    Exchange.HK: _HK_RE,
    Exchange.US: _US_RE,
    Exchange.KR: _KR_RE,
    Exchange.TW: _TW_RE,
    Exchange.JP: _JP_RE,
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


def parse_ticker_codes(text: str, exchange: Exchange) -> tuple[list[str], list[str]]:
    """Parse pasted ticker text and return (valid_codes, rejected_raw_tokens)."""
    out: list[str] = []
    rejected: list[str] = []
    seen: set[str] = set()
    for line in text.splitlines():
        line_rejected: list[str] = []
        line_has_valid_code = False
        for token in _candidate_tokens_from_line(line, exchange):
            code = _normalize_token(token, exchange)
            if not code:
                if not _is_ignored_token(token, exchange):
                    line_rejected.append(token)
                continue
            line_has_valid_code = True
            if code in seen:
                continue
            seen.add(code)
            out.append(code)
        if not line_has_valid_code:
            rejected.extend(line_rejected)
    return out, rejected


def _candidate_tokens(text: str, exchange: Exchange):
    for line in text.splitlines():
        yield from _candidate_tokens_from_line(line, exchange)


def _candidate_tokens_from_line(line: str, exchange: Exchange):
    for cell in _CELL_SPLIT_RE.split(line):
        cell = cell.strip()
        if not cell:
            continue
        tokens = [t for t in _SPACE_SPLIT_RE.split(cell) if t]
        if exchange == Exchange.US and len(tokens) > 1:
            yield tokens[0]
        else:
            yield from tokens


def _is_ignored_token(token: str, exchange: Exchange) -> bool:
    code = token.strip().strip("'\"`“”‘’()[]{}").upper().rstrip(".")
    if not code:
        return True
    if exchange == Exchange.US and code in _US_STOPWORDS:
        return True
    return any(0x4E00 <= ord(char) <= 0x9FFF for char in token)


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
    if exchange == Exchange.JP:
        code = code.removeprefix("JP")
        code = re.sub(r"\.(T|JP)$", "", code)
        return code.zfill(4)
    return code


class BatchImportDialog(QDialog):
    """Dialog for pasting a batch of ticker codes."""

    def __init__(self, exchange: Exchange, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.exchange = exchange
        self.setWindowTitle(
            ui_strings.BID_WINDOW_TITLE_FORMAT.format(
                exchange_name=self._exchange_name_for(exchange)
            )
        )
        self.setModal(True)
        self.setMinimumSize(520, 380)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(12)

        self.title_label = QLabel(ui_strings.BID_TITLE)
        self.title_label.setObjectName("CardTitle")
        layout.addWidget(self.title_label)

        self.hint_label = QLabel(ui_strings.BID_HINT)
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        self.editor = QTextEdit()
        self.editor.setPlaceholderText(self._placeholder_for(exchange))
        self.editor.setAcceptRichText(False)
        layout.addWidget(self.editor, 1)

        self.auto_confirm = QCheckBox(ui_strings.BID_AUTO_CONFIRM)
        self.auto_confirm.setChecked(True)
        layout.addWidget(self.auto_confirm)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText(ui_strings.BID_ADD)
        self.buttons.button(QDialogButtonBox.Cancel).setText(ui_strings.COMMON_CANCEL)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def codes(self) -> tuple[list[str], list[str]]:
        return parse_ticker_codes(self.editor.toPlainText(), self.exchange)

    def _retranslate(self, _lang: str = "") -> None:
        self.setWindowTitle(
            ui_strings.BID_WINDOW_TITLE_FORMAT.format(
                exchange_name=self._exchange_name_for(self.exchange)
            )
        )
        self.title_label.setText(ui_strings.BID_TITLE)
        self.hint_label.setText(ui_strings.BID_HINT)
        self.auto_confirm.setText(ui_strings.BID_AUTO_CONFIRM)
        self.buttons.button(QDialogButtonBox.Ok).setText(ui_strings.BID_ADD)
        self.buttons.button(QDialogButtonBox.Cancel).setText(ui_strings.COMMON_CANCEL)

    @staticmethod
    def _exchange_name_for(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: ui_strings.ES_NAME_A_SHARE,
            Exchange.HK: ui_strings.ES_NAME_HK,
            Exchange.US: ui_strings.ES_NAME_US,
            Exchange.KR: ui_strings.ES_NAME_KR,
            Exchange.TW: ui_strings.ES_NAME_TW,
            Exchange.JP: ui_strings.ES_NAME_JP,
        }[exchange]

    @staticmethod
    def _placeholder_for(exchange: Exchange) -> str:
        return {
            Exchange.A_SHARE: "600519\n000001\n300750",
            Exchange.HK: "00700\n09988\n00005",
            Exchange.US: "AAPL\nMSFT\nBRK.B",
            Exchange.KR: "005930\n000660\n035420",
            Exchange.TW: "2330\n2317\n2454",
            Exchange.JP: "7203\n6758\n9984",
        }[exchange]

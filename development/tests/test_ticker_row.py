from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from app.core.models import Exchange, Ticker
from app.ui.ticker_row import TickerRow


def _app() -> QApplication:
    return QApplication.instance() or QApplication([])


def test_resolved_normalized_code_keeps_ticker() -> None:
    _app()
    row = TickerRow(Exchange.HK)
    row.code_input.setText("0700")
    ticker = Ticker(
        exchange=Exchange.HK,
        code="00700",
        name="腾讯控股",
        external_id="7609",
    )

    row._on_resolved(ticker, "0700", "")

    assert row.code_input.text() == "00700"
    assert row.ticker == ticker
    assert row.is_resolved()

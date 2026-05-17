from __future__ import annotations

from app.core.models import Exchange
from app.ui.batch_import_dialog import parse_ticker_codes


def test_parse_ashare_codes_from_excel_cells() -> None:
    text = "600519\t贵州茅台\n000001, 平安银行\nSH300750\n600519.SH"

    assert parse_ticker_codes(text, Exchange.A_SHARE) == ["600519", "000001", "300750"]


def test_parse_hk_codes_normalizes_to_five_digits() -> None:
    text = "700 腾讯\nHK9988\n00005.HK\n00700"

    assert parse_ticker_codes(text, Exchange.HK) == ["00700", "09988", "00005"]


def test_parse_us_codes_keeps_class_suffix_and_skips_common_market_words() -> None:
    text = "AAPL Apple Inc.\nNYSE, MSFT\nbrk.b\nBRK-B"

    assert parse_ticker_codes(text, Exchange.US) == ["AAPL", "MSFT", "BRK.B", "BRK-B"]


def test_parse_kr_codes_zero_pads_and_deduplicates_suffixes() -> None:
    text = "5930 Samsung\n005930.KS\n000660;035420.KQ"

    assert parse_ticker_codes(text, Exchange.KR) == ["005930", "000660", "035420"]


def test_parse_tw_codes_normalizes_to_four_digits() -> None:
    text = "2330 台積電\nTW2317\n2454.TW\n2330"

    assert parse_ticker_codes(text, Exchange.TW) == ["2330", "2317", "2454"]

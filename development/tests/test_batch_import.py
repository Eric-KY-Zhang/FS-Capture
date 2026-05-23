from __future__ import annotations

from app.core.models import Exchange
from app.ui.batch_import_dialog import parse_ticker_codes


def test_parse_ashare_codes_from_excel_cells() -> None:
    text = "600519\t贵州茅台\n000001, 平安银行\nSH300750\n600519.SH"

    codes, rejected = parse_ticker_codes(text, Exchange.A_SHARE)

    assert codes == ["600519", "000001", "300750"]
    assert rejected == []


def test_parse_hk_codes_normalizes_to_five_digits() -> None:
    text = "700 腾讯\nHK9988\n00005.HK\n00700"

    codes, rejected = parse_ticker_codes(text, Exchange.HK)

    assert codes == ["00700", "09988", "00005"]
    assert rejected == []


def test_parse_us_codes_keeps_class_suffix_and_skips_common_market_words() -> None:
    text = "AAPL Apple Inc.\nNYSE, MSFT\nbrk.b\nBRK-B"

    codes, rejected = parse_ticker_codes(text, Exchange.US)

    assert codes == ["AAPL", "MSFT", "BRK.B", "BRK-B"]
    assert rejected == []


def test_parse_kr_codes_zero_pads_and_deduplicates_suffixes() -> None:
    text = "5930 Samsung\n005930.KS\n000660;035420.KQ"

    codes, rejected = parse_ticker_codes(text, Exchange.KR)

    assert codes == ["005930", "000660", "035420"]
    assert rejected == []


def test_parse_tw_codes_normalizes_to_four_digits() -> None:
    text = "2330 台積電\nTW2317\n2454.TW\n2330"

    codes, rejected = parse_ticker_codes(text, Exchange.TW)

    assert codes == ["2330", "2317", "2454"]
    assert rejected == []


def test_parse_returns_rejected_tokens() -> None:
    codes, rejected = parse_ticker_codes("600519\nINVALID\n000001", Exchange.A_SHARE)

    assert codes == ["600519", "000001"]
    assert rejected == ["INVALID"]


def test_parse_us_stopwords_not_in_rejected() -> None:
    codes, rejected = parse_ticker_codes("AAPL INC", Exchange.US)

    assert codes == ["AAPL"]
    assert rejected == []

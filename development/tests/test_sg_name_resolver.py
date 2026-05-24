from __future__ import annotations

import pytest

from app.core.models import Exchange, Ticker
from plugins.sg import name_resolver


def _row(code: str = "D05") -> dict:
    return {
        "issuers": [
            {
                "stock_code": code,
                "security_name": "DBS GROUP HOLDINGS LTD",
                "issuer_name": "DBS GROUP HOLDINGS LTD",
                "isin_code": "SG1L01001701",
                "ibm_code": "1L01",
            }
        ]
    }


@pytest.fixture(autouse=True)
def _no_disk_cache(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(name_resolver, "cached_or_load", lambda _key, load, expire=None: load())


@pytest.mark.parametrize(
    ("input_code", "expected"),
    [
        ("D05", "D05"),
        ("d05", "D05"),
        ("D05.SI", "D05"),
        ("d05.si", "D05"),
    ],
)
def test_resolve_normalizes_ticker(monkeypatch: pytest.MonkeyPatch, input_code: str, expected: str) -> None:
    monkeypatch.setattr(
        "plugins.sg.sgxnet_web.search_announcements",
        lambda *_args, **_kwargs: [_row(expected)],
    )

    ticker = name_resolver.resolve(input_code)

    assert ticker.code == expected
    assert ticker.name == "DBS GROUP HOLDINGS LTD"
    assert ticker.external_id == "SG1L01001701"
    assert ticker.exchange is Exchange.SG


def test_resolve_invalid_format_raises_value_error() -> None:
    with pytest.raises(ValueError, match="格式错误"):
        name_resolver.resolve("INVALID-CODE")


def test_resolve_unknown_ticker_raises_value_error(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("plugins.sg.sgxnet_web.search_announcements", lambda *_args, **_kwargs: [])
    monkeypatch.setattr("plugins.sg.sgxnet_web.list_ipo_prospectuses", lambda *_args, **_kwargs: [])

    with pytest.raises(ValueError, match="未找到新股代码"):
        name_resolver.resolve("ZZZZ")


def test_resolve_ipo_id_falls_back_to_ipo_prospectus(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("plugins.sg.sgxnet_web.search_announcements", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(
        "plugins.sg.sgxnet_web.list_ipo_prospectuses",
        lambda *_args, **_kwargs: [
            {
                "id": "3407",
                "name": "LION-CM EM ASIA INDEX ETF",
                "status": "Completed",
            }
        ],
    )

    ticker = name_resolver.resolve("3407")

    assert ticker.code == "3407"
    assert ticker.name == "LION-CM EM ASIA INDEX ETF"
    assert ticker.exchange is Exchange.SG


def test_fetch_company_returns_sgd_and_sgx_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(name_resolver, "resolve_one", lambda _code: name_resolver._issuer_from_rows("D05", [_row()]))

    company = name_resolver.fetch_company(
        Ticker(exchange=Exchange.SG, code="D05", name="DBS GROUP HOLDINGS LTD", external_id="SG1L01001701")
    )

    assert company.currency == "SGD"
    assert company.extra["security_name"] == "DBS GROUP HOLDINGS LTD"
    assert company.extra["ibm_code"] == "1L01"

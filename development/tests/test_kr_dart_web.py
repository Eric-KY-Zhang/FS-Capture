from __future__ import annotations

from pathlib import Path

from plugins.kr import dart_web


def _fixture(name: str) -> str:
    return (Path(__file__).parent / "fixtures" / name).read_text(encoding="utf-8")


def test_dart_web_resolve_corp_parses_search_result(monkeypatch) -> None:
    monkeypatch.setattr(
        dart_web,
        "_detail_search_html",
        lambda **_kwargs: _fixture("dart_search_005930.html"),
    )

    assert dart_web.resolve_corp("005930") == {
        "corp_code": "00126380",
        "corp_name": "삼성전자",
    }


def test_dart_web_resolve_corp_returns_none_for_unknown_code(monkeypatch) -> None:
    monkeypatch.setattr(dart_web, "_detail_search_html", lambda **_kwargs: "<tbody></tbody>")

    assert dart_web.resolve_corp("999999") is None

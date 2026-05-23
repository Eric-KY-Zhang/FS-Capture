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


def test_dart_web_list_filings_schema_matches_opendartreader(monkeypatch) -> None:
    calls = []

    def fake_detail_search(**kwargs):
        calls.append(kwargs)
        return _fixture("dart_search_005930.html")

    monkeypatch.setattr(dart_web, "_detail_search_html", fake_detail_search)

    df = dart_web.list_filings("00126380", "20240101", "20250630", "A001")

    assert list(df.columns) == ["rcept_no", "report_nm", "rcept_dt", "corp_code"]
    assert df.iloc[0].to_dict() == {
        "rcept_no": "20250311001085",
        "report_nm": "사업보고서 (2024.12)",
        "rcept_dt": "20250311",
        "corp_code": "00126380",
    }
    assert all(isinstance(value, str) for value in df.iloc[0].to_dict().values())
    assert calls == [
        {
            "corp_code": "00126380",
            "bgn_de": "20240101",
            "end_de": "20250630",
            "detail_type": "A001",
            "final": True,
        }
    ]

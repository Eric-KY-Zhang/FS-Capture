from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.core.settings import Settings
from plugins.kr import KRShare, dart_web


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


def test_dart_web_list_filings_retries_without_final_filter(monkeypatch) -> None:
    calls = []

    def fake_detail_search(**kwargs):
        calls.append(kwargs)
        if kwargs["final"]:
            return "<tbody></tbody>"
        return _fixture("dart_search_005930.html")

    monkeypatch.setattr(dart_web, "_detail_search_html", fake_detail_search)

    df = dart_web.list_filings("00126380", "20240101", "20250630", "A001")

    assert len(df) == 2
    assert [call["final"] for call in calls] == [True, False]


def test_dart_web_list_audit_filings_uses_i001_search(monkeypatch) -> None:
    calls = []
    html = """
    <tbody id="tbody">
      <tr>
        <td></td>
        <td><a href="javascript:openCorpInfoNew('00126380');">삼성전자</a></td>
        <td><a href="/dsaf001/main.do?rcpNo=20250218800508">감사보고서제출</a></td>
        <td></td>
        <td>2025.02.18</td>
      </tr>
    </tbody>
    """

    def fake_detail_search(**kwargs):
        calls.append(kwargs)
        return html

    monkeypatch.setattr(dart_web, "_detail_search_html", fake_detail_search)

    df = dart_web.list_audit_filings("00126380", "20240101", "20250630")

    assert df.iloc[0].to_dict() == {
        "rcept_no": "20250218800508",
        "report_nm": "감사보고서제출",
        "rcept_dt": "20250218",
        "corp_code": "00126380",
    }
    assert calls == [
        {
            "corp_code": "00126380",
            "bgn_de": "20240101",
            "end_de": "20250630",
            "detail_type": "I001",
            "final": False,
        }
    ]


def test_kr_no_api_key_falls_back_to_public_crawler(monkeypatch) -> None:
    from plugins.kr import dart_web as dart_web_module
    from plugins.kr import name_resolver

    monkeypatch.setattr(name_resolver, "load_settings", lambda: Settings())
    monkeypatch.setattr(
        name_resolver,
        "cached_or_load",
        lambda _key, loader, *, expire: loader(),
    )
    monkeypatch.setattr(
        dart_web_module,
        "resolve_corp",
        lambda _code: {"corp_code": "00126380", "corp_name": "삼성전자"},
    )

    ticker = KRShare().resolve_name("005930")

    assert ticker.code == "005930"
    assert ticker.name == "삼성전자"
    assert ticker.external_id == "00126380"


def test_kr_with_api_key_prefers_openapi_path(monkeypatch) -> None:
    from plugins.kr import dart_web as dart_web_module
    from plugins.kr import name_resolver

    class _Dart:
        corp_codes = pd.DataFrame(
            [{"stock_code": "005930", "corp_code": "00126380", "corp_name": "삼성전자"}]
        )

    monkeypatch.setattr(
        name_resolver,
        "load_settings",
        lambda: Settings.model_validate({"dart": {"api_key": "test-key"}}),
    )
    monkeypatch.setattr(
        name_resolver,
        "cached_or_load",
        lambda _key, loader, *, expire: loader(),
    )
    monkeypatch.setattr(name_resolver, "_dart_for_key", lambda _api_key: _Dart())
    monkeypatch.setattr(
        dart_web_module,
        "resolve_corp",
        lambda _code: (_ for _ in ()).throw(AssertionError("public crawler should not run")),
    )

    ticker = KRShare().resolve_name("005930")

    assert ticker.name == "삼성전자"
    assert ticker.external_id == "00126380"


def test_dart_web_list_ipo_filings_uses_c001_search(monkeypatch) -> None:
    calls = []

    def fake_detail_search(**kwargs):
        calls.append(kwargs)
        return _fixture("dart_ipo_454910.html") if kwargs["detail_type"] == "C001" else ""

    monkeypatch.setattr(dart_web, "_detail_search_html", fake_detail_search)

    df = dart_web.list_ipo_filings("01105153")

    assert "투자설명서" in set(df["report_nm"])
    assert "증권발행실적보고서" in set(df["report_nm"])
    assert set(df.columns) == {"rcept_no", "report_nm", "rcept_dt", "corp_code"}
    assert calls[0]["detail_type"] == "C001"
    assert "" in {call["detail_type"] for call in calls}

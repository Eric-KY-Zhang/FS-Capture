from __future__ import annotations

from types import SimpleNamespace

from app.core.models import Exchange, Ticker
from app.core.settings import Settings
from plugins.jp import edinet_api, name_resolver


def test_jp_resolve_uses_edinet_metadata(monkeypatch) -> None:
    monkeypatch.setattr(
        name_resolver,
        "resolve_one",
        lambda _code: {"edinet_code": "E02144", "filer_name": "トヨタ自動車株式会社"},
    )

    ticker = name_resolver.resolve("7203.T")

    assert ticker.exchange is Exchange.JP
    assert ticker.code == "7203"
    assert ticker.name == "トヨタ自動車株式会社"
    assert ticker.external_id == "E02144"


def test_jp_fetch_company_returns_jpy_and_edinet_extra(monkeypatch) -> None:
    monkeypatch.setattr(
        name_resolver,
        "resolve_one",
        lambda _code: {"edinet_code": "E01777", "jcn": "5010701001000"},
    )

    company = name_resolver.fetch_company(
        Ticker(exchange=Exchange.JP, code="6758", name="ソニーグループ株式会社", external_id="E01777")
    )

    assert company.currency == "JPY"
    assert company.extra["edinet_code"] == "E01777"
    assert company.extra["jcn"] == "5010701001000"


def test_jp_no_key_resolve_uses_public_search_once(monkeypatch) -> None:
    calls = []

    monkeypatch.setattr(name_resolver, "_edinet_api_key", lambda: "")
    monkeypatch.setattr(name_resolver, "_candidate_years", lambda: [2025, 2024, 2023, 2022])
    monkeypatch.setattr(
        name_resolver,
        "cached_or_load",
        lambda _key, loader, *, expire: loader(),
    )
    monkeypatch.setattr(
        "plugins.jp.edinet_web.search_filings_all",
        lambda code: calls.append(code)
        or [
            {
                "doc_id": "S100TR7I",
                "doc_type_code": "120",
                "submit_date_time": "2024-06-25 15:00",
                "edinet_code": "E02144",
                "sec_code": "7203",
                "filer_name": "トヨタ自動車株式会社",
            }
        ],
    )

    ticker = name_resolver.resolve("7203")

    assert ticker.name == "トヨタ自動車株式会社"
    assert ticker.external_id == "E02144"
    assert calls == ["7203"]


def test_edinet_api_sends_subscription_key_and_normalizes_rows(monkeypatch) -> None:
    calls = {}

    class _Client:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def get(self, url, **kwargs):
            calls["url"] = url
            calls["kwargs"] = kwargs
            return SimpleNamespace(
                raise_for_status=lambda: None,
                json=lambda: {
                    "metadata": {"status": "200", "message": "OK"},
                    "results": [
                        {
                            "docID": "S100TEST",
                            "docTypeCode": "120",
                            "submitDateTime": "2024-06-25 15:00",
                            "periodEnd": "2024-03-31",
                            "edinetCode": "E02144",
                            "secCode": "72030",
                            "JCN": "1180301018771",
                            "filerName": "トヨタ自動車株式会社",
                            "docDescription": "有価証券報告書",
                        }
                    ],
                },
            )

    monkeypatch.setattr(edinet_api, "default_client", lambda **_kwargs: _Client())
    monkeypatch.setattr(edinet_api, "cached_or_load", lambda _key, loader, expire: loader())
    monkeypatch.setattr(edinet_api, "_edinet_rate", lambda: 99.0)

    rows = edinet_api.list_documents("2024-06-25", api_key="secret")

    assert calls["kwargs"]["headers"]["Subscription-Key"] == "secret"
    assert calls["kwargs"]["params"]["Subscription-Key"] == "secret"
    assert rows[0]["doc_id"] == "S100TEST"
    assert rows[0]["sec_code"] == "7203"
    assert rows[0]["filer_name"] == "トヨタ自動車株式会社"


def test_edinet_key_setting_reads_environment(monkeypatch) -> None:
    monkeypatch.setenv("EDINET_API_KEY", "env-edinet")

    settings = Settings.model_validate({})

    assert settings.edinet.api_key == "env-edinet"

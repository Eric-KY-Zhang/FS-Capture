from __future__ import annotations

from app.core.models import Exchange, Ticker
from plugins.uk import name_resolver, nsm_web


def test_uk_resolve_alias_accepts_l_suffix() -> None:
    ticker = name_resolver.resolve("ulvr.l")

    assert ticker.exchange is Exchange.UK
    assert ticker.code == "ULVR"
    assert ticker.name == "Unilever PLC"
    assert ticker.external_id == "549300MKFYEKVRWML317"


def test_uk_resolve_fallback_uses_nsm_search(monkeypatch) -> None:
    monkeypatch.setattr(name_resolver, "cached_or_load", lambda _key, loader, expire: loader())
    monkeypatch.setattr(
        "plugins.uk.nsm_web.search_company",
        lambda _query, size=5: [{"company": "Example PLC", "lei": "EXAMPLELEI"}],
    )

    ticker = name_resolver.resolve("exam")

    assert ticker.code == "EXAM"
    assert ticker.name == "Example PLC"
    assert ticker.external_id == "EXAMPLELEI"


def test_uk_fetch_company_returns_gbp_and_lei(monkeypatch) -> None:
    monkeypatch.setattr(
        name_resolver,
        "resolve_one",
        lambda _code: {"name": "AstraZeneca PLC", "lei": "PY6ZZQWO2IZFZC3IOL08"},
    )

    company = name_resolver.fetch_company(
        Ticker(
            exchange=Exchange.UK,
            code="AZN",
            name="AstraZeneca PLC",
            external_id="PY6ZZQWO2IZFZC3IOL08",
        )
    )

    assert company.currency == "GBP"
    assert company.extra["lei"] == "PY6ZZQWO2IZFZC3IOL08"


def test_nsm_search_posts_index_payload_and_normalizes_hits(monkeypatch) -> None:
    calls = {}

    class _Client:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

    def fake_post_json(client, url, **kwargs):
        calls["client"] = client
        calls["url"] = url
        calls["kwargs"] = kwargs
        return {
            "hits": {
                "hits": [
                    {
                        "_id": "NI-1",
                        "_source": {
                            "company": "Unilever PLC",
                            "download_link": "NSM/Portal/report.pdf",
                        },
                    }
                ]
            }
        }

    monkeypatch.setattr(nsm_web, "default_client", lambda **_kwargs: _Client())
    monkeypatch.setattr(nsm_web, "post_json", fake_post_json)
    monkeypatch.setattr(nsm_web, "nsm_rate", lambda: 99.0)

    rows = nsm_web.search_company("Unilever PLC")

    assert calls["url"] == "https://api.data.fca.org.uk/search"
    assert calls["kwargs"]["params"]["index"] == "fca-nsm-searchdata"
    assert calls["kwargs"]["source"] == "nsm"
    assert calls["kwargs"]["rate"] == 99.0
    assert calls["kwargs"]["json_body"]["criteriaObj"]["criteria"][1]["name"] == "company_lei"
    assert rows[0]["_id"] == "NI-1"
    assert rows[0]["download_link"] == "NSM/Portal/report.pdf"

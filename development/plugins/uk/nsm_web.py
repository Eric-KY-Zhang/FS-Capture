"""UK NSM public search adapter."""

from __future__ import annotations

from typing import Any

from app.core.http import default_client, post_json
from app.core.settings import load_settings

_SEARCH_URL = "https://api.data.fca.org.uk/search"
_INDEX = "fca-nsm-searchdata"
_ARTEFACT_BASE = "https://data.fca.org.uk/artefacts/"
_HEADERS = {
    "Content-Type": "application/json",
    "Origin": "https://data.fca.org.uk",
    "Referer": "https://data.fca.org.uk/",
}

_SELECTORS = {
    "hits": "hits.hits",
    "download_link": "download_link",
    "html_link": "html_link",
}


def nsm_rate() -> float:
    return load_settings().rate_limits.nsm


def artifact_url(path: str) -> str:
    text = str(path or "").strip()
    if text.startswith(("http://", "https://")):
        return text
    return f"{_ARTEFACT_BASE}{text.lstrip('/')}"


def _rows(payload: dict[str, Any]) -> list[dict[str, Any]]:
    hits = ((payload.get("hits") or {}).get("hits") or [])
    rows: list[dict[str, Any]] = []
    for hit in hits:
        source = dict(hit.get("_source") or {})
        source["_id"] = hit.get("_id") or source.get("disclosure_id") or source.get("seq_id")
        rows.append(source)
    return rows


def search(
    criteria: list[dict[str, Any]],
    *,
    date_criteria: list[dict[str, Any]] | None = None,
    size: int = 50,
    from_: int = 0,
    sort: str = "submitted_date",
    sortorder: str = "desc",
) -> list[dict[str, Any]]:
    payload = {
        "from": from_,
        "size": size,
        "sort": sort,
        "sortorder": sortorder,
        "criteriaObj": {
            "criteria": criteria,
            "dateCriteria": date_criteria or [],
        },
    }
    with default_client(source="nsm", timeout=60.0) as client:
        data = post_json(
            client,
            _SEARCH_URL,
            source="nsm",
            rate=nsm_rate(),
            json_body=payload,
            params={"index": _INDEX},
            headers=_HEADERS,
        )
    return _rows(data)


def search_company(query: str, *, size: int = 5) -> list[dict[str, Any]]:
    return search(
        [
            {"name": "latest_flag", "value": "Y"},
            {"name": "company_lei", "value": [query, "", "disclose_org", "related_org"]},
        ],
        size=size,
    )


def search_filings(company: str, lei: str, year: int, headline: str) -> list[dict[str, Any]]:
    search_text = company or lei
    if not search_text:
        return []
    criteria = [
        {"name": "latest_flag", "value": "Y"},
        {"name": "headline", "value": headline},
        {"name": "company_lei", "value": [search_text, "", "disclose_org", "related_org"]},
    ]
    date_criteria = [
        {
            "name": "publication_date",
            "value": {
                "from": f"{year}-01-01T00:00:00Z",
                "to": f"{year}-12-31T23:59:59Z",
            },
        }
    ]
    rows = search(criteria, date_criteria=date_criteria)
    if rows or not lei or lei == search_text:
        return rows
    criteria[-1] = {"name": "company_lei", "value": [lei, "", "disclose_org", "related_org"]}
    return search(criteria, date_criteria=date_criteria)

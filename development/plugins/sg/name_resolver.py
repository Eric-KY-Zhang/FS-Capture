"""SG ticker resolution via SGXNet announcements."""

from __future__ import annotations

import datetime as dt
import re
from typing import Any

from app.core.cache import cached_or_load
from app.core.models import Company, Exchange, Ticker

_TTL = 24 * 3600


def _normalize_code(code: str) -> str:
    normalized = code.strip().upper().removesuffix(".SI")
    if not re.fullmatch(r"[A-Z0-9]{1,5}", normalized):
        raise ValueError(f"新股代码 {code} 格式错误（应为 SGX ticker，如 D05 或 D05.SI）")
    return normalized


def _issuer_from_rows(code: str, rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    for row in rows:
        for issuer in row.get("issuers") or []:
            if str(issuer.get("stock_code") or "").upper() != code:
                continue
            return {
                "name": str(
                    issuer.get("issuer_name") or issuer.get("security_name") or code
                ).strip(),
                "security_name": str(issuer.get("security_name") or "").strip(),
                "isin_code": str(issuer.get("isin_code") or "").strip(),
                "ibm_code": str(issuer.get("ibm_code") or "").strip(),
            }
    return None


def _ipo_from_rows(code: str, rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    for row in rows:
        if str(row.get("id") or "").strip().upper() != code:
            continue
        name = str(row.get("name") or "").strip()
        if not name:
            return None
        return {
            "name": name,
            "security_name": name,
            "isin_code": "",
            "ibm_code": "",
            "ipo_id": code,
        }
    return None


def resolve_one(code: str) -> dict[str, Any] | None:
    cache_key = f"sg:ticker:{code}:v1"

    def _fetch() -> dict[str, Any] | None:
        from .sgxnet_web import list_ipo_prospectuses, search_announcements

        today = dt.date.today()
        rows = search_announcements(
            code,
            start_year=today.year - 3,
            end_year=today.year + 1,
            page_size=20,
            max_pages=1,
        )
        issuer = _issuer_from_rows(code, rows)
        if issuer:
            return issuer
        return _ipo_from_rows(code, list_ipo_prospectuses())

    return cached_or_load(cache_key, _fetch, expire=_TTL)


def resolve(code: str) -> Ticker:
    normalized = _normalize_code(code)
    info = resolve_one(normalized)
    if not info:
        raise ValueError(f"未找到新股代码 {normalized}（请确认 SGX ticker）")
    return Ticker(
        exchange=Exchange.SG,
        code=normalized,
        name=str(info.get("name") or normalized),
        external_id=str(info.get("isin_code") or ""),
    )


def fetch_company(ticker: Ticker) -> Company:
    info = resolve_one(ticker.code) or {}
    return Company(
        ticker=ticker,
        listing_date=None,
        industry=None,
        currency="SGD",
        extra={
            "isin_code": ticker.external_id or info.get("isin_code") or "",
            "security_name": info.get("security_name") or ticker.name or "",
            "ibm_code": info.get("ibm_code") or "",
            "ipo_id": info.get("ipo_id") or "",
        },
    )

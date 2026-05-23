"""UK ticker resolution via FCA NSM issuer metadata."""

from __future__ import annotations

import re
from typing import Any

from app.core.cache import cached_or_load
from app.core.models import Company, Exchange, Ticker

_TTL = 7 * 24 * 3600

_ALIASES = {
    "ULVR": {"name": "Unilever PLC", "lei": "549300MKFYEKVRWML317"},
    "HSBA": {"name": "HSBC Holdings PLC", "lei": "MLU0ZO3ML4LN2LL2TL39"},
    "AZN": {"name": "AstraZeneca PLC", "lei": "PY6ZZQWO2IZFZC3IOL08"},
}


def _normalize_code(code: str) -> str:
    c = code.strip().upper()
    c = re.sub(r"\.L$", "", c)
    if not re.fullmatch(r"[A-Z]{1,5}", c):
        raise ValueError(f"英股代码 {code} 格式错误（应为 LSE ticker，如 ULVR 或 ULVR.L）")
    return c


def resolve_one(code: str) -> dict[str, Any] | None:
    if code in _ALIASES:
        return dict(_ALIASES[code])

    cache_key = f"uk:ticker:{code}:v1"

    def _fetch() -> dict[str, Any] | None:
        from .nsm_web import search_company

        rows = search_company(code, size=5)
        for row in rows:
            company = str(row.get("company") or "").strip()
            lei = str(row.get("lei") or "").strip()
            if company or lei:
                return {"name": company or code, "lei": lei}
        return None

    return cached_or_load(cache_key, _fetch, expire=_TTL)


def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)
    info = resolve_one(norm)
    if not info:
        raise ValueError(f"未找到英股代码 {code}（请确认 LSE ticker 或使用批量导入中的代码格式）")
    return Ticker(
        exchange=Exchange.UK,
        code=norm,
        name=str(info.get("name") or norm),
        external_id=str(info.get("lei") or ""),
    )


def fetch_company(ticker: Ticker) -> Company:
    info = resolve_one(ticker.code) or {}
    return Company(
        ticker=ticker,
        listing_date=None,
        industry=None,
        currency="GBP",
        extra={
            "lei": ticker.external_id or info.get("lei") or "",
            "nsm_company": ticker.name or info.get("name") or "",
        },
    )

"""JP ticker resolution via EDINET document metadata."""

from __future__ import annotations

import datetime as dt
import re
from functools import lru_cache
from typing import Any

from app.core.cache import cached_or_load
from app.core.models import Company, Exchange, Ticker
from app.core.settings import load_settings

_TTL = 7 * 24 * 3600


def _normalize_code(code: str) -> str:
    c = code.strip().upper()
    c = re.sub(r"^(JP)", "", c)
    c = re.sub(r"\.(T|JP)$", "", c)
    if not re.fullmatch(r"\d{1,4}", c):
        raise ValueError(f"日股代码 {code} 格式错误（应为 4 位数字，如 7203）")
    return c.zfill(4)


@lru_cache(maxsize=8)
def _edinet_api_key() -> str:
    return (load_settings().edinet.api_key or "").strip()


def reset_edinet_client() -> None:
    _edinet_api_key.cache_clear()


def _candidate_dates(today: dt.date | None = None) -> list[str]:
    ref = today or dt.date.today()
    dates: list[str] = []
    for year in range(ref.year, ref.year - 4, -1):
        start = dt.date(year, 6, 1)
        end = dt.date(year, 7, 15)
        current = start
        while current <= end:
            dates.append(current.isoformat())
            current += dt.timedelta(days=1)
    return dates


def _candidate_years(today: dt.date | None = None) -> list[int]:
    ref = today or dt.date.today()
    latest_filing_year = ref.year if ref >= dt.date(ref.year, 6, 1) else ref.year - 1
    return list(range(latest_filing_year, latest_filing_year - 4, -1))


def _list_documents(submit_date: str) -> list[dict[str, Any]]:
    api_key = _edinet_api_key()
    if api_key:
        from .edinet_api import list_documents

        return list_documents(submit_date, api_key=api_key)

    from .edinet_web import list_documents

    return list_documents(submit_date)


def _resolve_one_via_api(sec_code: str) -> dict[str, Any] | None:
    fallback: dict[str, Any] | None = None
    for submit_date in _candidate_dates():
        for row in _list_documents(submit_date):
            if row.get("sec_code") != sec_code:
                continue
            if fallback is None:
                fallback = row
            if str(row.get("doc_type_code")) == "120":
                return row
    return fallback


def _resolve_one_via_public(sec_code: str) -> dict[str, Any] | None:
    from .edinet_web import search_filings_all

    fallback: dict[str, Any] | None = None
    rows = search_filings_all(sec_code)
    for year in _candidate_years():
        year_rows = [
            row
            for row in rows
            if row.get("sec_code") == sec_code
            and str(row.get("submit_date_time") or "").startswith(str(year))
        ]
        for row in year_rows:
            if fallback is None:
                fallback = row
            if str(row.get("doc_type_code")) == "120":
                return row
    return fallback


def resolve_one(sec_code: str) -> dict[str, Any] | None:
    api_key = _edinet_api_key()
    mode = "api" if api_key else "web"
    cache_key = f"jp:sec-code:{sec_code}:{mode}:v2"

    def _fetch() -> dict[str, Any] | None:
        if api_key:
            return _resolve_one_via_api(sec_code)
        return _resolve_one_via_public(sec_code)

    return cached_or_load(cache_key, _fetch, expire=_TTL)


def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)
    info = resolve_one(norm)
    if not info:
        raise ValueError(
            f"未找到日股代码 {code}（请确认代码格式，或稍后重试 EDINET 公网页）"
        )
    return Ticker(
        exchange=Exchange.JP,
        code=norm,
        name=str(info.get("filer_name") or norm),
        external_id=str(info.get("edinet_code") or ""),
    )


def fetch_company(ticker: Ticker) -> Company:
    info = resolve_one(ticker.code) or {}
    extra = {
        "edinet_code": ticker.external_id or info.get("edinet_code") or "",
        "sec_code": ticker.code,
        "jcn": info.get("jcn") or "",
    }
    return Company(
        ticker=ticker,
        listing_date=None,
        industry=None,
        currency="JPY",
        extra=extra,
    )

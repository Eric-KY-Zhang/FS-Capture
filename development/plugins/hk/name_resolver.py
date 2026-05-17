"""HK ticker resolution via HKEXnews.

HKEX Title Search does not accept the five-digit stock code as ``stockId``.
It requires an internal numeric ``stockId`` returned by the HKEXnews prefix
JSONP endpoint.  Keep the user-facing ticker code as five digits and store the
real HKEX ``stockId`` in ``ticker.external_id``.
"""

from __future__ import annotations

import json
import re

from loguru import logger

from app.core.cache import get_cache
from app.core.http import default_client
from app.core.models import Company, Exchange, Ticker
from app.core.ratelimit import limiter

_PREFIX_URL = "https://www1.hkexnews.hk/search/prefix.do"
_CACHE_KEY_PREFIX = "hk:hkex_prefix:v3:"
_TTL = 24 * 3600
_JSONP_RE = re.compile(r"^[^(]*\((.*)\)\s*;?\s*$", re.DOTALL)


def _normalize_code(code: str) -> str:
    c = code.strip().upper()
    if c.endswith(".HK"):
        c = c[:-3]
    if c.startswith("HK"):
        c = c[2:]
    c = re.sub(r"\D", "", c)
    if not c:
        raise ValueError(f"Invalid HK stock code: {code}")
    return c.zfill(5)


def _parse_jsonp(text: str) -> dict:
    match = _JSONP_RE.match(text.strip())
    if not match:
        raise ValueError("HKEX prefix response is not JSONP")
    return json.loads(match.group(1))


def _cached_result(norm: str) -> dict | None:
    cached = get_cache().get(f"{_CACHE_KEY_PREFIX}{norm}")
    return cached if isinstance(cached, dict) else None


def _store_cache(norm: str, value: dict) -> None:
    get_cache().set(f"{_CACHE_KEY_PREFIX}{norm}", value, expire=_TTL)


def _fetch_hkex_prefix(norm: str) -> dict | None:
    params = {
        "callback": "callback",
        "lang": "EN",
        "type": "A",
        "name": norm,
        "market": "SEHK",
    }
    headers = {
        "Referer": "https://www1.hkexnews.hk/search/titlesearch.xhtml?lang=en",
        "Accept": "application/javascript,text/javascript,*/*;q=0.1",
    }

    try:
        with default_client(source="hkexnews") as client:
            limiter("hkexnews", 3.0).acquire_blocking()
            response = client.get(_PREFIX_URL, params=params, headers=headers)
            response.raise_for_status()
            payload = _parse_jsonp(response.text)
    except Exception as exc:
        logger.warning(f"HKEX prefix lookup failed for {norm}: {exc}")
        return None

    for item in payload.get("stockInfo") or []:
        code = str(item.get("code") or "").strip().zfill(5)
        if code != norm:
            continue
        stock_id = item.get("stockId")
        if stock_id is None or str(stock_id).strip() == "":
            continue
        return {
            "code": code,
            "name": str(item.get("name") or code).strip() or code,
            "stockId": str(stock_id).strip(),
        }
    return None


def _fetch_chinese_short_name(norm: str) -> str | None:
    try:
        import akshare as ak

        df = ak.stock_hk_security_profile_em(symbol=norm)
    except Exception as exc:
        logger.warning(f"Eastmoney HK security profile lookup failed for {norm}: {exc}")
        return None

    if df.empty:
        return None
    for col in ("证券简称", "股票简称", "简称", "名称"):
        if col not in df.columns:
            continue
        value = str(df.iloc[0][col] or "").strip()
        if value:
            return value
    return None


def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)
    result = _cached_result(norm)
    if not result:
        result = _fetch_hkex_prefix(norm)
        if result:
            zh_name = _fetch_chinese_short_name(norm)
            if zh_name:
                result["zh_name"] = zh_name
            _store_cache(norm, result)
    elif not result.get("zh_name"):
        zh_name = _fetch_chinese_short_name(norm)
        if zh_name:
            result = dict(result)
            result["zh_name"] = zh_name
            _store_cache(norm, result)

    if not result:
        raise ValueError(f"Unable to resolve HKEX stockId for HK stock code {norm}")

    return Ticker(
        exchange=Exchange.HK,
        code=norm,
        name=str(result.get("zh_name") or result.get("name") or norm),
        external_id=str(result["stockId"]),
    )


def fetch_company(ticker: Ticker) -> Company:
    industry: str | None = None
    extra: dict = {}
    try:
        import akshare as ak

        df = ak.stock_hk_company_profile_em(symbol=ticker.code)
        if not df.empty:
            if {"item", "value"}.issubset(df.columns):
                d = dict(zip(df["item"].astype(str), df["value"].astype(str), strict=False))
            else:
                d = {str(k): str(v) for k, v in df.iloc[0].dropna().items()}
            industry = d.get("industry") or d.get("Industry") or d.get("所属行业")
            extra = d
    except Exception as exc:
        logger.warning(f"akshare HK company profile failed for {ticker.code}: {exc}")

    return Company(
        ticker=ticker,
        listing_date=None,
        industry=industry,
        currency="HKD",
        extra=extra,
    )

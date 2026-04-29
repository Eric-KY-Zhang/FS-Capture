"""A-share name resolution + company metadata.

Strategy:
- For ticker -> name: use akshare `stock_info_a_code_name()` (cached daily).
  Covers Shanghai (600/601/603/688), Shenzhen (000/001/002/300/301), Beijing (8/4).
- For ticker -> cninfo orgId: hit cninfo's topSearch query endpoint.
  orgId format: e.g. "gssh0600519" (sh) or "gssz0000001" (sz). Required for the
  hisAnnouncement download API.
"""
from __future__ import annotations

import time
from typing import Optional

from loguru import logger

from app.core.cache import get_cache
from app.core.http import default_client, get_json, post_json
from app.core.models import Company, Exchange, Ticker


_CNINFO_TOPSEARCH = "http://www.cninfo.com.cn/new/information/topSearch/detailOfQuery"
_CACHE_KEY_NAME_MAP = "ashare:code_name_map:v1"
_CACHE_KEY_ORGID = "ashare:orgid:"
_NAME_MAP_TTL = 24 * 3600
_ORGID_TTL = 30 * 24 * 3600


def _normalize(code: str) -> str:
    """User may enter 600519, sh600519, SH600519, 600519.SH, etc. -> 600519"""
    c = code.strip().upper()
    for sep in (".", " "):
        if sep in c:
            c = c.split(sep)[0]
    if c.startswith(("SH", "SZ", "BJ")):
        c = c[2:]
    return c


def _market_prefix(code: str) -> str:
    """Determine 'sh' / 'sz' / 'bj' from numeric code."""
    if code.startswith(("60", "688", "900")):
        return "sh"
    if code.startswith(("000", "001", "002", "200", "300", "301")):
        return "sz"
    if code.startswith(("4", "8", "9")):  # BJ
        return "bj"
    # Fallback
    return "sh"


def _load_name_map() -> dict[str, str]:
    """code -> name, cached daily via akshare."""
    cache = get_cache()
    cached = cache.get(_CACHE_KEY_NAME_MAP)
    if cached:
        return cached  # type: ignore[return-value]

    import akshare as ak  # local import — heavy module
    logger.info("Loading A-share code↔name map from akshare ...")
    df = ak.stock_info_a_code_name()
    name_map = dict(zip(df["code"].astype(str), df["name"].astype(str)))
    cache.set(_CACHE_KEY_NAME_MAP, name_map, expire=_NAME_MAP_TTL)
    return name_map


def _fetch_orgid(code: str) -> Optional[str]:
    cache = get_cache()
    key = _CACHE_KEY_ORGID + code
    cached = cache.get(key)
    if cached:
        return cached  # type: ignore[return-value]

    headers = {
        "Origin": "http://www.cninfo.com.cn",
        "Referer": f"http://www.cninfo.com.cn/new/disclosure/stock?stockCode={code}",
        "Accept": "application/json, text/plain, */*",
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    }
    data = {"keyWord": code, "maxSecNum": 10, "maxListNum": 5, "_": int(time.time() * 1000)}

    with default_client(source="cninfo") as client:
        try:
            payload = post_json(
                client, _CNINFO_TOPSEARCH, source="cninfo", rate=5.0,
                data=data, headers=headers,
            )
        except Exception as exc:
            logger.warning(f"cninfo topSearch failed for {code}: {exc}")
            return None

    items = payload.get("keyBoardList") or []
    for item in items:
        if str(item.get("code")) == code:
            org_id = item.get("orgId")
            if org_id:
                cache.set(key, org_id, expire=_ORGID_TTL)
                return org_id
    return None


def resolve(code: str) -> Ticker:
    norm = _normalize(code)
    name_map = _load_name_map()
    name = name_map.get(norm)
    if not name:
        raise ValueError(f"A股代码 {code} 未找到。请确认代码格式（6位数字）")

    org_id = _fetch_orgid(norm)
    market = _market_prefix(norm)
    # Always store something in external_id; orgId is preferred but the downloader
    # has a fallback if it's None.
    ext = org_id or f"{market.upper()}{norm}"
    return Ticker(exchange=Exchange.A_SHARE, code=norm, name=name, external_id=ext)


def fetch_company(ticker: Ticker) -> Company:
    """Best-effort metadata via akshare. Fields filled where available."""
    market = _market_prefix(ticker.code)
    sym = f"{market}{ticker.code}"
    industry: Optional[str] = None
    listing_date: Optional[str] = None
    extra: dict = {}

    try:
        import akshare as ak
        df = ak.stock_individual_info_em(symbol=ticker.code)
        d = dict(zip(df["item"].astype(str), df["value"].astype(str)))
        industry = d.get("行业")
        listing_date = d.get("上市时间")
        extra = {k: v for k, v in d.items() if k not in {"行业", "上市时间"}}
    except Exception as exc:
        logger.warning(f"akshare stock_individual_info_em failed for {ticker.code}: {exc}")

    return Company(
        ticker=ticker,
        listing_date=listing_date,
        industry=industry,
        currency="CNY",
        extra=extra,
    )

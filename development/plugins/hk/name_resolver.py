"""HK ticker resolution.

Primary: per-code lookup via Eastmoney `stock/get` endpoint (lightweight, reliable).
Fallback: akshare `stock_hk_spot_em()` bulk list (only if user wants name preview
without exact code).
"""
from __future__ import annotations

from typing import Optional

from loguru import logger

from app.core.cache import get_cache
from app.core.http import default_client, get_json
from app.core.models import Company, Exchange, Ticker


_CACHE_KEY = "hk:code_name_map:v1"
_TTL = 24 * 3600


def _normalize_code(code: str) -> str:
    c = code.strip().upper()
    if c.endswith(".HK"):
        c = c[:-3]
    c = c.lstrip("HK")
    c = c.lstrip("0") or "0"
    return c.zfill(5)


def _eastmoney_single(code: str) -> Optional[str]:
    """Direct lookup via Eastmoney stock/get. Returns name or None.
    secid format: 116.{code} for HK.
    """
    url = "https://push2delay.eastmoney.com/api/qt/stock/get"
    params = {
        "secid": f"116.{code}",
        "fields": "f57,f58",
        "_": "0",
    }
    headers = {
        "Referer": "https://quote.eastmoney.com/",
        "Accept": "application/json, text/plain, */*",
    }
    try:
        with default_client(source="eastmoney") as client:
            payload = get_json(
                client, url, source="eastmoney", rate=4.0,
                params=params, headers=headers,
            )
        data = payload.get("data") or {}
        name = data.get("f58")
        return str(name).strip() if name else None
    except Exception as exc:
        logger.warning(f"Eastmoney HK single-stock lookup failed for {code}: {exc}")
        return None


def _financial_report_name(code: str) -> Optional[str]:
    try:
        import akshare as ak
        df = ak.stock_financial_hk_report_em(stock=code, symbol="资产负债表", indicator="年度")
        if not df.empty and "SECURITY_NAME_ABBR" in df.columns:
            name = df["SECURITY_NAME_ABBR"].dropna().astype(str).iloc[0]
            return name.strip() if name else None
    except Exception as exc:
        logger.warning(f"akshare HK financial-name lookup failed for {code}: {exc}")
    return None


def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)

    # Try cache first (the bulk map populated by previous successful akshare runs).
    cache = get_cache()
    bulk: dict[str, str] = cache.get(_CACHE_KEY) or {}
    name = bulk.get(norm)

    if not name:
        name = _eastmoney_single(norm)

    if not name:
        name = _financial_report_name(norm)

    if not name:
        logger.warning(f"未能识别港股 {code} 名称，将继续使用代码作为名称")
        name = norm

    return Ticker(exchange=Exchange.HK, code=norm, name=name, external_id=norm)


def fetch_company(ticker: Ticker) -> Company:
    industry: Optional[str] = None
    extra: dict = {}
    try:
        import akshare as ak
        df = ak.stock_hk_company_profile_em(symbol=ticker.code)
        if not df.empty:
            if {"item", "value"}.issubset(df.columns):
                d = dict(zip(df["item"].astype(str), df["value"].astype(str)))
            else:
                d = {str(k): str(v) for k, v in df.iloc[0].dropna().items()}
            industry = d.get("行业") or d.get("所属行业")
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

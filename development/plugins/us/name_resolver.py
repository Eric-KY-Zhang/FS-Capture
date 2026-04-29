"""US ticker resolution via SEC EDGAR company_tickers.json (cached daily)."""
from __future__ import annotations

from typing import Optional

from loguru import logger

from app.core.cache import get_cache
from app.core.http import default_client, get_json
from app.core.models import Company, Exchange, Ticker


_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
_CACHE_KEY = "us:ticker_map:v1"
_TTL = 24 * 3600


def _load_map() -> dict[str, dict]:
    """ticker -> {cik, name}"""
    cache = get_cache()
    cached = cache.get(_CACHE_KEY)
    if cached:
        return cached  # type: ignore[return-value]

    logger.info("Loading SEC ticker↔CIK map ...")
    with default_client(source="sec") as client:
        payload = get_json(client, _TICKERS_URL, source="sec", rate=8.0)

    out: dict[str, dict] = {}
    # company_tickers.json is keyed by row index, values have ticker / cik_str / title
    for _, row in (payload or {}).items():
        ticker = str(row.get("ticker", "")).upper()
        if not ticker:
            continue
        out[ticker] = {
            "cik": int(row["cik_str"]),
            "name": str(row.get("title", "")),
        }
    cache.set(_CACHE_KEY, out, expire=_TTL)
    return out


def resolve(code: str) -> Ticker:
    norm = code.strip().upper()
    m = _load_map()
    info = m.get(norm)
    if not info:
        # Fuzzy match: maybe user entered with class suffix (e.g. BRK.B -> BRK-B in SEC)
        alt = norm.replace(".", "-")
        info = m.get(alt)
        if info:
            norm = alt
    if not info:
        raise ValueError(f"未找到美股代码 {code}（请确认代码与 SEC EDGAR 一致）")
    cik = f"{info['cik']:010d}"
    return Ticker(exchange=Exchange.US, code=norm, name=info["name"], external_id=cik)


def fetch_company(ticker: Ticker) -> Company:
    cik = ticker.external_id or ""
    listing_date: Optional[str] = None
    industry: Optional[str] = None
    extra: dict = {}

    if cik:
        url = f"https://data.sec.gov/submissions/CIK{cik}.json"
        with default_client(source="sec") as client:
            try:
                payload = get_json(client, url, source="sec", rate=8.0)
                industry = payload.get("sicDescription")
                listing_date = payload.get("ein")  # placeholder; SEC doesn't expose listing date directly
                extra = {
                    "exchange": payload.get("exchanges", []),
                    "sic": payload.get("sic"),
                    "fiscalYearEnd": payload.get("fiscalYearEnd"),
                    "category": payload.get("category"),
                    "stateOfIncorporation": payload.get("stateOfIncorporation"),
                }
            except Exception as exc:
                logger.warning(f"SEC submissions fetch failed for {ticker.code}: {exc}")

    return Company(
        ticker=ticker,
        listing_date=None,
        industry=industry,
        currency="USD",
        extra=extra,
    )

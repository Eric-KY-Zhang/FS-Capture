"""KR ticker resolution via DART OpenAPI or public DART disclosure pages."""

from __future__ import annotations

from functools import lru_cache

from loguru import logger

from app.core.cache import get_cache
from app.core.models import Company, Exchange, Ticker
from app.core.settings import load_settings

_CACHE_KEY_CORP_MAP = "kr:corpcode_map:v1"
_TTL = 7 * 24 * 3600


def _normalize_code(code: str) -> str:
    c = code.strip().upper()
    if c.endswith(".KS") or c.endswith(".KQ"):
        c = c[:-3]
    return c.zfill(6)


@lru_cache(maxsize=4)
def _dart_for_key(api_key: str):
    import OpenDartReader

    return OpenDartReader(api_key)


def _dart():
    api_key = (load_settings().dart.api_key or "").strip()
    if not api_key:
        return None
    return _dart_for_key(api_key)


def reset_dart_client() -> None:
    _dart_for_key.cache_clear()


def resolve_one(stock_code: str) -> dict | None:
    """stock_code -> {corp_code, corp_name}; cache permanently."""
    cache = get_cache()
    cache_key = f"kr:corp:{stock_code}"
    cached = cache.get(cache_key)
    if cached:
        return cached  # type: ignore[return-value]

    dart = _dart()
    if dart is not None:
        try:
            df = dart.corp_codes
            df = df[df["stock_code"].astype(str).str.zfill(6) == stock_code]
            if not df.empty:
                row = df.iloc[0]
                info = {
                    "corp_code": str(row["corp_code"]),
                    "corp_name": str(row["corp_name"]),
                }
                cache.set(cache_key, info)
                return info
        except Exception as exc:
            logger.warning(f"DART OpenAPI corp lookup failed for {stock_code}: {exc}")

    from .dart_web import resolve_corp

    logger.info(f"DART OpenAPI not configured; falling back to public crawler for {stock_code}")
    info = resolve_corp(stock_code)
    if info is not None:
        cache.set(cache_key, info)
    return info


def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)
    info = resolve_one(norm)
    if not info:
        raise ValueError(f"未找到韩股代码 {code}（请确认代码格式，如 005930）")
    return Ticker(
        exchange=Exchange.KR,
        code=norm,
        name=info["corp_name"],
        external_id=info["corp_code"],
    )


def fetch_company(ticker: Ticker) -> Company:
    industry: str | None = None
    extra: dict = {}
    dart = _dart()
    if dart is not None:
        try:
            info_df = dart.company(corp=ticker.external_id or ticker.code)
            if info_df is not None and not info_df.empty:
                row = info_df.iloc[0]
                # DART OpenAPI spells this field as "induty_code"; keep the upstream name.
                industry_value = row.get("induty_code", row.get("industry_code", ""))
                industry = str(industry_value) or None
                extra = {k: str(v) for k, v in row.items()}
        except Exception as exc:
            logger.warning(f"DART company info failed for {ticker.code}: {exc}")

    return Company(
        ticker=ticker,
        listing_date=None,
        industry=industry,
        currency="KRW",
        extra=extra,
    )

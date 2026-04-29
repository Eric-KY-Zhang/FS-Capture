"""KR ticker resolution via DART OpenAPI corpCode.

DART OpenAPI requires a free API key (set in Settings → DART API key, persisted
in config.toml). The corpCode endpoint returns a ZIP containing all listed
companies' corp_code, corp_name, stock_code, modify_date.

We rely on OpenDartReader to manage the ZIP/XML extraction.
"""
from __future__ import annotations

from functools import lru_cache
from typing import Optional

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


@lru_cache(maxsize=1)
def _dart():
    s = load_settings()
    if not s.dart.api_key:
        raise ValueError(
            "尚未配置 DART API 密钥。韩股官方披露数据需要 DART OpenAPI，"
            "请在『设置』中填入密钥；如暂时无法注册，可先取消勾选韩股。"
        )
    import OpenDartReader
    return OpenDartReader(s.dart.api_key)


def _load_map() -> dict[str, dict]:
    """stock_code -> {corp_code, corp_name}"""
    cache = get_cache()
    cached = cache.get(_CACHE_KEY_CORP_MAP)
    if cached:
        return cached  # type: ignore[return-value]

    logger.info("Loading DART corpCode map ...")
    dart = _dart()
    df = dart.corp_codes  # full DataFrame; columns include stock_code, corp_code, corp_name
    df = df[df["stock_code"].astype(str).str.strip() != ""]
    out: dict[str, dict] = {}
    for _, r in df.iterrows():
        sc = str(r["stock_code"]).strip().zfill(6)
        out[sc] = {
            "corp_code": str(r["corp_code"]),
            "corp_name": str(r["corp_name"]),
        }
    cache.set(_CACHE_KEY_CORP_MAP, out, expire=_TTL)
    return out


def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)
    m = _load_map()
    info = m.get(norm)
    if not info:
        raise ValueError(f"未找到韩股代码 {code}（请确认代码格式，如 005930）")
    return Ticker(
        exchange=Exchange.KR,
        code=norm,
        name=info["corp_name"],
        external_id=info["corp_code"],
    )


def fetch_company(ticker: Ticker) -> Company:
    industry: Optional[str] = None
    extra: dict = {}
    try:
        dart = _dart()
        info_df = dart.company(corp=ticker.external_id or ticker.code)
        if info_df is not None and not info_df.empty:
            row = info_df.iloc[0]
            industry = str(row.get("induty_code", "")) or None
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

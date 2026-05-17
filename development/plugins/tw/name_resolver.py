"""TW (台股) ticker resolution via TWSE ISIN service.

Strategy:
- Download the company directory from https://isin.twse.com.tw/isin/C_public.jsp
  - strMode=2 → 上市 (TWSE main board)
  - strMode=4 → 上柜 (TPEx OTC)
- Parse the HTML table for stock_code, company_name, listing_date, industry.
- Cache the merged map in diskcache for 30 days.
- Resolve a 4-digit code (e.g. ``2330``) to ticker / company metadata.

No authentication required. Both TWSE and TPEx issuers file their disclosures
into the same MOPS portal, so a single plugin covers both boards.
"""

from __future__ import annotations

import re

from loguru import logger

from app.core.cache import get_cache
from app.core.http import default_client
from app.core.models import Company, Exchange, Ticker
from app.core.ratelimit import limiter

_ISIN_URL = "https://isin.twse.com.tw/isin/C_public.jsp?strMode={mode}"
_ISIN_MODES = {"sii": 2, "otc": 4}
_CACHE_KEY = "tw:isin_map:v1"
_TTL = 30 * 24 * 3600
_RATE = 2.0


def _normalize_code(code: str) -> str:
    """User may enter `2330`, `tw2330`, `2330.TW`, `2330.TWO` → `2330`."""
    c = code.strip().upper()
    if "." in c:
        c = c.split(".")[0]
    if c.startswith("TW"):
        c = c[2:]
    return c


def _parse_isin_html(html: str, board: str) -> list[dict]:
    """Parse the ISIN HTML table.

    Each data row begins with `<tr><td ...>{code}　{name}</td>...` (full-width
    space between code and name). The table also includes section dividers like
    "股票" / "ETF" / "存託憑證" — we only keep rows whose first cell starts with
    a 4-digit stock code.
    """
    rows: list[dict] = []
    # Row pattern: capture every <tr>...</tr>; we'll regex columns inside.
    tr_re = re.compile(r"<tr[^>]*>(.*?)</tr>", re.IGNORECASE | re.DOTALL)
    td_re = re.compile(r"<td[^>]*>(.*?)</td>", re.IGNORECASE | re.DOTALL)
    tag_re = re.compile(r"<[^>]+>")

    for match in tr_re.finditer(html):
        cells_html = td_re.findall(match.group(1))
        cells = [tag_re.sub("", c).replace(" ", " ").strip() for c in cells_html]
        if len(cells) < 7:
            continue
        first = cells[0].replace("　", " ").strip()
        # First cell looks like: "2330　台積電" (code, full-width space, name)
        m = re.match(r"^(\d{4,6})\s+(.+)$", first)
        if not m:
            continue
        code = m.group(1)
        name = m.group(2).strip()
        # Skip non-common-stock instruments (warrants / DRs / ETF) — keep only
        # 4-digit numeric codes which on TW are common shares.
        if not re.fullmatch(r"\d{4}", code):
            continue
        # cells layout: [code+name, ISIN, 上市日, 市場別, 產業別, CFICODE, 備註]
        listing_date = cells[2] if len(cells) > 2 else ""
        industry = cells[4] if len(cells) > 4 else ""
        rows.append(
            {
                "code": code,
                "name": name,
                "listing_date": listing_date,
                "industry": industry,
                "board": board,
            }
        )
    return rows


def _fetch_isin_board(board: str) -> list[dict]:
    mode = _ISIN_MODES[board]
    url = _ISIN_URL.format(mode=mode)
    with default_client(source="twse", timeout=60.0) as client:
        limiter("twse", _RATE).acquire_blocking()
        resp = client.get(url)
        resp.raise_for_status()
        # TWSE ISIN page is Big5 — try declared encoding then fall back.
        try:
            resp.encoding = resp.encoding or "big5"
            html = resp.text
            if "台" not in html and "股票" not in html:
                resp.encoding = "big5"
                html = resp.text
        except Exception:
            html = resp.content.decode("big5", errors="replace")
        return _parse_isin_html(html, board)


def _load_map() -> dict[str, dict]:
    """Return ``{code: {name, listing_date, industry, board}}``.

    Merges 上市 + 上柜; 上市 wins on conflicts (rare).
    """
    cache = get_cache()
    cached = cache.get(_CACHE_KEY)
    if cached:
        return cached  # type: ignore[return-value]

    logger.info("Loading TW ISIN directory (上市 + 上櫃) ...")
    out: dict[str, dict] = {}
    for board in ("otc", "sii"):  # sii second so its entries overwrite otc on conflict
        try:
            rows = _fetch_isin_board(board)
        except Exception as exc:
            logger.warning(f"TW ISIN fetch failed for board={board}: {exc}")
            continue
        for row in rows:
            out[row["code"]] = {
                "name": row["name"],
                "listing_date": row["listing_date"] or None,
                "industry": row["industry"] or None,
                "board": row["board"],
            }
    if not out:
        raise RuntimeError("无法加载台股 ISIN 列表（TWSE/TPEx 均不可达）")
    cache.set(_CACHE_KEY, out, expire=_TTL)
    return out


def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)
    if not re.fullmatch(r"\d{4}", norm):
        raise ValueError(f"台股代码 {code} 格式错误（应为 4 位数字，如 2330）")
    m = _load_map()
    info = m.get(norm)
    if not info:
        raise ValueError(f"未找到台股代码 {code}（上市 / 上柜均未匹配）")
    return Ticker(
        exchange=Exchange.TW,
        code=norm,
        name=info["name"],
        external_id=norm,  # MOPS uses the 4-digit code directly as `co_id`
    )


def fetch_company(ticker: Ticker) -> Company:
    info = _load_map().get(ticker.code, {})
    return Company(
        ticker=ticker,
        listing_date=info.get("listing_date"),
        industry=info.get("industry"),
        currency="TWD",
        extra={"board": info.get("board") or ""},
    )

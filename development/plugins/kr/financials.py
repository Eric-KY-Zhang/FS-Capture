"""KR financials via DART OpenAPI fnlttSinglAcntAll.

DART's `finstate_all` returns a long table where each row is one account line
with columns: sj_div (BS/IS/CF), account_nm (Korean), thstrm_amount (current).
We pivot into our canonical FinancialStatement structure.
"""
from __future__ import annotations

from typing import Optional

import pandas as pd
from loguru import logger

from app.core.models import (
    FinancialStatement,
    Period,
    PeriodType,
    StatementType,
    Ticker,
)
from .name_resolver import _dart


_REPRT_CODE = {
    PeriodType.Q1: "11013",
    PeriodType.Q2: "11012",
    PeriodType.Q3: "11014",
    PeriodType.ANNUAL: "11011",
}


_SJ_DIV_MAP = {
    "BS": StatementType.BALANCE_SHEET,
    "IS": StatementType.INCOME,
    "CIS": StatementType.INCOME,    # comprehensive income statement
    "CF": StatementType.CASH_FLOW,
    "SCE": None,                    # statement of changes in equity (skip)
}


def _to_float(value) -> Optional[float]:
    if value is None or value == "" or pd.isna(value):
        return None
    try:
        return float(str(value).replace(",", ""))
    except (ValueError, TypeError):
        return None


def fetch(ticker: Ticker, period: Period) -> list[FinancialStatement]:
    dart = _dart()
    reprt_code = _REPRT_CODE[period.type]

    try:
        df = dart.finstate_all(
            corp=ticker.external_id or ticker.code,
            bsns_year=str(period.year),
            reprt_code=reprt_code,
        )
    except Exception as exc:
        logger.warning(f"DART finstate_all failed for {ticker.code}: {exc}")
        return []
    if df is None or df.empty:
        return []

    by_stmt: dict[StatementType, dict[str, Optional[float]]] = {}
    for _, row in df.iterrows():
        sj = str(row.get("sj_div", "")).upper()
        st_type = _SJ_DIV_MAP.get(sj)
        if st_type is None:
            continue
        name = str(row.get("account_nm", "")).strip()
        if not name:
            continue
        val = _to_float(row.get("thstrm_amount"))
        bucket = by_stmt.setdefault(st_type, {})
        # Some accounts repeat (e.g. parent vs consolidated); keep first non-null.
        if name not in bucket or bucket[name] is None:
            bucket[name] = val

    out: list[FinancialStatement] = []
    for st_type, lines in by_stmt.items():
        if lines:
            out.append(FinancialStatement(
                ticker=ticker,
                period=period,
                statement_type=st_type,
                currency="KRW",
                unit="韩元",
                lines=lines,
            ))
    return out

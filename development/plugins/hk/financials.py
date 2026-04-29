"""HK financials via akshare (Eastmoney HK endpoints).

akshare exposes:
  ak.stock_financial_hk_report_em(stock="00700", symbol="资产负债表", indicator="年度")

`indicator` options: "年度" | "报告期" — the latter returns 半年/季度.
"""
from __future__ import annotations

import re
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


_STATEMENT_SYMBOL = {
    StatementType.BALANCE_SHEET: "资产负债表",
    StatementType.INCOME: "利润表",
    StatementType.CASH_FLOW: "现金流量表",
}

_ALIASES = {
    "总资产": "资产总计",
    "资产总额": "资产总计",
    "总负债": "负债合计",
    "股东权益": "股东权益合计",
    "总权益": "所有者权益(或股东权益)合计",
    "现金及等价物": "货币资金",
    "应收帐款": "应收账款",
    "短期贷款": "短期借款",
    "长期贷款": "长期借款",
    "营业额": "营业收入",
    "营运收入": "营业收入",
    "营运支出": "营业成本",
    "经营溢利": "营业利润",
    "除税前溢利": "利润总额",
    "除税后溢利": "净利润",
    "股东应占溢利": "归属于母公司所有者的净利润",
    "每股基本盈利": "基本每股收益",
    "每股摊薄盈利": "稀释每股收益",
    "经营业务现金净额": "经营活动现金流量净额",
    "投资业务现金净额": "投资活动现金流量净额",
    "融资业务现金净额": "筹资活动现金流量净额",
    "现金净额": "现金净增加额",
}


def _date_for_period(period: Period) -> str:
    md = {
        PeriodType.Q1: "0331",
        PeriodType.Q2: "0630",
        PeriodType.Q3: "0930",
        PeriodType.ANNUAL: "1231",
    }[period.type]
    return f"{period.year}{md}"


def _compact_date(value) -> str:
    return re.sub(r"\D", "", str(value))[:8]


def _select_period_row(df: pd.DataFrame, period: Period) -> Optional[pd.Series]:
    if df is None or df.empty:
        return None
    target = _date_for_period(period)
    for col in ("REPORT_DATE", "STD_REPORT_DATE", "报告期", "报告日期"):
        if col in df.columns:
            mask = df[col].map(_compact_date).eq(target)
            sel = df.loc[mask]
            if not sel.empty:
                if {"STD_ITEM_NAME", "AMOUNT"}.issubset(sel.columns):
                    lines = {}
                    for _, r in sel.iterrows():
                        name = str(r.get("STD_ITEM_NAME") or "").strip()
                        if not name:
                            continue
                        lines[name] = r.get("AMOUNT")
                    return pd.Series(lines)
                return sel.iloc[0]
    return None


def _parse_number(value) -> Optional[float]:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)) and pd.notna(value):
        return float(value)
    if not isinstance(value, str):
        return None
    text = value.strip().replace(",", "")
    if not text or text in {"--", "-", "False", "None", "nan"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _row_to_lines(row: pd.Series) -> dict[str, Optional[float]]:
    skip = {
        "SECUCODE", "SECURITY_CODE", "SECURITY_NAME_ABBR", "REPORT_DATE",
        "REPORT_TYPE", "DATE_TYPE_CODE", "STD_REPORT_DATE", "FISCAL_YEAR",
        "CURRENCY", "ORG_TYPE", "ORG_CODE",
    }
    out: dict[str, Optional[float]] = {}
    for k, v in row.items():
        if k in skip:
            continue
        parsed = _parse_number(v)
        if parsed is not None:
            out[str(k)] = parsed
    for source, target in _ALIASES.items():
        if source in out and target not in out:
            out[target] = out[source]
    return out


def fetch(ticker: Ticker, period: Period) -> list[FinancialStatement]:
    import akshare as ak

    indicator = "年度" if period.type is PeriodType.ANNUAL else "报告期"
    out: list[FinancialStatement] = []
    for st_type, sym in _STATEMENT_SYMBOL.items():
        try:
            df = ak.stock_financial_hk_report_em(
                stock=ticker.code, symbol=sym, indicator=indicator
            )
        except Exception as exc:
            logger.warning(f"akshare HK report failed for {ticker.code} {sym}: {exc}")
            continue
        row = _select_period_row(df, period)
        if row is None:
            logger.info(f"[{ticker.code}] no HK {st_type} for {period.label()}")
            continue
        lines = _row_to_lines(row)
        if not lines:
            continue
        out.append(FinancialStatement(
            ticker=ticker,
            period=period,
            statement_type=st_type,
            currency="HKD",
            unit="元",
            lines=lines,
        ))
    return out

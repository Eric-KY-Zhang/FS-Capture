"""A-share financial data via akshare (Eastmoney-backed endpoints).

Returns canonical 资产负债表 / 利润表 / 现金流量表 keyed by Chinese metric names
matching the existing 瑞华底稿 schema.

Period semantics:
- ANNUAL  -> fiscal year ending Dec 31 of `period.year`
- Q1      -> Mar 31
- Q2      -> Jun 30  (半年报)
- Q3      -> Sep 30
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


_PERIOD_MONTH_DAY = {
    PeriodType.Q1: ("03", "31"),
    PeriodType.Q2: ("06", "30"),
    PeriodType.Q3: ("09", "30"),
    PeriodType.ANNUAL: ("12", "31"),
}


def _period_key(period: Period) -> str:
    """Return the 'YYYY-MM-DD' string used by akshare/Eastmoney as a column header."""
    mm, dd = _PERIOD_MONTH_DAY[period.type]
    return f"{period.year}-{mm}-{dd}"


def _period_key_compact(period: Period) -> str:
    return _period_key(period).replace("-", "")


def _akshare_symbol(ticker: Ticker) -> str:
    """akshare Eastmoney endpoints want 'SH600519' / 'SZ000001' uppercase."""
    if ticker.code.startswith(("60", "688", "900")):
        return f"SH{ticker.code}"
    if ticker.code.startswith(("000", "001", "002", "200", "300", "301")):
        return f"SZ{ticker.code}"
    return f"BJ{ticker.code}"


def _sina_symbol(ticker: Ticker) -> str:
    return _akshare_symbol(ticker).lower()


def _select_period_column(df: pd.DataFrame, period: Period) -> Optional[pd.Series]:
    """The Eastmoney sheet endpoints return a DataFrame with a 'REPORT_DATE'
    column (or sometimes 'REPORT_DATE_NAME'). We pick the row matching our period.
    """
    if df is None or df.empty:
        return None
    target = _period_key(period)
    target_compact = _period_key_compact(period)
    # Some akshare versions use snake_case, others raw EM column names; check both.
    for col in ("REPORT_DATE", "报告日", "报告期", "报告日期"):
        if col in df.columns:
            values = df[col].astype(str)
            compact = values.str.replace(r"\D", "", regex=True).str[:8]
            mask = values.str.startswith(target) | compact.eq(target_compact)
            sel = df.loc[mask]
            if not sel.empty:
                return sel.iloc[0]
    # Try parsing any column that starts with the target.
    for col in df.columns:
        if str(col).startswith(target):
            return df[col]
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

    multiplier = 1.0
    if text.endswith("亿"):
        multiplier = 1e8
        text = text[:-1]
    elif text.endswith("万"):
        multiplier = 1e4
        text = text[:-1]
    elif text.endswith("%"):
        multiplier = 0.01
        text = text[:-1]

    try:
        return float(text) * multiplier
    except ValueError:
        return None


_META_COLUMNS = {
    "SECUCODE", "SECURITY_CODE", "SECURITY_NAME_ABBR", "ORG_CODE",
    "ORG_TYPE", "REPORT_DATE", "REPORT_TYPE", "REPORT_DATE_NAME",
    "SECURITY_TYPE_CODE", "NOTICE_DATE", "UPDATE_DATE", "CURRENCY",
    "BCKEY", "ROW_ID",
    "股票代码", "股票简称", "报告日", "报告期", "报告日期", "公告日期",
    "数据源", "是否审计", "币种", "类型", "更新日期", "序号",
}


_ALIASES = {
    "资产-货币资金": "货币资金",
    "资产-应收账款": "应收账款",
    "资产-存货": "存货",
    "资产-总资产": "资产总计",
    "负债-总负债": "负债合计",
    "股东权益合计": "所有者权益(或股东权益)合计",
    "营业总收入": "营业收入",
    "营业总支出-营业支出": "营业成本",
    "经营性现金流-现金流量净额": "经营活动产生的现金流量净额",
    "投资性现金流-现金流量净额": "投资活动产生的现金流量净额",
    "融资性现金流-现金流量净额": "筹资活动产生的现金流量净额",
}


def _row_to_lines(row: pd.Series) -> dict[str, Optional[float]]:
    """Convert a row (metric -> value) into our canonical dict.

    Skips meta columns and non-numeric rows.
    """
    out: dict[str, Optional[float]] = {}
    for k, v in row.items():
        if k in _META_COLUMNS:
            continue
        parsed = _parse_number(v)
        if parsed is not None:
            out[str(k)] = parsed

    for source, target in _ALIASES.items():
        if source in out and target not in out:
            out[target] = out[source]
    return out


def _safe_call(fn, *args, **kwargs) -> Optional[pd.DataFrame]:
    try:
        return fn(*args, **kwargs)
    except Exception as exc:
        logger.warning(f"akshare call {fn.__name__} failed: {exc}")
        return None


def _normalize_report_date(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    if df is None or df.empty:
        return df
    if "REPORT_DATE" in df.columns:
        return df
    for col in ("报告日", "报告期", "报告日期"):
        if col not in df.columns:
            continue
        out = df.copy()
        compact = out[col].astype(str).str.replace(r"\D", "", regex=True).str[:8]
        dates = pd.to_datetime(compact, format="%Y%m%d", errors="coerce")
        out["REPORT_DATE"] = dates.dt.strftime("%Y-%m-%d")
        return out
    return df


def _filter_ticker(df: Optional[pd.DataFrame], ticker: Ticker) -> Optional[pd.DataFrame]:
    if df is None or df.empty:
        return df
    for col in ("股票代码", "SECURITY_CODE"):
        if col not in df.columns:
            continue
        codes = df[col].astype(str).map(lambda value: _extract_code(value))
        return df.loc[codes == ticker.code.zfill(6)]
    return df


def _extract_code(value: str) -> str:
    match = re.search(r"\d{6}", value)
    return match.group(0) if match else value.strip().zfill(6)


def _sina_statement(ak, ticker: Ticker, symbol: str) -> Optional[pd.DataFrame]:
    df = _safe_call(ak.stock_financial_report_sina, stock=_sina_symbol(ticker), symbol=symbol)
    return _normalize_report_date(df)


def _eastmoney_summary(ak, ticker: Ticker, period: Period, fn_name: str) -> Optional[pd.DataFrame]:
    fn = getattr(ak, fn_name, None)
    if fn is None:
        return None
    df = _safe_call(fn, date=_period_key_compact(period))
    df = _filter_ticker(df, ticker)
    if df is None or df.empty:
        return None
    out = df.copy()
    out["REPORT_DATE"] = _period_key(period)
    return out


def _with_fallbacks(ak, ticker: Ticker, period: Period, bs, ic, cf):
    missing = {
        StatementType.BALANCE_SHEET: _select_period_column(bs, period) is None,
        StatementType.INCOME: _select_period_column(ic, period) is None,
        StatementType.CASH_FLOW: _select_period_column(cf, period) is None,
    }
    if not any(missing.values()):
        return bs, ic, cf

    logger.info(f"[{ticker.code}] trying Sina financial report fallback for {period.label()}")
    if missing[StatementType.BALANCE_SHEET]:
        fallback = _sina_statement(ak, ticker, "资产负债表")
        if _select_period_column(fallback, period) is not None:
            bs = fallback
            missing[StatementType.BALANCE_SHEET] = False
    if missing[StatementType.INCOME]:
        fallback = _sina_statement(ak, ticker, "利润表")
        if _select_period_column(fallback, period) is not None:
            ic = fallback
            missing[StatementType.INCOME] = False
    if missing[StatementType.CASH_FLOW]:
        fallback = _sina_statement(ak, ticker, "现金流量表")
        if _select_period_column(fallback, period) is not None:
            cf = fallback
            missing[StatementType.CASH_FLOW] = False

    if not any(missing.values()):
        return bs, ic, cf

    logger.info(f"[{ticker.code}] trying Eastmoney datacenter fallback for {period.label()}")
    if missing[StatementType.BALANCE_SHEET]:
        fn = "stock_zcfz_bj_em" if _akshare_symbol(ticker).startswith("BJ") else "stock_zcfz_em"
        fallback = _eastmoney_summary(ak, ticker, period, fn)
        if _select_period_column(fallback, period) is not None:
            bs = fallback
    if missing[StatementType.INCOME]:
        fallback = _eastmoney_summary(ak, ticker, period, "stock_lrb_em")
        if _select_period_column(fallback, period) is not None:
            ic = fallback
    if missing[StatementType.CASH_FLOW]:
        fallback = _eastmoney_summary(ak, ticker, period, "stock_xjll_em")
        if _select_period_column(fallback, period) is not None:
            cf = fallback
    return bs, ic, cf


def fetch(ticker: Ticker, period: Period) -> list[FinancialStatement]:
    import akshare as ak

    sym = _akshare_symbol(ticker)
    statements: list[FinancialStatement] = []

    # Annual uses the *_yearly_em endpoints; quarterly uses the regular *_by_report_em.
    if period.type is PeriodType.ANNUAL:
        bs = _safe_call(ak.stock_balance_sheet_by_yearly_em, symbol=sym)
        ic = _safe_call(ak.stock_profit_sheet_by_yearly_em, symbol=sym)
        cf = _safe_call(ak.stock_cash_flow_sheet_by_yearly_em, symbol=sym)
    else:
        bs = _safe_call(ak.stock_balance_sheet_by_report_em, symbol=sym)
        ic = _safe_call(ak.stock_profit_sheet_by_report_em, symbol=sym)
        cf = _safe_call(ak.stock_cash_flow_sheet_by_report_em, symbol=sym)

    bs, ic, cf = _with_fallbacks(ak, ticker, period, bs, ic, cf)

    for st_type, df in (
        (StatementType.BALANCE_SHEET, bs),
        (StatementType.INCOME, ic),
        (StatementType.CASH_FLOW, cf),
    ):
        row = _select_period_column(df, period)
        if row is None:
            logger.info(f"[{ticker.code}] no {st_type} for {period.label()}")
            continue
        lines = _row_to_lines(row)
        if not lines:
            continue
        statements.append(FinancialStatement(
            ticker=ticker,
            period=period,
            statement_type=st_type,
            currency="CNY",
            unit="元",
            lines=lines,
        ))
    return statements

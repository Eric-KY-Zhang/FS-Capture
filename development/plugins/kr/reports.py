"""KR filings download via DART OpenAPI.

OpenDartReader.document(rcp_no) returns the filing document XML text. It does
not save a file to disk in the current OpenDartReader release, so this module
writes that returned XML directly under the output tree.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import pandas as pd
from loguru import logger

from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path
from .name_resolver import _dart


_REPORT_KEYWORDS = {
    PeriodType.ANNUAL: ("사업보고서",),
    PeriodType.Q1: ("분기보고서",),
    PeriodType.Q2: ("반기보고서",),
    PeriodType.Q3: ("분기보고서",),
}

_KIND = {
    PeriodType.ANNUAL: "annual_report",
    PeriodType.Q1: "q1_report",
    PeriodType.Q2: "interim_report",
    PeriodType.Q3: "q3_report",
}

_DETAIL_KIND = {
    PeriodType.ANNUAL: "A001",
    PeriodType.Q1: "A003",
    PeriodType.Q2: "A002",
    PeriodType.Q3: "A003",
}


def _corp(ticker: Ticker) -> str:
    return ticker.external_id or ticker.code


def _list_filings(ticker: Ticker, period: Period) -> pd.DataFrame:
    dart = _dart()
    bgn = f"{period.year}0101"
    end = f"{period.year + 1}0630" if period.type is PeriodType.ANNUAL else f"{period.year}1231"
    try:
        df = dart.list(
            corp=_corp(ticker),
            start=bgn,
            end=end,
            kind="A",
            kind_detail=_DETAIL_KIND[period.type],
            final=True,
        )
    except Exception as exc:
        logger.warning(f"DART list failed for {ticker.code}: {exc}")
        return pd.DataFrame()
    return df if isinstance(df, pd.DataFrame) else pd.DataFrame()


def _date_text(row: dict) -> str:
    return str(row.get("rcept_dt") or row.get("rcept_de") or row.get("date") or "")


def _date_month(row: dict) -> Optional[int]:
    match = re.match(r"^\d{4}(\d{2})\d{2}$", _date_text(row))
    return int(match.group(1)) if match else None


def _select_filing(df: pd.DataFrame, period: Period) -> Optional[dict]:
    if df is None or df.empty:
        return None
    title_col = "report_nm" if "report_nm" in df.columns else "title"
    if title_col not in df.columns:
        return None

    keywords = _REPORT_KEYWORDS[period.type]
    rows = [
        row.to_dict()
        for _, row in df.iterrows()
        if any(k in str(row.get(title_col, "")) for k in keywords)
        and str(period.year) in str(row.get(title_col, ""))
    ]
    if not rows:
        return None

    if period.type is PeriodType.Q1:
        preferred_months = {4, 5, 6}
        preferred = [r for r in rows if _date_month(r) in preferred_months]
        if preferred:
            rows = preferred
    elif period.type is PeriodType.Q3:
        preferred_months = {10, 11, 12}
        preferred = [r for r in rows if _date_month(r) in preferred_months]
        if preferred:
            rows = preferred

    rows.sort(key=_date_text)
    return rows[-1] if period.type in {PeriodType.ANNUAL, PeriodType.Q3} else rows[0]


def _select_audit_filing(df: pd.DataFrame, period: Period) -> Optional[dict]:
    if df is None or df.empty:
        return None
    title_col = "report_nm" if "report_nm" in df.columns else "title"
    if title_col not in df.columns:
        return None
    rows = [
        row.to_dict()
        for _, row in df.iterrows()
        if "감사보고서" in str(row.get(title_col, ""))
        and str(period.year) in str(row.get(title_col, ""))
    ]
    if not rows:
        return None
    rows.sort(key=_date_text)
    return rows[-1]


def _write_document(rcept_no: str, dest: Path) -> Optional[int]:
    dart = _dart()
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        text = dart.document(rcept_no)
    except Exception as exc:
        logger.warning(f"DART document fetch failed for rcept {rcept_no}: {exc}")
        return None

    if isinstance(text, bytes):
        data = text
    else:
        data = str(text).encode("utf-8")

    tmp_dest = dest.with_name(f"{dest.name}.part")
    try:
        tmp_dest.write_bytes(data)
        tmp_dest.replace(dest)
    except Exception:
        tmp_dest.unlink(missing_ok=True)
        raise
    return len(data)


def _report_file(ticker: Ticker, period: Period, row: dict, output_root: Path, kind: str) -> Optional[ReportFile]:
    rcept = str(row.get("rcept_no") or row.get("rcept_no".upper()) or "")
    if not rcept:
        return None
    dest = report_output_path(output_root, ticker, period, kind, ".xml")
    n_bytes = _write_document(rcept, dest)
    if n_bytes is None:
        return None
    return ReportFile(
        ticker=ticker,
        period=period,
        kind=kind,
        local_path=str(dest),
        source_url=f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept}",
        title=str(row.get("report_nm") or row.get("title") or ""),
        file_size_bytes=n_bytes,
    )


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    df = _list_filings(ticker, period)
    out: list[ReportFile] = []

    main = _select_filing(df, period)
    if main:
        report = _report_file(ticker, period, main, output_root, _KIND[period.type])
        if report:
            out.append(report)

    if period.type is PeriodType.ANNUAL:
        audit = _select_audit_filing(df, period)
        if audit:
            report = _report_file(ticker, period, audit, output_root, "audit_report")
            if report:
                out.append(report)

    if not out:
        logger.warning(f"[{ticker.code}] no KR {period.label()} filings found")
    return out

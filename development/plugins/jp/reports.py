"""JP filings download via EDINET."""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any

import pandas as pd
from loguru import logger

from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path

from .name_resolver import _edinet_api_key

_KIND = {
    PeriodType.ANNUAL: "annual_report",
    PeriodType.Q1: "q1_report",
    PeriodType.Q2: "interim_report",
    PeriodType.Q3: "q3_report",
}

_DOC_TYPE_CODES = {
    PeriodType.ANNUAL: {"120", "130"},
    PeriodType.Q1: {"140", "150"},
    PeriodType.Q2: {"160", "170"},
    PeriodType.Q3: {"140", "150"},
}


def _date_range(start: dt.date, end: dt.date) -> list[str]:
    dates: list[str] = []
    current = start
    while current <= end:
        dates.append(current.isoformat())
        current += dt.timedelta(days=1)
    return dates


def _scan_dates(period: Period) -> list[str]:
    if period.type is PeriodType.ANNUAL:
        return _date_range(dt.date(period.year, 1, 1), dt.date(period.year, 12, 31))
    return _date_range(dt.date(period.year, 1, 1), dt.date(period.year, 12, 31))


def _matches_ticker(row: dict[str, Any], ticker: Ticker) -> bool:
    if ticker.external_id and str(row.get("edinet_code") or "") == ticker.external_id:
        return True
    return str(row.get("sec_code") or "") == ticker.code


def _list_documents(submit_date: str) -> list[dict[str, Any]]:
    api_key = _edinet_api_key()
    if api_key:
        from .edinet_api import list_documents

        return list_documents(submit_date, api_key=api_key)

    from .edinet_web import list_documents

    return list_documents(submit_date)


def _matching_rows(
    documents: list[dict[str, Any]],
    ticker: Ticker,
    doc_types: set[str],
    seen: set[str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in documents:
        doc_id = str(row.get("doc_id") or "")
        if not doc_id or doc_id in seen:
            continue
        if str(row.get("doc_type_code") or "") not in doc_types:
            continue
        if not _matches_ticker(row, ticker):
            continue
        seen.add(doc_id)
        rows.append(row)
    return rows


def _list_filings_via_api(
    ticker: Ticker, period: Period, doc_types: set[str], api_key: str
) -> list[dict[str, Any]]:
    from .edinet_api import list_documents

    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for submit_date in _scan_dates(period):
        try:
            documents = list_documents(submit_date, api_key=api_key)
        except Exception as exc:
            logger.warning(f"EDINET list failed for {submit_date}: {exc}")
            continue
        rows.extend(_matching_rows(documents, ticker, doc_types, seen))
    return rows


def _list_filings_via_public(
    ticker: Ticker, period: Period, doc_types: set[str]
) -> list[dict[str, Any]]:
    from .edinet_web import search_filings

    try:
        documents = search_filings(ticker.code, period.year)
    except Exception as exc:
        logger.warning(f"EDINET public search failed for {ticker.code} {period.year}: {exc}")
        return []
    return _matching_rows(documents, ticker, doc_types, set())


def _list_filings(ticker: Ticker, period: Period) -> pd.DataFrame:
    doc_types = _DOC_TYPE_CODES.get(period.type)
    if not doc_types:
        return pd.DataFrame()

    api_key = _edinet_api_key()
    if api_key:
        rows = _list_filings_via_api(ticker, period, doc_types, api_key)
    else:
        rows = _list_filings_via_public(ticker, period, doc_types)

    return pd.DataFrame(rows)


def _select_filing(df: pd.DataFrame, period: Period) -> dict[str, Any] | None:
    if df is None or df.empty:
        return None
    rows = [row.to_dict() for _, row in df.iterrows()]
    if period.type is PeriodType.ANNUAL:
        preferred = [row for row in rows if str(row.get("doc_type_code") or "") == "120"]
        if preferred:
            rows = preferred
    rows.sort(key=lambda row: str(row.get("submit_date_time") or ""))
    return rows[-1]


def _download_doc_as_pdf(doc_id: str, dest: Path) -> tuple[str, str, int]:
    api_key = _edinet_api_key()
    if api_key:
        from .edinet_api import download_document_pdf

        n_bytes = download_document_pdf(doc_id, dest, api_key=api_key)
        source_url = f"https://api.edinet-fsa.go.jp/api/v2/documents/{doc_id}?type=2"
    else:
        from .edinet_web import download_document_pdf

        n_bytes = download_document_pdf(doc_id, dest)
        source_url = f"https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/{doc_id}.pdf"
    return source_url, "pdf", n_bytes


def _filing_date(row: dict[str, Any]) -> str | None:
    value = str(row.get("submit_date_time") or "").strip()
    if not value:
        return None
    return value[:10]


def _report_file(
    ticker: Ticker, period: Period, row: dict[str, Any], output_root: Path, kind: str
) -> ReportFile | None:
    doc_id = str(row.get("doc_id") or "").strip()
    if not doc_id:
        return None
    dest = report_output_path(output_root, ticker, period, kind, ".pdf")
    try:
        source_url, source_format, n_bytes = _download_doc_as_pdf(doc_id, dest)
    except Exception:
        return None
    return ReportFile(
        ticker=ticker,
        period=period,
        kind=kind,
        local_path=str(dest),
        source_url=source_url,
        title=str(row.get("title") or ""),
        file_size_bytes=n_bytes,
        filing_date=_filing_date(row),
        report_date=str(row.get("period_end") or "") or None,
        form=f"EDINET-{row.get('doc_type_code') or ''}",
        source_id=doc_id,
        accession_number=doc_id,
        source_format=source_format,
        output_format="pdf",
    )


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if period.type is PeriodType.IPO_PROSPECTUS:
        logger.warning(f"[{ticker.code}] JP IPO prospectus search is not implemented")
        return []
    if period.type not in _KIND:
        logger.warning(f"[{ticker.code}] unsupported JP period type {period.type}")
        return []

    df = _list_filings(ticker, period)
    row = _select_filing(df, period)
    if row is None:
        logger.warning(f"[{ticker.code}] no JP {period.label()} filings found")
        return []

    report = _report_file(ticker, period, row, output_root, _KIND[period.type])
    return [report] if report else []

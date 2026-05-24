"""SG filings download via SGXNet public APIs."""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any

from loguru import logger

from app.core.cache import cached_or_load
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path, report_output_path_for_filing

from .sgxnet_web import (
    download_pdf,
    extract_pdf_links,
    list_financial_reports,
    list_ipo_prospectuses,
    search_announcements,
)

_TTL = 24 * 3600

_KIND = {
    PeriodType.ANNUAL: "annual_report",
    PeriodType.Q2: "interim_report",
    PeriodType.IPO_PROSPECTUS: "ipo_prospectus",
}


def _date_from_ms(value: int | float | None) -> dt.date | None:
    if value is None:
        return None
    return dt.datetime.fromtimestamp(value / 1000, dt.UTC).date()


def _date_from_yyyymmdd(value: str | None) -> dt.date | None:
    if not value:
        return None
    try:
        return dt.datetime.strptime(value, "%Y%m%d").date()
    except ValueError:
        return None


def _financial_reports() -> list[dict[str, Any]]:
    return cached_or_load("sg:financialreports:v1", list_financial_reports, expire=_TTL)


def _ipo_prospectuses() -> list[dict[str, Any]]:
    return cached_or_load("sg:ipoprospectus:v1", list_ipo_prospectuses, expire=_TTL)


def _matches_ticker(row: dict[str, Any], ticker: Ticker) -> bool:
    target = str(ticker.name or ticker.code).strip().upper()
    if not target:
        return False
    fields = [
        str(row.get("companyName") or "").strip().upper(),
        str(row.get("securityName") or "").strip().upper(),
    ]
    return any(field == target or target == field for field in fields)


def _list_annual_filings(ticker: Ticker, period: Period) -> list[dict[str, Any]]:
    rows = []
    for row in _financial_reports():
        document_date = _date_from_ms(row.get("documentDate"))
        if not document_date or document_date.year != period.year:
            continue
        if str(row.get("title") or "").strip().upper() != "ANNUAL REPORT":
            continue
        if not _matches_ticker(row, ticker):
            continue
        rows.append(row)
    return rows


def _list_h1_filings(ticker: Ticker, period: Period) -> list[dict[str, Any]]:
    rows = []
    for row in search_announcements(ticker.code, start_year=period.year, end_year=period.year):
        submission_date = _date_from_yyyymmdd(str(row.get("submission_date") or ""))
        if not submission_date or submission_date.year != period.year:
            continue
        category = str(row.get("category_name") or "").upper()
        title = str(row.get("title") or "").upper()
        if "FINANCIAL STATEMENTS" not in category:
            continue
        if "HALF YEARLY RESULTS" not in title and "SECOND QUARTER" not in title:
            continue
        rows.append(row)
    return rows


def _list_ipo_filings(ticker: Ticker, period: Period) -> list[dict[str, Any]]:
    target = str(ticker.name or ticker.code).strip().upper()
    rows = []
    for row in _ipo_prospectuses():
        closing_date = _date_from_ms(row.get("closing_date"))
        if not closing_date or closing_date.year != period.year:
            continue
        name = str(row.get("name") or "").strip().upper()
        if target and target not in name and name not in target:
            continue
        rows.append(row)
    return rows


def _date_sort_key(row: dict[str, Any]) -> str:
    if row.get("documentDate") is not None:
        value = _date_from_ms(row.get("documentDate"))
        return value.isoformat() if value else ""
    if row.get("closing_date") is not None:
        value = _date_from_ms(row.get("closing_date"))
        return value.isoformat() if value else ""
    return str(row.get("submission_date") or "")


def _select_filing(rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not rows:
        return None
    rows = list(rows)
    rows.sort(key=_date_sort_key)
    return rows[-1]


def _choose_pdf_link(kind: str, links: list[str]) -> str | None:
    def score(link: str) -> tuple[int, str]:
        text = link.lower()
        if kind == "annual_report":
            value = 0 if "annual%20report" in text or "annual report" in text else 5
            if "sustainability" in text or "shareholder" in text or "letter" in text:
                value += 10
            return value, text
        if kind == "interim_report":
            value = 0
            if "interim" in text or "financial%20statement" in text:
                value -= 5
            if any(term in text for term in ("news", "presentation", "slides")):
                value += 10
            return value, text
        if kind == "ipo_prospectus":
            value = 0
            if "summary" in text or "highlights" in text:
                value += 10
            return value, text
        return 0, text

    if not links:
        return None
    return sorted(links, key=score)[0]


def _download_page_as_pdf(row: dict[str, Any], dest: Path, kind: str) -> tuple[str, str, int]:
    page_url = str(row.get("url") or "").strip()
    if not page_url:
        raise ValueError("SGXNet row has no page URL")
    links = extract_pdf_links(page_url)
    pdf_url = _choose_pdf_link(kind, links)
    if not pdf_url:
        raise ValueError(f"SGXNet page has no PDF attachment: {page_url}")
    n_bytes = download_pdf(pdf_url, dest)
    return pdf_url, "pdf", n_bytes


def _filing_date(row: dict[str, Any]) -> str | None:
    if row.get("submission_date"):
        value = _date_from_yyyymmdd(str(row.get("submission_date")))
        return value.isoformat() if value else None
    if row.get("closing_date") is not None:
        value = _date_from_ms(row.get("closing_date"))
        return value.isoformat() if value else None
    return None


def _report_date(row: dict[str, Any]) -> str | None:
    value = _date_from_ms(row.get("documentDate"))
    return value.isoformat() if value else None


def _source_id(row: dict[str, Any]) -> str:
    return str(row.get("id") or row.get("ref_id") or "").strip()


def _report_file(
    ticker: Ticker,
    period: Period,
    row: dict[str, Any],
    output_root: Path,
    kind: str,
    *,
    sequence: int | None = None,
) -> ReportFile | None:
    if kind == "ipo_prospectus":
        dest = report_output_path_for_filing(
            output_root,
            ticker,
            "ipo",
            sequence,
            kind,
            ".pdf",
            filing_date=_filing_date(row),
            source_id=_source_id(row) or None,
        )
    else:
        dest = report_output_path(output_root, ticker, period, kind, ".pdf")
    try:
        source_url, source_format, n_bytes = _download_page_as_pdf(row, dest, kind)
    except Exception as exc:
        logger.warning(f"SGXNet download failed for {ticker.code}: {exc}")
        return None
    source_id = _source_id(row)
    return ReportFile(
        ticker=ticker,
        period=period,
        kind=kind,
        local_path=str(dest),
        source_url=source_url,
        title=str(row.get("title") or row.get("name") or ""),
        file_size_bytes=n_bytes,
        filing_date=_filing_date(row),
        report_date=_report_date(row),
        form=str(row.get("category_name") or row.get("status") or "") or None,
        source_id=source_id or None,
        accession_number=source_id or None,
        sequence=sequence,
        source_format=source_format,
        output_format="pdf",
    )


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    kind = _KIND.get(period.type)
    if not kind:
        logger.warning(f"[{ticker.code}] unsupported SG period type {period.type}")
        return []

    if period.type is PeriodType.ANNUAL:
        rows = _list_annual_filings(ticker, period)
    elif period.type is PeriodType.Q2:
        rows = _list_h1_filings(ticker, period)
    else:
        rows = _list_ipo_filings(ticker, period)

    row = _select_filing(rows)
    if row is None:
        logger.warning(f"[{ticker.code}] no SG {period.label()} filings found")
        return []

    report = _report_file(ticker, period, row, output_root, kind, sequence=1)
    return [report] if report else []

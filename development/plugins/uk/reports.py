"""UK filings download via FCA National Storage Mechanism."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from loguru import logger

from app.core.http import default_client, stream_to_file
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path

from .nsm_web import artifact_url, nsm_rate, search_filings

_KIND = {
    PeriodType.ANNUAL: "annual_report",
    PeriodType.Q1: "q1_report",
    PeriodType.Q2: "interim_report",
    PeriodType.Q3: "q3_report",
}

_KEYWORDS = {
    PeriodType.ANNUAL: ("Annual Report", "Annual Financial Report"),
    PeriodType.Q1: ("Q1 Trading Update", "First Quarter Trading Update"),
    PeriodType.Q2: ("Half-year Report", "Half-yearly Financial Report", "Interim Results"),
    PeriodType.Q3: ("Q3 Trading Update", "Third Quarter Trading Update"),
}


def _row_id(row: dict[str, Any]) -> str:
    return str(row.get("_id") or row.get("disclosure_id") or row.get("seq_id") or "")


def _search_years(period: Period) -> list[int]:
    if period.type is PeriodType.ANNUAL:
        return [period.year + 1, period.year]
    return [period.year]


def _matches_period(row: dict[str, Any], period: Period) -> bool:
    if period.type is not PeriodType.ANNUAL:
        return True
    year_text = str(period.year)
    document_date = str(row.get("document_date") or "")
    headline = str(row.get("headline") or "")
    return document_date.startswith(year_text) or year_text in headline


def _list_filings(ticker: Ticker, period: Period) -> pd.DataFrame:
    keywords = _KEYWORDS.get(period.type)
    if not keywords:
        return pd.DataFrame()

    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    company = ticker.name or ticker.code
    lei = ticker.external_id or ""
    for search_year in _search_years(period):
        for keyword in keywords:
            try:
                found = search_filings(company, lei, search_year, keyword)
            except Exception as exc:
                logger.warning(f"NSM search failed for {ticker.code} {keyword}: {exc}")
                continue
            for row in found:
                if not _matches_period(row, period):
                    continue
                key = _row_id(row) or str(row.get("download_link") or row.get("html_link") or "")
                if not key or key in seen:
                    continue
                seen.add(key)
                rows.append(row)

    return pd.DataFrame(rows)


def _date_text(row: dict[str, Any]) -> str:
    return str(row.get("publication_date") or row.get("submitted_date") or "")


def _date_iso(row: dict[str, Any]) -> str | None:
    text = _date_text(row)
    return text[:10] if text else None


def _link_text(row: dict[str, Any]) -> str:
    return str(row.get("download_link") or row.get("html_link") or "")


def _selection_score(row: dict[str, Any], period: Period) -> tuple[int, int, str]:
    headline = str(row.get("headline") or "").lower()
    link = _link_text(row).lower()
    html_link = str(row.get("html_link") or "").strip()
    if link.endswith(".pdf"):
        format_score = 0
    elif html_link or link.endswith((".html", ".htm")):
        format_score = 1
    else:
        format_score = 5

    title_score = 0
    if period.type is PeriodType.ANNUAL:
        if "annual report" in headline:
            title_score -= 1
        if "annual financial report" in headline:
            title_score += 1
    return format_score, title_score, _date_text(row)


def _select_filing(df: pd.DataFrame, period: Period) -> dict[str, Any] | None:
    if df is None or df.empty:
        return None
    rows = [row.to_dict() for _, row in df.iterrows()]
    rows.sort(key=lambda row: _selection_score(row, period))
    return rows[0]


def _is_pdf_file(path: Path) -> bool:
    try:
        return path.read_bytes()[:4] == b"%PDF"
    except OSError:
        return False


def _render_url_to_pdf(url: str, dest: Path) -> int:
    from app.core.pdf_renderer import render_url_to_pdf

    n_bytes = render_url_to_pdf(url, dest, source="nsm")
    if not _is_pdf_file(dest):
        dest.unlink(missing_ok=True)
        raise ValueError(f"NSM render did not produce a PDF for {url}")
    return n_bytes


def _download_row_as_pdf(row: dict[str, Any], dest: Path) -> tuple[str, str, int]:
    download = str(row.get("download_link") or "").strip()
    html = str(row.get("html_link") or "").strip()
    source_path = html if html else download
    if not source_path:
        raise ValueError("NSM row has no downloadable artefact")

    source_url = artifact_url(source_path)
    if download.lower().endswith(".pdf"):
        source_url = artifact_url(download)
        with default_client(source="nsm", timeout=120.0) as client:
            n_bytes = stream_to_file(
                client,
                source_url,
                dest,
                source="nsm",
                rate=nsm_rate(),
                read_timeout=180.0,
            )
        if not _is_pdf_file(dest):
            dest.unlink(missing_ok=True)
            raise ValueError(f"NSM artefact is not a PDF: {source_url}")
        return source_url, "pdf", n_bytes

    n_bytes = _render_url_to_pdf(source_url, dest)
    return source_url, "html", n_bytes


def _report_file(
    ticker: Ticker, period: Period, row: dict[str, Any], output_root: Path, kind: str
) -> ReportFile | None:
    dest = report_output_path(output_root, ticker, period, kind, ".pdf")
    try:
        source_url, source_format, n_bytes = _download_row_as_pdf(row, dest)
    except Exception as exc:
        logger.warning(f"NSM download failed for {ticker.code}: {exc}")
        return None

    row_id = _row_id(row)
    return ReportFile(
        ticker=ticker,
        period=period,
        kind=kind,
        local_path=str(dest),
        source_url=source_url,
        title=str(row.get("headline") or row.get("type") or ""),
        file_size_bytes=n_bytes,
        filing_date=_date_iso(row),
        report_date=str(row.get("document_date") or "") or None,
        form=str(row.get("type") or row.get("type_code") or "") or None,
        source_id=row_id or None,
        accession_number=row_id or None,
        source_format=source_format,
        output_format="pdf",
    )


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if period.type is PeriodType.IPO_PROSPECTUS:
        logger.warning(f"[{ticker.code}] UK IPO prospectus search is not implemented")
        return []
    if period.type not in _KIND:
        logger.warning(f"[{ticker.code}] unsupported UK period type {period.type}")
        return []

    df = _list_filings(ticker, period)
    row = _select_filing(df, period)
    if row is None:
        logger.warning(f"[{ticker.code}] no UK {period.label()} filings found")
        return []

    report = _report_file(ticker, period, row, output_root, _KIND[period.type])
    return [report] if report else []

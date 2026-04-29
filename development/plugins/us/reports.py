"""US filings download via SEC EDGAR submissions API.

Form types we surface:
- 10-K, 10-K/A    -> annual_report (US domestic issuer)
- 20-F            -> annual_report (foreign private issuer)
- 10-Q, 10-Q/A    -> q1/q3 (period_of_report drives which)
- 6-K             -> interim_report (foreign issuer)

The auditor's report ("Report of Independent Registered Public Accounting Firm")
is contained inside the 10-K / 20-F document — not filed separately.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Optional

from loguru import logger

from app.core.http import default_client, get_json, stream_to_file
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path


_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
_ARCHIVE_BASE = "https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_no_clean}/{filename}"


_FORMS_BY_PERIOD = {
    PeriodType.ANNUAL: ("10-K", "10-K/A", "20-F", "20-F/A", "40-F"),
    PeriodType.Q1: ("10-Q", "10-Q/A"),
    PeriodType.Q2: ("10-Q", "10-Q/A", "6-K"),
    PeriodType.Q3: ("10-Q", "10-Q/A"),
}


def _parse_date(value: str) -> Optional[dt.date]:
    try:
        return dt.date.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def _table_rows(table: dict) -> list[dict]:
    forms = table.get("form") or []
    accs = table.get("accessionNumber") or []
    primary_docs = table.get("primaryDocument") or []
    primary_doc_descrs = table.get("primaryDocDescription") or []
    filing_dates = table.get("filingDate") or []
    report_dates = table.get("reportDate") or table.get("periodOfReport") or []

    rows: list[dict] = []
    for i, form in enumerate(forms):
        rows.append({
            "form": form,
            "accessionNumber": accs[i] if i < len(accs) else "",
            "primaryDocument": primary_docs[i] if i < len(primary_docs) else "",
            "primaryDocDescription": primary_doc_descrs[i] if i < len(primary_doc_descrs) else "",
            "filingDate": filing_dates[i] if i < len(filing_dates) else "",
            "reportDate": report_dates[i] if i < len(report_dates) else "",
        })
    return rows


def _filter_table(table: dict, period: Period) -> list[dict]:
    """Walk a flat columnar dict (recent / paged files) for matching rows."""
    rows = _table_rows(table)
    target_forms = set(_FORMS_BY_PERIOD[period.type])
    if period.type is PeriodType.ANNUAL:
        return [
            r for r in rows
            if r["form"] in target_forms
            and r["reportDate"].startswith(str(period.year))
        ]

    annual_dates = sorted(
        d for d in (
            _parse_date(r["reportDate"])
            for r in rows
            if r["form"] in _FORMS_BY_PERIOD[PeriodType.ANNUAL]
        )
        if d is not None
    )
    fiscal_year_end = next((d for d in annual_dates if d.year == period.year), None)
    previous_year_end = (
        max((d for d in annual_dates if d < fiscal_year_end), default=None)
        if fiscal_year_end else None
    )

    if fiscal_year_end and previous_year_end:
        quarter_rows = [
            r for r in rows
            if r["form"] in target_forms
            and (report_date := _parse_date(r["reportDate"])) is not None
            and previous_year_end < report_date < fiscal_year_end
        ]
        quarter_rows.sort(key=lambda r: _parse_date(r["reportDate"]) or dt.date.min)
        target_index = {
            PeriodType.Q1: 0,
            PeriodType.Q2: 1,
            PeriodType.Q3: 2,
        }[period.type]
        if len(quarter_rows) > target_index:
            return [quarter_rows[target_index]]

    # Fallback for calendar-year issuers or foreign 6-K issuers where the
    # annual cycle cannot be reconstructed from the current submissions page.
    calendar_end = {
        PeriodType.Q1: "03-31",
        PeriodType.Q2: "06-30",
        PeriodType.Q3: "09-30",
    }[period.type]
    return [
        r for r in rows
        if r["form"] in target_forms
        and r["reportDate"].startswith(str(period.year))
        and r["reportDate"].endswith(calendar_end)
    ]


def _filter_filings(client, filings: dict, period: Period) -> list[dict]:
    """Search recent + paged files for matches."""
    rows = _filter_table(filings.get("recent", {}), period)
    if rows:
        return rows
    # Paged older filings live in `files`
    for f in filings.get("files") or []:
        name = f.get("name")
        if not name:
            continue
        url = f"https://data.sec.gov/submissions/{name}"
        try:
            page = get_json(client, url, source="sec", rate=8.0)
        except Exception as exc:
            logger.warning(f"failed to fetch paged submissions {name}: {exc}")
            continue
        rows = _filter_table(page, period)
        if rows:
            return rows
    return []


def _accession_to_archive(cik: str, acc_no: str, filename: str) -> str:
    cik_int = str(int(cik))
    return _ARCHIVE_BASE.format(cik_int=cik_int, acc_no_clean=acc_no.replace("-", ""), filename=filename)


def _kind_for(form: str, period: Period) -> str:
    if form.startswith(("10-K", "20-F", "40-F")):
        return "annual_report"
    if form.startswith("10-Q"):
        return f"{period.type.value}_report"
    if form.startswith("6-K"):
        return "interim_report"
    return "other"


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    cik = ticker.external_id
    if not cik:
        raise ValueError(f"美股下载需要 CIK：{ticker.code}")

    url = _SUBMISSIONS_URL.format(cik=cik)
    with default_client(source="sec") as client:
        payload = get_json(client, url, source="sec", rate=8.0)
        filings = payload.get("filings") or {}
        rows = _filter_filings(client, filings, period)

        if not rows:
            logger.warning(f"[{ticker.code}] no SEC filings for {period.label()}")
            return []

        # Prefer the canonical form (no /A amendment) and earliest filing date for the period.
        rows.sort(key=lambda r: (("/A" in r["form"]), r["filingDate"]))
        chosen = rows[0]

        acc_no = chosen["accessionNumber"]
        primary = chosen["primaryDocument"]
        url_doc = _accession_to_archive(cik, acc_no, primary)

        kind = _kind_for(chosen["form"], period)
        ext = Path(primary).suffix or ".htm"
        dest = report_output_path(output_root, ticker, period, kind, ext)
        n_bytes = stream_to_file(client, url_doc, dest, source="sec", rate=8.0)

        return [ReportFile(
            ticker=ticker,
            period=period,
            kind=kind,
            local_path=str(dest),
            source_url=url_doc,
            title=f"{chosen['form']} ({chosen['reportDate']})",
            file_size_bytes=n_bytes,
        )]

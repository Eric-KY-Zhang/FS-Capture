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
import re
from pathlib import Path
from typing import Any

from loguru import logger

from app.core.http import default_client, get_json, stream_to_file
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path, report_output_path_for_filing

_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
_ARCHIVE_BASE = "https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_no_clean}/{filename}"


_FORMS_BY_PERIOD = {
    PeriodType.ANNUAL: ("10-K", "10-K/A", "20-F", "20-F/A", "40-F"),
    PeriodType.Q1: ("10-Q", "10-Q/A"),
    PeriodType.Q2: ("10-Q", "10-Q/A", "6-K"),
    PeriodType.Q3: ("10-Q", "10-Q/A"),
}

_IPO_INITIAL_FORMS = {"S-1", "F-1"}
_IPO_AMENDMENT_FORMS = {"S-1/A", "F-1/A"}
_IPO_REGISTRATION_FORMS = _IPO_INITIAL_FORMS | _IPO_AMENDMENT_FORMS
_PROSPECTUS_424B_RE = re.compile(r"^424B\d+$", re.IGNORECASE)


def _parse_date(value: str) -> dt.date | None:
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
    acceptance_dates = table.get("acceptanceDateTime") or []
    file_numbers = table.get("fileNumber") or []
    sizes = table.get("size") or []

    rows: list[dict] = []
    for i, form in enumerate(forms):
        rows.append(
            {
                "form": form,
                "accessionNumber": accs[i] if i < len(accs) else "",
                "primaryDocument": primary_docs[i] if i < len(primary_docs) else "",
                "primaryDocDescription": primary_doc_descrs[i]
                if i < len(primary_doc_descrs)
                else "",
                "filingDate": filing_dates[i] if i < len(filing_dates) else "",
                "reportDate": report_dates[i] if i < len(report_dates) else "",
                "acceptanceDateTime": acceptance_dates[i] if i < len(acceptance_dates) else "",
                "fileNumber": file_numbers[i] if i < len(file_numbers) else "",
                "size": sizes[i] if i < len(sizes) else "",
            }
        )
    return rows


def _filter_table(table: dict, period: Period) -> list[dict]:
    """Walk a flat columnar dict (recent / paged files) for matching rows."""
    rows = _table_rows(table)
    target_forms = set(_FORMS_BY_PERIOD[period.type])
    if period.type is PeriodType.ANNUAL:
        return [
            r
            for r in rows
            if r["form"] in target_forms and r["reportDate"].startswith(str(period.year))
        ]

    annual_dates = sorted(
        d
        for d in (
            _parse_date(r["reportDate"])
            for r in rows
            if r["form"] in _FORMS_BY_PERIOD[PeriodType.ANNUAL]
        )
        if d is not None
    )
    fiscal_year_end = next((d for d in annual_dates if d.year == period.year), None)
    previous_year_end = (
        max((d for d in annual_dates if d < fiscal_year_end), default=None)
        if fiscal_year_end
        else None
    )

    if fiscal_year_end and previous_year_end:
        quarter_rows = [
            r
            for r in rows
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
        r
        for r in rows
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


def _all_filing_rows(client, filings: dict) -> list[dict]:
    """Load submissions recent + every paged file, oldest/newest ordering later."""
    rows = _table_rows(filings.get("recent", {}))
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
        rows.extend(_table_rows(page))
    return _dedupe_rows(rows)


def _dedupe_rows(rows: list[dict]) -> list[dict]:
    seen: set[tuple[str, str]] = set()
    out: list[dict] = []
    for row in rows:
        key = (
            str(row.get("accessionNumber") or ""),
            str(row.get("primaryDocument") or ""),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def _accession_to_archive(cik: str, acc_no: str, filename: str) -> str:
    cik_int = str(int(cik))
    return _ARCHIVE_BASE.format(
        cik_int=cik_int, acc_no_clean=acc_no.replace("-", ""), filename=filename
    )


def _normalize_form(row_or_form: dict | str) -> str:
    if isinstance(row_or_form, dict):
        value = row_or_form.get("form")
    else:
        value = row_or_form
    return str(value or "").strip().upper()


def _accession(row: dict) -> str:
    return str(row.get("accessionNumber") or "").strip()


def _file_number(row: dict) -> str:
    return str(row.get("fileNumber") or "").strip()


def _filing_sort_key(row: dict) -> tuple[dt.date, str, str]:
    return (
        _parse_date(str(row.get("filingDate") or "")) or dt.date.min,
        str(row.get("acceptanceDateTime") or ""),
        _accession(row),
    )


def _is_amendment_form(form: str) -> bool:
    return _normalize_form(form).endswith("/A")


def _is_424b(form: str) -> bool:
    return bool(_PROSPECTUS_424B_RE.match(_normalize_form(form)))


def _source_format(primary_document: str) -> str:
    return "pdf" if Path(primary_document).suffix.lower() == ".pdf" else "html"


def _render_url_to_pdf(url: str, dest: Path) -> int:
    from app.core.pdf_renderer import render_url_to_pdf

    result = render_url_to_pdf(url, dest, source="sec")
    if isinstance(result, int):
        return result
    if dest.exists():
        return dest.stat().st_size
    raise FileNotFoundError(f"PDF renderer did not create {dest}")


def _download_primary_as_pdf(client: Any, cik: str, row: dict, dest: Path) -> tuple[str, str, int]:
    acc_no = _accession(row)
    primary = str(row.get("primaryDocument") or "").strip()
    if not acc_no or not primary:
        raise ValueError(f"SEC filing row is missing accessionNumber/primaryDocument: {row}")

    url_doc = _accession_to_archive(cik, acc_no, primary)
    source_format = _source_format(primary)
    if source_format == "pdf":
        n_bytes = stream_to_file(client, url_doc, dest, source="sec", rate=8.0)
    else:
        dest.parent.mkdir(parents=True, exist_ok=True)
        n_bytes = _render_url_to_pdf(url_doc, dest)
    return url_doc, source_format, n_bytes


def _kind_for(form: str, period: Period) -> str:
    form = _normalize_form(form)
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
    if period.type is PeriodType.IPO_PROSPECTUS:
        reports = download_ipo_documents(ticker, output_root)
        for report in reports:
            report.period = period
            report.kind = "ipo_prospectus"
        return reports

    url = _SUBMISSIONS_URL.format(cik=cik)
    with default_client(source="sec") as client:
        payload = get_json(client, url, source="sec", rate=8.0)
        filings = payload.get("filings") or {}
        rows = _filter_filings(client, filings, period)

        if not rows:
            logger.warning(f"[{ticker.code}] no SEC filings for {period.label()}")
            return []

        # Prefer the canonical form (no /A amendment) and earliest filing date for the period.
        rows.sort(key=lambda r: ((_is_amendment_form(r["form"])), r["filingDate"]))
        chosen = rows[0]

        kind = _kind_for(chosen["form"], period)
        dest = report_output_path(output_root, ticker, period, kind, ".pdf")
        url_doc, source_format, n_bytes = _download_primary_as_pdf(client, cik, chosen, dest)
        form = _normalize_form(chosen)

        return [
            ReportFile(
                ticker=ticker,
                period=period,
                kind=kind,
                local_path=str(dest),
                source_url=url_doc,
                title=f"{chosen['form']} ({chosen['reportDate']})",
                file_size_bytes=n_bytes,
                filing_date=chosen.get("filingDate") or None,
                report_date=chosen.get("reportDate") or None,
                form=form,
                source_id=_file_number(chosen) or None,
                accession_number=_accession(chosen) or None,
                is_amendment=_is_amendment_form(form),
                source_format=source_format,
                output_format="pdf",
            )
        ]


def _select_final_424b(rows: list[dict]) -> dict | None:
    prospectuses = [r for r in rows if _is_424b(str(r.get("form") or ""))]
    if not prospectuses:
        return None

    preferred = [r for r in prospectuses if _normalize_form(r) == "424B4"]
    candidates = preferred or prospectuses
    candidates.sort(key=_filing_sort_key)
    return candidates[-1]


def _sorted_unique_chain(rows: list[dict]) -> list[dict]:
    return sorted(_dedupe_rows(rows), key=_filing_sort_key)


def _select_ipo_chain_by_file_number(rows: list[dict]) -> list[dict]:
    initials = [r for r in rows if _normalize_form(r) in _IPO_INITIAL_FORMS and _file_number(r)]
    if not initials:
        return []

    file_numbers = sorted({_file_number(r) for r in initials})
    candidates: list[tuple[tuple[dt.date, str, str], str]] = []
    for file_no in file_numbers:
        same_file = [r for r in rows if _file_number(r) == file_no]
        same_initials = [r for r in same_file if _normalize_form(r) in _IPO_INITIAL_FORMS]
        if not same_initials:
            continue
        same_initials.sort(key=_filing_sort_key)
        candidates.append(
            (
                _filing_sort_key(same_initials[0]),
                file_no,
            )
        )
    if not candidates:
        return []

    candidates.sort()
    selected_file_no = candidates[0][1]
    selected_rows = [r for r in rows if _file_number(r) == selected_file_no]
    chain = [r for r in selected_rows if _normalize_form(r) in _IPO_REGISTRATION_FORMS]
    final = _select_final_424b(selected_rows)
    if final:
        chain.append(final)

    return _sorted_unique_chain(chain)


def _select_ipo_chain_by_window(rows: list[dict]) -> list[dict]:
    initials = [r for r in rows if _normalize_form(r) in _IPO_INITIAL_FORMS]
    if not initials:
        return []

    initials.sort(key=_filing_sort_key)
    first_initial = initials[0]
    start_key = _filing_sort_key(first_initial)

    prospectuses_after_start = [
        r for r in rows if _is_424b(str(r.get("form") or "")) and _filing_sort_key(r) >= start_key
    ]
    prospectuses_after_start.sort(key=_filing_sort_key)
    first_424b = prospectuses_after_start[0] if prospectuses_after_start else None
    end_key = _filing_sort_key(first_424b) if first_424b else None

    chain = [
        r
        for r in rows
        if _normalize_form(r) in _IPO_REGISTRATION_FORMS
        and _filing_sort_key(r) >= start_key
        and (end_key is None or _filing_sort_key(r) <= end_key)
    ]
    if end_key is not None:
        final = _select_final_424b(
            [r for r in prospectuses_after_start if _filing_sort_key(r) <= end_key]
        )
        if final:
            chain.append(final)

    return _sorted_unique_chain(chain)


def _select_ipo_chain(rows: list[dict]) -> list[dict]:
    sorted_rows = sorted(rows, key=_filing_sort_key)
    file_number_chain = _select_ipo_chain_by_file_number(sorted_rows)
    if file_number_chain and _select_final_424b(file_number_chain):
        return file_number_chain

    window_chain = _select_ipo_chain_by_window(sorted_rows)
    if window_chain and _select_final_424b(window_chain):
        return window_chain

    return file_number_chain or window_chain


def _filing_label(row: dict) -> str:
    return _normalize_form(row).replace("/", "-").lower() or "filing"


def download_ipo_documents(ticker: Ticker, output_root: Path) -> list[ReportFile]:
    cik = ticker.external_id
    if not cik:
        raise ValueError(f"美股 IPO 下载需要 CIK：{ticker.code}")

    url = _SUBMISSIONS_URL.format(cik=cik)
    with default_client(source="sec") as client:
        payload = get_json(client, url, source="sec", rate=8.0)
        filings = payload.get("filings") or {}
        chain = _select_ipo_chain(_all_filing_rows(client, filings))

        if not chain:
            logger.warning(f"[{ticker.code}] no SEC IPO filings found")
            return []

        out: list[ReportFile] = []
        for sequence, row in enumerate(chain, start=1):
            form = _normalize_form(row)
            filing_date = row.get("filingDate") or None
            accession = _accession(row)
            dest = report_output_path_for_filing(
                output_root,
                ticker,
                "ipo",
                sequence,
                _filing_label(row),
                ".pdf",
                filing_date=filing_date,
                is_amendment=_is_amendment_form(form),
                source_id=accession.replace("-", "") if accession else None,
            )
            url_doc, source_format, n_bytes = _download_primary_as_pdf(client, cik, row, dest)
            title = row.get("primaryDocDescription") or f"{form} ({filing_date or 'unknown date'})"
            out.append(
                ReportFile(
                    ticker=ticker,
                    kind="ipo_prospectus",
                    local_path=str(dest),
                    source_url=url_doc,
                    title=title,
                    file_size_bytes=n_bytes,
                    filing_date=filing_date,
                    report_date=row.get("reportDate") or None,
                    form=form,
                    source_id=_file_number(row) or None,
                    accession_number=accession or None,
                    is_amendment=_is_amendment_form(form),
                    sequence=sequence,
                    source_format=source_format,
                    output_format="pdf",
                )
            )

        return out

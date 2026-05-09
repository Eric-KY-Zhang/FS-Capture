"""KR filings download via DART OpenAPI and disclosure pages."""
from __future__ import annotations

import datetime as dt
import re
from contextlib import contextmanager
from pathlib import Path
from typing import Optional

import pandas as pd
from loguru import logger

from app.core.http import default_client
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path, report_output_path_for_filing
from .name_resolver import _dart


_DART_RATE = 5.0
_DART_MAIN_URL = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}"
_DART_PDF_URL = "https://dart.fss.or.kr/pdf/download/pdf.do?rcp_no={rcept_no}&dcm_no={dcm_no}"
_DART_PDF_MAIN_URL = "https://dart.fss.or.kr/pdf/download/main.do?rcp_no={rcept_no}&dcm_no={dcm_no}"
_DART_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}

_REPORT_KEYWORDS = {
    PeriodType.ANNUAL: ("\uc0ac\uc5c5\ubcf4\uace0\uc11c",),
    PeriodType.Q1: ("\ubd84\uae30\ubcf4\uace0\uc11c",),
    PeriodType.Q2: ("\ubc18\uae30\ubcf4\uace0\uc11c",),
    PeriodType.Q3: ("\ubd84\uae30\ubcf4\uace0\uc11c",),
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

_IPO_KEEP_KEYWORDS = (
    "\uc99d\uad8c\uc2e0\uace0\uc11c",
    "\ud22c\uc790\uc124\uba85\uc11c",
    "\uc608\ube44\ud22c\uc790\uc124\uba85\uc11c",
)

_IPO_CHAIN_KEYWORDS = (
    "\uc815\uc815",
    "\uae30\uc7ac\uc815\uc815",
    "\uc815\uc815\uc2e0\uace0",
    "\ucd5c\uc885",
    "\ud655\uc815",
    "\ubc1c\ud589\uac00\uc561",
    "\ubc1c\ud589\uac00\uaca9",
    "\uacf5\ubaa8\uac00",
)

_IPO_EXCLUDE_KEYWORDS = (
    "\uc720\uc0c1\uc99d\uc790",
    "\uc8fc\uc8fc\ubc30\uc815",
    "\uc2e4\uad8c\uc8fc",
    "\uc2e4\uad8c",
    "\uccad\uc57d\uacb0\uacfc",
    "\ubc30\uc815",
    "\ubc30\uc815\uacb0\uacfc",
    "\ubc1c\ud589\uacb0\uacfc",
    "\uc99d\uad8c\ubc1c\ud589\uc2e4\uc801",
    "\ucd94\uac00\uc0c1\uc7a5",
    "\uc0c1\uc7a5\uc548\ub0b4",
    "\uc548\ub0b4\uacf5\uc2dc",
    "\uc548\ub0b4\ubb38",
    "\uacb0\uacfc\ubcf4\uace0",
    "\uc8fc\uc694\uc0ac\ud56d\ubcf4\uace0",
    "\uc0ac\uc5c5\ubcf4\uace0\uc11c",
    "\ubc18\uae30\ubcf4\uace0\uc11c",
    "\ubd84\uae30\ubcf4\uace0\uc11c",
)

_IPO_BOUNDARY_KEYWORDS = (
    "\uc99d\uad8c\ubc1c\ud589\uc2e4\uc801",
    "\ubc1c\ud589\uc2e4\uc801\ubcf4\uace0\uc11c",
    "\uacb0\uacfc\ubcf4\uace0\uc11c",
)

_IPO_WINDOW_DAYS = 366

_DART_PDF_RE = re.compile(
    r"openPdfDownload\(\s*['\"](?P<rcept>\d+)['\"]\s*,\s*['\"](?P<dcm>\d+)['\"]\s*\)"
)
_DART_DCM_RE = re.compile(r"(?:dcmNo|dcm_no)['\"=:\s]*(?P<dcm>\d+)")


@contextmanager
def _dart_client():
    try:
        from curl_cffi import requests as curl_requests

        session = curl_requests.Session(impersonate="chrome")
        session.headers.update(_DART_HEADERS)
        try:
            yield session
        finally:
            session.close()
    except ImportError:
        with default_client(source="dart", timeout=60.0) as client:
            yield client


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


def _list_ipo_filings(ticker: Ticker) -> pd.DataFrame:
    dart = _dart()
    frames: list[pd.DataFrame] = []
    for detail in ("C001", ""):
        try:
            df = dart.list(
                corp=_corp(ticker),
                kind="C",
                kind_detail=detail,
                final=False,
            )
        except Exception as exc:
            logger.warning(f"DART IPO list failed for {ticker.code}: {exc}")
            continue
        if isinstance(df, pd.DataFrame) and not df.empty:
            frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def _parse_date(value: str) -> Optional[dt.date]:
    try:
        text = str(value or "").strip()
        if len(text) == 8 and text.isdigit():
            return dt.date(int(text[:4]), int(text[4:6]), int(text[6:8]))
        return dt.date.fromisoformat(text)
    except (TypeError, ValueError):
        return None


def _date_text(row: dict) -> str:
    return str(row.get("rcept_dt") or row.get("rcept_de") or row.get("date") or "")


def _date_iso(row: dict) -> Optional[str]:
    text = _date_text(row)
    parsed = _parse_date(text)
    return parsed.isoformat() if parsed else (text or None)


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
        preferred = [r for r in rows if _date_month(r) in {4, 5, 6}]
        if preferred:
            rows = preferred
    elif period.type is PeriodType.Q3:
        preferred = [r for r in rows if _date_month(r) in {10, 11, 12}]
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
        if "\uac10\uc0ac\ubcf4\uace0\uc11c" in str(row.get(title_col, ""))
        and str(period.year) in str(row.get(title_col, ""))
    ]
    if not rows:
        return None
    rows.sort(key=_date_text)
    return rows[-1]


def _extract_dcm_no(html: str, rcept_no: str) -> Optional[str]:
    for match in _DART_PDF_RE.finditer(html or ""):
        if match.group("rcept") == rcept_no:
            return match.group("dcm")
    match = _DART_DCM_RE.search(html or "")
    return match.group("dcm") if match else None


def _is_pdf_file(path: Path) -> bool:
    try:
        with path.open("rb") as f:
            return f.read(4) == b"%PDF"
    except OSError:
        return False


def _download_pdf_url(client, url: str, dest: Path) -> Optional[int]:
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")
    tmp_dest.unlink(missing_ok=True)
    try:
        from app.core.ratelimit import limiter

        limiter("dart", _DART_RATE).acquire_blocking()
        response = client.get(url, timeout=120.0)
        response.raise_for_status()
        tmp_dest.write_bytes(response.content)
        if not _is_pdf_file(tmp_dest):
            tmp_dest.unlink(missing_ok=True)
            return None
        tmp_dest.replace(dest)
        return dest.stat().st_size
    except Exception as exc:
        tmp_dest.unlink(missing_ok=True)
        logger.warning(f"DART PDF download failed for {url}: {exc}")
        return None


def _render_url_to_pdf(url: str, dest: Path) -> Optional[int]:
    from app.core.pdf_renderer import render_url_to_pdf

    try:
        result = render_url_to_pdf(url, dest, source="dart")
        if not _is_pdf_file(dest):
            dest.unlink(missing_ok=True)
            return None
        if isinstance(result, int):
            return result
        if dest.exists():
            return dest.stat().st_size
    except Exception as exc:
        logger.warning(f"DART render fallback failed for {url}: {exc}")
    return None


def _download_rcept_as_pdf(rcept_no: str, dest: Path) -> tuple[str, str, int] | None:
    main_url = _DART_MAIN_URL.format(rcept_no=rcept_no)
    with _dart_client() as client:
        dcm_no: Optional[str] = None
        try:
            from app.core.ratelimit import limiter

            limiter("dart", _DART_RATE).acquire_blocking()
            response = client.get(main_url, timeout=60.0)
            response.raise_for_status()
            dcm_no = _extract_dcm_no(response.text, rcept_no)
        except Exception as exc:
            logger.warning(f"DART disclosure page fetch failed for rcept {rcept_no}: {exc}")

        if dcm_no:
            pdf_url = _DART_PDF_URL.format(rcept_no=rcept_no, dcm_no=dcm_no)
            n_bytes = _download_pdf_url(client, pdf_url, dest)
            if n_bytes is not None:
                return pdf_url, "pdf", n_bytes

            pdf_main_url = _DART_PDF_MAIN_URL.format(rcept_no=rcept_no, dcm_no=dcm_no)
            n_bytes = _download_pdf_url(client, pdf_main_url, dest)
            if n_bytes is not None:
                return pdf_main_url, "pdf", n_bytes

        n_bytes = _render_url_to_pdf(main_url, dest)
        if n_bytes is not None:
            return main_url, "html", n_bytes
    return None


def _report_file(ticker: Ticker, period: Period, row: dict, output_root: Path, kind: str) -> Optional[ReportFile]:
    rcept = str(row.get("rcept_no") or row.get("RCEPT_NO") or "")
    if not rcept:
        return None
    dest = report_output_path(output_root, ticker, period, kind, ".pdf")
    downloaded = _download_rcept_as_pdf(rcept, dest)
    if downloaded is None:
        return None
    source_url, source_format, n_bytes = downloaded
    return ReportFile(
        ticker=ticker,
        period=period,
        kind=kind,
        local_path=str(dest),
        source_url=source_url,
        title=str(row.get("report_nm") or row.get("title") or ""),
        file_size_bytes=n_bytes,
        filing_date=_date_iso(row),
        accession_number=rcept,
        source_format=source_format,
        output_format="pdf",
    )


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if period.type is PeriodType.IPO_PROSPECTUS:
        reports = download_ipo_documents(ticker, output_root)
        for report in reports:
            report.period = period
            report.kind = "ipo_prospectus"
        return reports

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


def _title(row: dict) -> str:
    return str(row.get("report_nm") or row.get("title") or "")


def _rcept(row: dict) -> str:
    return str(row.get("rcept_no") or row.get("RCEPT_NO") or "").strip()


def _is_ipo_document(row: dict) -> bool:
    title = _title(row)
    if not title:
        return False
    if any(keyword in title for keyword in _IPO_EXCLUDE_KEYWORDS):
        return False
    return any(keyword in title for keyword in _IPO_KEEP_KEYWORDS)


def _is_ipo_amendment(row: dict) -> bool:
    title = _title(row)
    return any(keyword in title for keyword in _IPO_CHAIN_KEYWORDS)


def _is_ipo_boundary(row: dict) -> bool:
    title = _title(row)
    return any(keyword in title for keyword in _IPO_BOUNDARY_KEYWORDS)


def _dedupe_ipo_rows(rows: list[dict]) -> list[dict]:
    seen: set[str] = set()
    out: list[dict] = []
    for row in rows:
        key = _rcept(row) or f"{_date_text(row)}:{_title(row)}"
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def _select_ipo_rows(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []
    rows = [row.to_dict() for _, row in df.iterrows()]
    rows.sort(key=lambda row: (_date_text(row), _rcept(row), _title(row)))
    selected = [row for row in rows if _is_ipo_document(row)]
    if not selected:
        return []

    start = selected[0]
    start_date = _parse_date(_date_text(start))
    if start_date is not None:
        result_dates = [
            parsed
            for row in rows
            if (parsed := _parse_date(_date_text(row))) is not None
            and parsed >= start_date
            and _is_ipo_boundary(row)
        ]
        boundary_date = min(result_dates) if result_dates else start_date + dt.timedelta(days=_IPO_WINDOW_DAYS)
        selected = [
            row for row in selected
            if (parsed := _parse_date(_date_text(row))) is not None
            and start_date <= parsed <= boundary_date
        ]

    selected = _dedupe_ipo_rows(selected)
    selected.sort(key=lambda row: (_date_text(row), _rcept(row), _title(row)))
    return selected


def _filing_label(row: dict) -> str:
    title = _title(row)
    if "\uc99d\uad8c\uc2e0\uace0\uc11c" in title:
        return "securities_registration"
    if "\uc608\ube44\ud22c\uc790\uc124\uba85\uc11c" in title:
        return "preliminary_prospectus"
    if "\ud22c\uc790\uc124\uba85\uc11c" in title:
        return "prospectus"
    return "ipo_document"


def download_ipo_documents(ticker: Ticker, output_root: Path) -> list[ReportFile]:
    rows = _select_ipo_rows(_list_ipo_filings(ticker))
    if not rows:
        logger.warning(f"[{ticker.code}] no DART IPO documents found")
        return []

    out: list[ReportFile] = []
    for sequence, row in enumerate(rows, start=1):
        rcept = _rcept(row)
        if not rcept:
            continue
        filing_date = _date_iso(row)
        is_amendment = _is_ipo_amendment(row)
        dest = report_output_path_for_filing(
            output_root,
            ticker,
            "ipo",
            sequence,
            _filing_label(row),
            ".pdf",
            filing_date=filing_date,
            is_amendment=is_amendment,
            source_id=rcept,
        )
        downloaded = _download_rcept_as_pdf(rcept, dest)
        if downloaded is None:
            continue
        source_url, source_format, n_bytes = downloaded
        out.append(ReportFile(
            ticker=ticker,
            period=None,
            kind="ipo_prospectus",
            local_path=str(dest),
            source_url=source_url,
            title=_title(row),
            file_size_bytes=n_bytes,
            filing_date=filing_date,
            form="DART-C001",
            source_id=rcept,
            accession_number=rcept,
            is_amendment=is_amendment,
            sequence=sequence,
            source_format=source_format,
            output_format="pdf",
        ))

    return out

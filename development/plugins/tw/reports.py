"""TW (台股) report download via MOPS / doc.twse.com.tw.

Single endpoint covers everything:

    GET https://doc.twse.com.tw/server-java/t57sb01?step=1
        &co_id=<4-digit>&year=<ROC>&mtype=<F|A|B>

Filename grammars differ by mtype:

- **mtype=F** (股東會 / shareholder meeting; the annual report lives here):
  ``{meeting_AD_year}_{co_id}_{YYYYMMDD}{typecode}.pdf``
  typecode ``F04`` = 中文年報, ``FE4`` = English annual report. The filename's
  leading 4-digit year is the **data year (FY)**, NOT the meeting year. Query
  ``year=ROC(FY+1)`` (i.e. the meeting year after the fiscal year ends).

- **mtype=A** (財務報告書 / financial statements; quarterly + annual):
  ``{YYYY}{Q}_{co_id}_{typecode}.pdf``
  where ``YYYY`` is the western fiscal year and ``Q`` is ``01``/``02``/``03``/
  ``04`` for Q1/H1/Q3/annual. typecode ``AI1`` = 合併中文 (preferred),
  ``AIA`` = 個體 (parent-only fallback). Query ``year=ROC(FY)``.

- **mtype=B** (公開說明書 / IPO prospectus): older filings; sweep ROC years.

ROC year conversion: ``roc = ad - 1911`` (西元 2024 → 民國 113).
"""

from __future__ import annotations

import datetime as dt
import re
import time
from pathlib import Path

from loguru import logger

from app.core.http import default_client, stream_to_file
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path, report_output_path_for_filing
from app.core.ratelimit import limiter

_DOC_URL = "https://doc.twse.com.tw/server-java/t57sb01"
_RATE = 2.0

_KIND = {
    PeriodType.ANNUAL: "annual_report",
    PeriodType.Q1: "q1_report",
    PeriodType.Q2: "interim_report",
    PeriodType.Q3: "q3_report",
}

# Quarter suffix used in mtype=A filenames: YYYY + suffix (e.g. 202401 = 2024 Q1).
_QUARTER_SUFFIX = {
    PeriodType.Q1: "01",
    PeriodType.Q2: "02",
    PeriodType.Q3: "03",
}

# Filename parsers, one per mtype.
# mtype=F  → 2024_2330_20250603F04.pdf  → (yyyy=2024, co=2330, date=20250603, type=F04)
_FNAME_F_RE = re.compile(
    r"(?P<fname>(?P<yyyy>\d{4})_(?P<co>\d{4,6})_(?P<date>\d{8})(?P<type>[A-Z]{1,2}\d{1,2}))\.pdf",
    re.IGNORECASE,
)
# mtype=A  → 202401_2330_AI1.pdf  → (yyyy=2024, qq=01, co=2330, type=AI1)
_FNAME_A_RE = re.compile(
    r"(?P<fname>(?P<yyyy>\d{4})(?P<qq>0[1-4])_(?P<co>\d{4,6})_(?P<type>[A-Z]{2,3}\d?))\.pdf",
    re.IGNORECASE,
)
# mtype=B  (公開說明書) — filename varies; lenient match.
_FNAME_B_RE = re.compile(
    r"(?P<fname>(?P<prefix>[\dA-Z]+)_(?P<co>\d{4,6})_(?P<type>[A-Z0-9]+))\.pdf",
    re.IGNORECASE,
)

# Also accept the function-call form: readfile2("F","2330","2024_2330_20250603F04.pdf")
_READFILE_RE = re.compile(
    r"readfile2?\(\s*['\"](?P<mtype>[A-Z])['\"]\s*,\s*['\"](?P<co>\d{4,6})['\"]"
    r"\s*,\s*['\"](?P<fname>[^'\"]+)['\"]\s*\)",
    re.IGNORECASE,
)

# Soft rate limit: MOPS sometimes returns 200 with no rows.
_EMPTY_RETRIES = 3
_EMPTY_BACKOFF_BASE = 1.5


def roc_year(ad_year: int) -> int:
    """Western year → ROC year (民國紀年)."""
    if ad_year < 1912:
        raise ValueError(f"ROC year is undefined before 1912 (got {ad_year})")
    return ad_year - 1911


def _fetch_listing(co_id: str, roc: int, mtype: str) -> str:
    """GET the t57sb01 listing page. Returns Big5-decoded HTML body.

    Returns "" if the endpoint persistently yields an empty filename listing.
    """
    params = {
        "step": "1",
        "colorchg": "1",
        "co_id": co_id,
        "year": str(roc),
        "mtype": mtype,
    }
    last_body = ""
    with default_client(source="twse", timeout=60.0) as client:
        for attempt in range(_EMPTY_RETRIES):
            limiter("twse", _RATE).acquire_blocking()
            try:
                resp = client.get(_DOC_URL, params=params)
                resp.raise_for_status()
            except Exception as exc:
                logger.warning(
                    f"TWSE listing fetch failed ({co_id} year={roc} mtype={mtype}): {exc}"
                )
                time.sleep(_EMPTY_BACKOFF_BASE**attempt)
                continue
            body = resp.content.decode("big5", errors="replace")
            last_body = body
            if _READFILE_RE.search(body):
                return body
            logger.debug(
                f"TWSE empty listing on attempt {attempt + 1}/{_EMPTY_RETRIES} for "
                f"{co_id} roc_year={roc} mtype={mtype}"
            )
            time.sleep(_EMPTY_BACKOFF_BASE**attempt)
    return last_body


def _extract_pdf_links(html: str, mtype: str) -> list[dict]:
    """Extract every PDF filename + its inferred metadata. Dedup, preserve order."""
    rows: list[dict] = []
    seen: set[str] = set()
    fname_re = {"F": _FNAME_F_RE, "A": _FNAME_A_RE, "B": _FNAME_B_RE}[mtype]
    for call in _READFILE_RE.finditer(html or ""):
        if call.group("mtype").upper() != mtype.upper():
            continue
        fname = call.group("fname")
        if fname in seen:
            continue
        m = fname_re.match(fname)
        if not m:
            continue
        seen.add(fname)
        info: dict = {"filename": fname, "typecode": m.group("type").upper()}
        gd = m.groupdict()
        if "yyyy" in gd and gd["yyyy"]:
            info["yyyy"] = gd["yyyy"]
        if "qq" in gd and gd["qq"]:
            info["qq"] = gd["qq"]
        if "date" in gd and gd["date"]:
            info["filedate"] = gd["date"]
        rows.append(info)
    return rows


def _annual_score(typecode: str) -> int:
    """Lower is better for annual (mtype=F). Prefer F04 (Chinese annual),
    fall back to FE4 (English). Reject F01/F02/F05/etc (meeting notices, minutes).
    """
    t = typecode.upper()
    if t == "F04":
        return 0
    if t == "FE4":
        return 1
    return 99  # not an annual report


def _quarter_score(typecode: str) -> int:
    """Lower is better for quarterly (mtype=A). Prefer AI1 (consolidated Chinese),
    fall back to AIA (parent-only), then AE1 (English consolidated)."""
    t = typecode.upper()
    if t == "AI1":
        return 0
    if t == "AIA":
        return 1
    if t == "AE1":
        return 2
    if t.startswith("AI"):
        return 3
    return 99


def _select_annual(rows: list[dict], target_fy: int) -> dict | None:
    """For mtype=F: filename's leading 4-digit YYYY is the data year (FY).
    Pick F04/FE4 entries with yyyy == target_fy.
    """
    cands = [
        r for r in rows if r.get("yyyy") == str(target_fy) and _annual_score(r["typecode"]) < 99
    ]
    if not cands:
        return None
    cands.sort(key=lambda r: (_annual_score(r["typecode"]), r["filename"]))
    return cands[0]


def _select_quarter(rows: list[dict], target_fy: int, period_type: PeriodType) -> dict | None:
    """For mtype=A: pick {YYYY}{QQ} prefix matching target year + quarter."""
    qq = _QUARTER_SUFFIX[period_type]
    cands = [
        r
        for r in rows
        if r.get("yyyy") == str(target_fy)
        and r.get("qq") == qq
        and _quarter_score(r["typecode"]) < 99
    ]
    if not cands:
        return None
    cands.sort(key=lambda r: (_quarter_score(r["typecode"]), r["filename"]))
    return cands[0]


def _build_pdf_url(co_id: str, filename: str, mtype: str) -> str:
    """Construct the step=9 'launch' URL — recorded as ``source_url`` even though
    the actual binary is fetched in a second hop. Kept GET-shaped for readability;
    the real download goes through ``_download_pdf`` which POSTs the form.
    """
    return f"{_DOC_URL}?step=9&kind={mtype}&co_id={co_id}&filename={filename}&colorchg=1"


_PDF_HREF_RE = re.compile(r"href=['\"](/pdf/[^'\"]+\.pdf)['\"]", re.IGNORECASE)


def _download_pdf(url: str, dest: Path) -> int | None:
    """Two-step download:

    1. POST ``step=9`` with the form fields the JS handler submits → returns a
       small HTML page containing ``<a href='/pdf/<filename>_<timestamp>.pdf'>``.
    2. GET that temporary URL → the actual PDF bytes (session-bound for ~minutes).
    """
    # Parse mtype + co_id + filename out of the launch URL we just built.
    co_id = _qparam(url, "co_id")
    fname = _qparam(url, "filename")
    mtype = _qparam(url, "kind")
    if not (co_id and fname and mtype):
        logger.warning(f"TWSE _download_pdf: malformed launch URL {url}")
        return None

    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        with default_client(source="twse", timeout=120.0) as client:
            # Step 1: POST to get the temporary PDF URL.
            limiter("twse", _RATE).acquire_blocking()
            form = {
                "step": "9",
                "kind": mtype,
                "co_id": co_id,
                "filename": fname,
                "colorchg": "1",
                "DEBUG": "",
                "SKEY1": "",
                "SKEY2": "",
                "YEAR": "",
                "MDATE": "",
                "TYPE": "",
            }
            resp = client.post(_DOC_URL, data=form, timeout=60.0)
            resp.raise_for_status()
            body = resp.content.decode("big5", errors="replace")
            href_match = _PDF_HREF_RE.search(body)
            if not href_match:
                logger.warning(
                    f"TWSE step=9 did not yield a /pdf/ link for {fname}: body[:200]={body[:200]!r}"
                )
                return None
            pdf_url = f"https://doc.twse.com.tw{href_match.group(1)}"

            # Step 2: GET the temporary PDF.
            n = stream_to_file(
                client,
                pdf_url,
                dest,
                source="twse",
                rate=_RATE,
                read_timeout=180.0,
            )
        with dest.open("rb") as f:
            head = f.read(4)
        if head != b"%PDF":
            logger.warning(f"TWSE final download was not a PDF (got {head!r}): {pdf_url}")
            dest.unlink(missing_ok=True)
            return None
        return n
    except Exception as exc:
        logger.warning(f"TWSE download failed for {url}: {exc}")
        return None


def _qparam(url: str, name: str) -> str | None:
    """Cheap query-string param extraction (no urllib parsing — handles our own URLs)."""
    m = re.search(rf"[?&]{re.escape(name)}=([^&]+)", url)
    return m.group(1) if m else None


def _filing_date(row: dict) -> str | None:
    """Best-effort filing date from filename fields."""
    if row.get("filedate") and len(row["filedate"]) == 8 and row["filedate"].isdigit():
        d = row["filedate"]
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    if row.get("yyyy") and row.get("qq"):
        return f"{row['yyyy']}-{row['qq']}-01"
    return None


def _download_annual(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    """ANNUAL report lives in mtype=F (shareholder meeting materials).

    The annual report for FY=X is filed at the meeting in calendar year X+1
    → query at ROC(X+1). The filename's leading YYYY is the FY (X).
    """
    meeting_roc = roc_year(period.year + 1)
    html = _fetch_listing(ticker.code, meeting_roc, "F")
    rows = _extract_pdf_links(html, "F")
    chosen = _select_annual(rows, period.year)
    if chosen is None:
        # Fallback: many companies file the annual report at the same-year meeting
        # if FY differs. Try ROC(X) too.
        fallback_html = _fetch_listing(ticker.code, roc_year(period.year), "F")
        fallback_rows = _extract_pdf_links(fallback_html, "F")
        chosen = _select_annual(fallback_rows, period.year)
        if chosen is None:
            logger.warning(
                f"[{ticker.code}] no TW {period.label()} annual report (F04/FE4) "
                f"found at ROC {meeting_roc} or {roc_year(period.year)}"
            )
            return []
    return _materialize(ticker, period, output_root, chosen, "F")


def _download_quarter(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    """Q1 / Q2 / Q3 financial statements live in mtype=A."""
    html = _fetch_listing(ticker.code, roc_year(period.year), "A")
    rows = _extract_pdf_links(html, "A")
    chosen = _select_quarter(rows, period.year, period.type)
    if chosen is None:
        logger.warning(
            f"[{ticker.code}] no TW {period.label()} financial statement (AI1/AIA) found"
        )
        return []
    return _materialize(ticker, period, output_root, chosen, "A")


def _materialize(
    ticker: Ticker, period: Period, output_root: Path, row: dict, mtype: str
) -> list[ReportFile]:
    kind = _KIND[period.type]
    dest = report_output_path(output_root, ticker, period, kind, ".pdf")
    url = _build_pdf_url(ticker.code, row["filename"], mtype)
    n_bytes = _download_pdf(url, dest)
    if n_bytes is None:
        return []
    return [
        ReportFile(
            ticker=ticker,
            period=period,
            kind=kind,
            local_path=str(dest),
            source_url=url,
            title=row["filename"],
            file_size_bytes=n_bytes,
            filing_date=_filing_date(row),
            source_id=row["filename"],
            source_format="pdf",
            output_format="pdf",
        )
    ]


def _download_ipo(ticker: Ticker, output_root: Path) -> list[ReportFile]:
    """Sweep mtype=B across recent ROC years. IPO prospectuses are rare per company."""
    seen: set[str] = set()
    rows: list[dict] = []
    current_roc = roc_year(dt.datetime.now().year)
    # Sweep current..current-30 ROC years (covers most listed-since-1992 issuers).
    for r in range(current_roc, max(current_roc - 30, 80), -1):
        try:
            html = _fetch_listing(ticker.code, r, "B")
        except Exception as exc:
            logger.warning(f"TWSE IPO sweep failed at ROC {r}: {exc}")
            continue
        for row in _extract_pdf_links(html, "B"):
            if row["filename"] in seen:
                continue
            seen.add(row["filename"])
            rows.append(row)

    if not rows:
        logger.warning(f"[{ticker.code}] no TW IPO prospectus found on MOPS")
        return []

    rows.sort(key=lambda r: r["filename"])
    out: list[ReportFile] = []
    for sequence, row in enumerate(rows, start=1):
        filing_date = _filing_date(row)
        dest = report_output_path_for_filing(
            output_root,
            ticker,
            "ipo",
            sequence,
            "ipo_prospectus",
            ".pdf",
            filing_date=filing_date,
            source_id=row["filename"],
        )
        url = _build_pdf_url(ticker.code, row["filename"], "B")
        n_bytes = _download_pdf(url, dest)
        if n_bytes is None:
            continue
        out.append(
            ReportFile(
                ticker=ticker,
                period=None,
                kind="ipo_prospectus",
                local_path=str(dest),
                source_url=url,
                title=row["filename"],
                file_size_bytes=n_bytes,
                filing_date=filing_date,
                form="MOPS-B",
                source_id=row["filename"],
                sequence=sequence,
                source_format="pdf",
                output_format="pdf",
            )
        )
    return out


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if period.type is PeriodType.IPO_PROSPECTUS:
        return _download_ipo(ticker, output_root)
    if period.type is PeriodType.ANNUAL:
        return _download_annual(ticker, period, output_root)
    if period.type in (PeriodType.Q1, PeriodType.Q2, PeriodType.Q3):
        return _download_quarter(ticker, period, output_root)
    logger.warning(f"[{ticker.code}] unsupported TW period type {period.type}")
    return []

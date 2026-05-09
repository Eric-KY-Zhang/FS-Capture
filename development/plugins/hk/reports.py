"""HK report and IPO document download via HKEXnews Title Search."""
from __future__ import annotations

import datetime as dt
import html
import re
from pathlib import Path
from typing import Optional
from urllib.parse import urljoin, urlparse

from loguru import logger
from selectolax.parser import HTMLParser

from app.core.http import default_client, stream_to_file
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path
from app.core.ratelimit import limiter
from ._pdf_verify import verify_pdf_year_and_kind
from .fiscal_year import fiscal_year_end_month

try:
    from app.core.output_paths import report_output_path_for_filing as _report_output_path_for_filing
except ImportError:  # Main agent adds this public helper; keep HK importable meanwhile.
    _report_output_path_for_filing = None


_TITLESEARCH_URL = "https://www1.hkexnews.hk/search/titlesearch.xhtml"
_PDF_HOST = "https://www1.hkexnews.hk"
_HKEX_RATE = 3.0
_TITLE_SEARCH_START = "19990401"

_STOCK_CODE_RE = re.compile(r"\b\d{5}\b")
_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# HKEXnews Title Search headline category codes.
_DOC_TYPE_BY_PERIOD = {
    PeriodType.ANNUAL: ("40000", "40100", "annual_report"),
    PeriodType.Q2: ("40000", "40200", "interim_report"),
    PeriodType.Q1: ("40000", "40400", "q1_report"),
    PeriodType.Q3: ("40000", "40300", "q3_report"),
}

_REPORT_KEYWORDS = {
    PeriodType.ANNUAL: ("annual report", "年報", "年度報告", "年报", "年度报告"),
    PeriodType.Q1: ("first quarterly report", "first quarter report", "第一季度報告", "第一季度报告"),
    PeriodType.Q2: ("interim report", "half-year report", "interim/half-year report", "中期報告", "中期报告", "半年報", "半年报"),
    PeriodType.Q3: ("third quarterly report", "third quarter report", "第三季度報告", "第三季度报告"),
}

_REPORT_EXCLUDE_KEYWORDS = (
    "esg",
    "environmental, social and governance",
    "sustainability",
    "notification",
    "letter",
    "annual results",
    "interim results",
    "quarterly results",
    "circular",
    "proxy form",
    "form of proxy",
    "announcement",
    "環境、社會及管治",
    "环境、社会及管治",
    "通知",
    "通函",
    "業績",
    "业绩",
    "補充",
    "补充",
    "更正",
    "勘誤",
    "勘误",
    "supplementary",
)

_IPO_PRIMARY_DOC_TYPES = (
    "offer for subscription",
    "offer for sale",
    "introduction",
    "placing of securities of a class new to listing",
    "authorised collective investment scheme",
    "deemed new listing",
)

_IPO_PRIMARY_TITLE_KEYWORDS = (
    "prospectus",
    "listing document",
    "global offering",
    "initial public offering",
    "ipo",
    "招股章程",
    "上市文件",
)

_IPO_SUPPLEMENT_KEYWORDS = (
    "supplement",
    "supplemental",
    "supplementary",
    "addendum",
    "revised",
    "replacement",
    "補充",
    "修訂",
    "补充",
    "修订",
)

_IPO_FALLBACK_KEYWORDS = (
    "application proof",
    "phip",
    "post hearing information pack",
    "申請版本",
    "申请版本",
    "聆訊後資料集",
    "聆讯后资料集",
)

_IPO_EXCLUDE_KEYWORDS = (
    "oc announcement",
    "overall coordinator announcement",
    "warning statement",
    "formal notice",
    "allotment results",
    "announcement",
    "application form",
)


def _squash(text: str) -> str:
    return " ".join(html.unescape(text or "").split())


def _node_text(node) -> str:
    if node is None:
        return ""
    try:
        return _squash(node.text(separator=" ", strip=True))
    except TypeError:
        return _squash(node.text(strip=True))


def _strip_mobile_heading(text: str) -> str:
    return re.sub(r"^(Release Time|Stock Code|Stock Short Name|Document):\s*", "", text).strip()


def _parse_release_date(value: str) -> Optional[dt.date]:
    value = _strip_mobile_heading(value)
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def _source_id_from_url(url: str) -> str:
    stem = Path(urlparse(url).path).stem
    return stem or url


def _suffix_from_url(url: str) -> str:
    suffix = Path(urlparse(url).path).suffix.lower()
    return suffix or ".pdf"


def _format_from_suffix(suffix: str) -> str:
    value = suffix.lower().lstrip(".")
    if value in {"htm", "html"}:
        return "html"
    return value or "pdf"


def _extract_doc_type(headline: str) -> str:
    match = re.search(r"\[([^\]]+)\]", headline or "")
    if match:
        return _squash(match.group(1))
    if " - " in headline:
        return _squash(headline.rsplit(" - ", 1)[-1])
    return _squash(headline)


def _parse_results(html_text: str) -> list[dict]:
    """Extract Title Search rows with URL, title, headline and stock codes."""
    tree = HTMLParser(html_text)
    rows: list[dict] = []
    trs = tree.css("table.table-scroll tbody tr") or tree.css("table.table-scroll tr")
    for tr in trs:
        cells = tr.css("td")
        if len(cells) < 4:
            continue

        date_txt = _strip_mobile_heading(_node_text(cells[0]))
        stock_text = _strip_mobile_heading(_node_text(cells[1]))
        name_text = _strip_mobile_heading(_node_text(cells[2]))
        doc_cell = cells[3]
        link = doc_cell.css_first("div.doc-link a") or doc_cell.css_first("a")
        if link is None:
            continue
        href = link.attributes.get("href") or ""
        if not href:
            continue

        headline = _node_text(doc_cell.css_first("div.headline"))
        url = href if href.startswith(("http://", "https://")) else urljoin(_PDF_HOST, href)
        title = _node_text(link)
        rows.append({
            "url": url,
            "title": title,
            "headline": headline,
            "doc_type": _extract_doc_type(headline),
            "date": date_txt,
            "filing_date": _parse_release_date(date_txt),
            "stock_codes": tuple(_STOCK_CODE_RE.findall(stock_text)),
            "stock_names": name_text,
            "source_id": _source_id_from_url(url),
        })
    return rows


def _require_stock_id(ticker: Ticker) -> str:
    if not ticker.external_id:
        raise ValueError(f"HKEX download requires ticker.external_id stockId for {ticker.code}")
    return str(ticker.external_id)


def _title_search_params(
    ticker: Ticker,
    *,
    from_date: str,
    to_date: str,
    t1code: str,
    t2code: str = "-2",
    t2gcode: str = "-2",
    title: str = "",
) -> dict:
    return {
        "lang": "EN",
        "category": "0",
        "market": "SEHK",
        "searchType": "1",
        "documentType": "-1",
        "t1code": t1code,
        "t2Gcode": t2gcode,
        "t2code": t2code,
        "stockId": _require_stock_id(ticker),
        "from": from_date,
        "to": to_date,
        "MB-Daterange": "0",
        "title": title,
    }


def _post_title_search(client, params: dict) -> str:
    headers = {
        "Origin": "https://www1.hkexnews.hk",
        "Referer": "https://www1.hkexnews.hk/search/titlesearch.xhtml?lang=en",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    limiter("hkexnews", _HKEX_RATE).acquire_blocking()
    response = client.post(_TITLESEARCH_URL, data=params, headers=headers)
    response.raise_for_status()
    return response.text


def _search_rows(
    client,
    ticker: Ticker,
    *,
    from_date: str,
    to_date: str,
    t1code: str,
    t2code: str = "-2",
    t2gcode: str = "-2",
    title: str = "",
) -> list[dict]:
    params = _title_search_params(
        ticker,
        from_date=from_date,
        to_date=to_date,
        t1code=t1code,
        t2code=t2code,
        t2gcode=t2gcode,
        title=title,
    )
    return _parse_results(_post_title_search(client, params))


def _row_matches_ticker(row: dict, ticker: Ticker) -> bool:
    codes = {str(code).zfill(5) for code in row.get("stock_codes") or ()}
    return ticker.code.zfill(5) in codes


def _contains_any(text: str, keywords: tuple[str, ...]) -> bool:
    lowered = text.lower()
    return any(keyword.lower() in lowered for keyword in keywords)


def _is_report_candidate(row: dict, period: Period) -> bool:
    title = row.get("title") or ""
    doc_type = row.get("doc_type") or ""

    if _contains_any(title, _REPORT_EXCLUDE_KEYWORDS):
        return False

    keywords = _REPORT_KEYWORDS[period.type]
    if _contains_any(title, keywords):
        return True
    return _contains_any(doc_type, keywords)


def _period_verify_kind(period_type: PeriodType) -> str:
    return {
        PeriodType.ANNUAL: "annual",
        PeriodType.Q1: "q1",
        PeriodType.Q2: "interim",
        PeriodType.Q3: "q3",
    }[period_type]


def _add_months(value: dt.date, months: int) -> dt.date:
    month_index = value.month - 1 + months
    year = value.year + month_index // 12
    month = month_index % 12 + 1
    return dt.date(year, month, 1)


def _filing_window(period: Period, fye_month: int) -> tuple[dt.date, dt.date]:
    if fye_month == 12:
        start = dt.date(period.year + 1, 1, 1)
    else:
        start = _add_months(dt.date(period.year, fye_month, 1), 1)
    end = _add_months(start, 6) - dt.timedelta(days=1)
    return start, end


def _filing_in_expected_window(filing_date: Optional[dt.date], period: Period, ticker: Ticker) -> bool:
    if filing_date is None:
        return False
    start, end = _filing_window(period, fiscal_year_end_month(ticker.code))
    return start <= filing_date <= end


def _file_size(row: dict) -> int:
    value = row.get("file_size_bytes", row.get("file_size"))
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _title_has_year_and_kind(title: str, period: Period) -> bool:
    return str(period.year) in title and _contains_any(title, _REPORT_KEYWORDS[period.type])


def _base_report_score(row: dict, ticker: Ticker, period: Period) -> int:
    title = row.get("title") or ""
    score = 0
    if _title_has_year_and_kind(title, period):
        score += 20
    if _filing_in_expected_window(row.get("filing_date"), period, ticker):
        score += 5
    if _file_size(row) > 1_000_000:
        score += 3
    if _contains_any(title, ("補充", "补充", "更正", "勘誤", "勘误", "supplementary")):
        score -= 30
    return score


def _select_main(rows: list[dict], ticker: Ticker, period: Period) -> Optional[dict]:
    matching_rows = [row for row in rows if _row_matches_ticker(row, ticker)]
    if not matching_rows:
        return None

    year_str = str(period.year)
    candidates = [
        row for row in matching_rows
        if _is_report_candidate(row, period) and year_str in (row.get("title") or "")
    ]
    if not candidates:
        candidates = [row for row in matching_rows if _is_report_candidate(row, period)]
    if not candidates:
        return None

    scored: list[tuple[int, dict]] = [
        (_base_report_score(row, ticker, period), row) for row in candidates
    ]

    if len(candidates) > 1:
        top = sorted(
            scored,
            key=lambda item: (
                -item[0],
                item[1].get("filing_date") or dt.date.max,
                item[1].get("source_id") or "",
            ),
        )[:5]
        verified_ids: set[int] = set()
        verify_kind = _period_verify_kind(period.type)
        for score, row in top:
            if verify_pdf_year_and_kind(row.get("url") or "", period.year, verify_kind):
                verified_ids.add(id(row))
        scored = [
            (score + (10 if id(row) in verified_ids else 0), row)
            for score, row in scored
        ]

    scored.sort(
        key=lambda item: (
            -item[0],
            item[1].get("filing_date") or dt.date.max,
            item[1].get("source_id") or "",
        )
    )
    return scored[0][1]


def _fallback_safe_filename(value: str, *, fallback: str = "file", max_len: int = 120) -> str:
    cleaned = _BAD_FILENAME_CHARS.sub("_", value).strip(" ._")
    if not cleaned:
        cleaned = fallback
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" ._")
    return cleaned or fallback


def _filing_output_path(
    output_root: Path,
    ticker: Ticker,
    *,
    sequence: int,
    label: str,
    filing_date: Optional[dt.date],
    is_amendment: bool,
    suffix: str,
    source_id: Optional[str],
) -> Path:
    filing_date_text = filing_date.isoformat() if filing_date else None
    path_label = _fallback_safe_filename(label, fallback="ipo_document")
    if _report_output_path_for_filing is not None:
        return _report_output_path_for_filing(
            output_root,
            ticker,
            "ipo",
            sequence,
            path_label,
            suffix,
            filing_date=filing_date_text,
            is_amendment=is_amendment,
            source_id=source_id,
        )

    date_part = filing_date_text or "undated"
    amend_part = "_amendment" if is_amendment else ""
    filename = _fallback_safe_filename(
        f"{ticker.exchange.value}_{ticker.code}_ipo_{sequence:03d}_{date_part}{amend_part}_{path_label}{suffix}",
        fallback=f"{ticker.exchange.value}_{ticker.code}_ipo_{sequence:03d}{suffix}",
        max_len=180,
    )
    return output_root / filename


def _make_report_file(**kwargs) -> ReportFile:
    try:
        return ReportFile(**kwargs)
    except Exception:
        if kwargs.get("period") is None and hasattr(ReportFile, "model_construct"):
            return ReportFile.model_construct(**kwargs)
        raise


def _is_ipo_amendment(row: dict) -> bool:
    text = " ".join((row.get("title") or "", row.get("doc_type") or "", row.get("headline") or ""))
    return _contains_any(text, _IPO_SUPPLEMENT_KEYWORDS)


def _is_formal_ipo_primary(row: dict) -> bool:
    title = row.get("title") or ""
    doc_type = row.get("doc_type") or ""
    headline = row.get("headline") or ""
    full = " ".join((title, doc_type, headline))
    if _contains_any(full, _IPO_FALLBACK_KEYWORDS) or _contains_any(full, _IPO_EXCLUDE_KEYWORDS):
        return False
    if _contains_any(doc_type, _IPO_PRIMARY_DOC_TYPES):
        return True
    return "listing documents" in headline.lower() and _contains_any(full, _IPO_PRIMARY_TITLE_KEYWORDS)


def _is_formal_ipo_supplement(row: dict) -> bool:
    title = row.get("title") or ""
    doc_type = row.get("doc_type") or ""
    headline = row.get("headline") or ""
    full = " ".join((title, doc_type, headline))
    if _contains_any(full, _IPO_FALLBACK_KEYWORDS) or _contains_any(full, _IPO_EXCLUDE_KEYWORDS):
        return False
    if not _contains_any(full, _IPO_SUPPLEMENT_KEYWORDS):
        return False
    return "listing documents" in headline.lower() or "listing document" in doc_type.lower()


def _is_ipo_fallback(row: dict) -> bool:
    full = " ".join((row.get("title") or "", row.get("doc_type") or "", row.get("headline") or ""))
    if _contains_any(full, _IPO_EXCLUDE_KEYWORDS):
        return False
    return _contains_any(full, _IPO_FALLBACK_KEYWORDS)


def _near_primary_date(row: dict, primary_dates: list[dt.date]) -> bool:
    filing_date = row.get("filing_date")
    if not filing_date:
        return True
    return any(abs((filing_date - primary_date).days) <= 365 for primary_date in primary_dates)


def _dedupe_rows(rows: list[dict]) -> list[dict]:
    seen: set[tuple[str, str]] = set()
    out: list[dict] = []
    for row in rows:
        key = (row.get("source_id") or "", row.get("url") or "")
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def _sort_filing_rows(rows: list[dict]) -> list[dict]:
    return sorted(
        rows,
        key=lambda row: (
            row.get("filing_date") or dt.date.min,
            row.get("source_id") or "",
            row.get("title") or "",
        ),
    )


def _select_ipo_rows(formal_rows: list[dict], fallback_rows: list[dict], ticker: Ticker) -> list[dict]:
    formal_rows = [row for row in formal_rows if _row_matches_ticker(row, ticker)]
    primary_rows = [row for row in formal_rows if _is_formal_ipo_primary(row)]
    supplement_rows = [row for row in formal_rows if _is_formal_ipo_supplement(row)]

    selected: list[dict] = []
    if primary_rows:
        primary_dates = [row["filing_date"] for row in primary_rows if row.get("filing_date")]
        selected = primary_rows + [
            row for row in supplement_rows
            if not primary_dates or _near_primary_date(row, primary_dates)
        ]
    else:
        selected = [
            row for row in supplement_rows
            if _contains_any(row.get("title") or "", _IPO_PRIMARY_TITLE_KEYWORDS)
        ]

    selected = _sort_filing_rows(_dedupe_rows(selected))
    if selected:
        return selected

    fallback = [
        row for row in fallback_rows
        if _row_matches_ticker(row, ticker) and _is_ipo_fallback(row)
    ]
    return _sort_filing_rows(_dedupe_rows(fallback))


def _ipo_form(row: dict) -> str:
    doc_type = row.get("doc_type") or ""
    full = " ".join((row.get("title") or "", doc_type, row.get("headline") or ""))
    if _contains_any(full, _IPO_FALLBACK_KEYWORDS):
        return "PHIP/Application Proof"
    if _contains_any(full, _IPO_SUPPLEMENT_KEYWORDS):
        return "IPO Supplement"
    return doc_type or "IPO"


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if period.type not in _DOC_TYPE_BY_PERIOD:
        return []

    t1code, t2code, kind = _DOC_TYPE_BY_PERIOD[period.type]
    from_date = f"{period.year}0101"
    to_date = f"{period.year + 1}0630" if period.type is PeriodType.ANNUAL else f"{period.year + 1}0331"

    with default_client(source="hkexnews") as client:
        try:
            rows = _search_rows(
                client,
                ticker,
                from_date=from_date,
                to_date=to_date,
                t1code=t1code,
                t2code=t2code,
            )
        except Exception as exc:
            logger.warning(f"HKEXnews search failed for {ticker.code}: {exc}")
            return []

        chosen = _select_main(rows, ticker, period)
        if not chosen:
            logger.warning(f"[{ticker.code}] no HK {period.label()} filing found")
            return []

        url = chosen["url"]
        suffix = _suffix_from_url(url)
        dest = report_output_path(output_root, ticker, period, kind, suffix)
        n_bytes = stream_to_file(client, url, dest, source="hkexnews", rate=_HKEX_RATE)
        file_format = _format_from_suffix(suffix)

        return [ReportFile(
            ticker=ticker,
            period=period,
            kind=kind,
            local_path=str(dest),
            source_url=url,
            title=chosen.get("title"),
            file_size_bytes=n_bytes,
            form=chosen.get("doc_type"),
            filing_date=chosen.get("filing_date"),
            report_date=None,
            source_id=chosen.get("source_id"),
            accession_number=None,
            is_amendment=False,
            sequence=1,
            source_format=file_format,
            output_format=file_format,
        )]


def download_ipo_documents(ticker: Ticker, output_root: Path) -> list[ReportFile]:
    today = dt.date.today().strftime("%Y%m%d")

    with default_client(source="hkexnews") as client:
        try:
            formal_rows = _search_rows(
                client,
                ticker,
                from_date=_TITLE_SEARCH_START,
                to_date=today,
                t1code="30000",
            )
            selected = _select_ipo_rows(formal_rows, [], ticker)
            fallback_rows: list[dict] = []
            if not selected:
                fallback_rows = _search_rows(
                    client,
                    ticker,
                    from_date=_TITLE_SEARCH_START,
                    to_date=today,
                    t1code="91000",
                )
                selected = _select_ipo_rows(formal_rows, fallback_rows, ticker)
        except Exception as exc:
            logger.warning(f"HKEXnews IPO search failed for {ticker.code}: {exc}")
            return []

        if not selected:
            logger.warning(f"[{ticker.code}] no HKEX IPO documents found")
            return []

        out: list[ReportFile] = []
        for sequence, row in enumerate(selected, start=1):
            url = row["url"]
            suffix = _suffix_from_url(url)
            filing_date = row.get("filing_date")
            is_amendment = _is_ipo_amendment(row)
            label = row.get("title") or row.get("doc_type") or "ipo_document"
            dest = _filing_output_path(
                output_root,
                ticker,
                sequence=sequence,
                label=label,
                filing_date=filing_date,
                is_amendment=is_amendment,
                suffix=suffix,
                source_id=row.get("source_id"),
            )
            n_bytes = stream_to_file(client, url, dest, source="hkexnews", rate=_HKEX_RATE)
            file_format = _format_from_suffix(suffix)

            out.append(_make_report_file(
                ticker=ticker,
                period=None,
                kind="ipo_document",
                local_path=str(dest),
                source_url=url,
                title=label,
                file_size_bytes=n_bytes,
                form=_ipo_form(row),
                filing_date=filing_date,
                report_date=None,
                source_id=row.get("source_id"),
                accession_number=None,
                is_amendment=is_amendment,
                sequence=sequence,
                source_format=file_format,
                output_format=file_format,
            ))

        return out

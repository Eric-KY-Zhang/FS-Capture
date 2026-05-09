"""A-share annual / quarterly report download via cninfo (巨潮资讯网).

cninfo's hisAnnouncement endpoint returns disclosure announcements filtered by
stock code, category, and announcement date range. Each announcement carries an
`adjunctUrl` like `finalpage/2024-04-12/1219842831.PDF`; the absolute PDF URL is
`http://static.cninfo.com.cn/{adjunctUrl}`.

For A-share the auditor's report is bundled inside the annual report PDF (no
separate filing is published), so we surface a single 'annual_report' file for
annual periods.
"""

from __future__ import annotations

import datetime as dt
import html
import re
from pathlib import Path

from loguru import logger

from app.core.http import default_client, post_json, stream_to_file
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path

try:
    from app.core.output_paths import (
        report_output_path_for_filing as _report_output_path_for_filing,
    )
except ImportError:  # Main agent adds this public helper; keep this module importable meanwhile.
    _report_output_path_for_filing = None


_HISANNOUNCEMENT_URL = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
_PDF_BASE = "http://static.cninfo.com.cn/"
_CNINFO_RATE = 5.0
_PAGE_SIZE = 30


_CATEGORY = {
    PeriodType.ANNUAL: "category_ndbg_szsh",
    PeriodType.Q1: "category_yjdbg_szsh",
    PeriodType.Q2: "category_bndbg_szsh",  # 半年报
    PeriodType.Q3: "category_sjdbg_szsh",
}


_KIND = {
    PeriodType.ANNUAL: "annual_report",
    PeriodType.Q1: "q1_report",
    PeriodType.Q2: "interim_report",
    PeriodType.Q3: "q3_report",
}


_IPO_SEARCH_KEYS = (
    "招股说明书",
    "招股意向书",
    "首次公开发行",
    "申报稿",
    "注册稿",
    "修订稿",
    "更正",
    "补充",
    "更新",
)

_IPO_PRIMARY_KEYWORDS = (
    "招股说明书",
    "招股意向书",
    "首次公开发行",
    "申报稿",
    "注册稿",
)

_IPO_EXCLUDE_KEYWORDS = (
    "摘要",
    "上市公告书",
    "发行公告",
    "申购公告",
    "认购公告",
    "中签",
    "摇号",
    "配售结果",
    "发行结果",
    "路演",
    "保荐书",
    "保荐工作报告",
    "法律意见书",
    "律师工作报告",
    "审计报告",
    "验资报告",
    "评估报告",
    "问询回复",
    "问询函回复",
    "回复意见",
    "反馈意见回复",
    "落实函回复",
    "审核问询函",
    "承销",
    "战略配售",
    "初步询价",
    "询价公告",
    "定价公告",
    "缴款",
    "风险提示公告",
    "提示性公告",
)

_AMENDMENT_KEYWORDS = ("修订", "更正", "补充", "更新")

_TITLE_TAG_RE = re.compile(r"<[^>]+>")
_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _column_for(code: str) -> str:
    if code.startswith(("60", "688", "900")):
        return "sse"
    if code.startswith(("000", "001", "002", "200", "300", "301")):
        return "szse"
    return "bj"


def _clean_title(title: str) -> str:
    return html.unescape(_TITLE_TAG_RE.sub("", title or "")).strip()


def _announcement_source_id(announcement: dict) -> str:
    return str(
        announcement.get("announcementId")
        or announcement.get("announcement_id")
        or announcement.get("adjunctUrl")
        or ""
    )


def _announcement_key(announcement: dict) -> tuple[str, str, str]:
    source_id = _announcement_source_id(announcement)
    if source_id:
        return ("source", source_id, "")
    return (
        "fallback",
        str(announcement.get("adjunctUrl") or ""),
        str(announcement.get("announcementTime") or ""),
    )


def _announcement_time(announcement: dict) -> int:
    value = announcement.get("announcementTime") or announcement.get("announcementDate") or 0
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _announcement_date(announcement: dict) -> dt.date | None:
    raw = _announcement_time(announcement)
    if not raw:
        return None
    seconds = raw / 1000 if raw > 10_000_000_000 else raw
    try:
        return dt.datetime.fromtimestamp(seconds).date()
    except (OSError, OverflowError, ValueError):
        return None


def _pdf_url(adjunct_url: str) -> str:
    if adjunct_url.startswith(("http://", "https://")):
        return adjunct_url
    return _PDF_BASE + adjunct_url.lstrip("/")


def _cninfo_headers(ticker: Ticker) -> dict:
    return {
        "Origin": "http://www.cninfo.com.cn",
        "Referer": (f"http://www.cninfo.com.cn/new/disclosure/stock?stockCode={ticker.code}"),
        "Accept": "application/json, text/plain, */*",
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    }


def _base_query_data(ticker: Ticker) -> dict:
    return {
        "stock": f"{ticker.code},{ticker.external_id}" if ticker.external_id else ticker.code,
        "tabName": "fulltext",
        "pageSize": _PAGE_SIZE,
        "pageNum": 1,
        "column": _column_for(ticker.code),
        "category": "",
        "seDate": "",
        "searchkey": "",
        "secid": "",
        "plate": "",
        "isHLtitle": "true",
    }


def _payload_total(payload: dict) -> int | None:
    for key in ("totalAnnouncement", "totalRecordNum", "totalRecords", "total"):
        value = payload.get(key)
        if value is None:
            continue
        try:
            return int(value)
        except (TypeError, ValueError):
            continue
    return None


def _query_cninfo_announcements(
    client,
    ticker: Ticker,
    *,
    category: str = "",
    se_date: str = "",
    searchkey: str = "",
) -> list[dict]:
    """Query cninfo hisAnnouncement with pagination and source-id de-duplication."""
    headers = _cninfo_headers(ticker)
    seen: set[tuple[str, str, str]] = set()
    out: list[dict] = []
    page_num = 1

    while True:
        data = _base_query_data(ticker)
        data.update(
            {
                "pageNum": page_num,
                "category": category,
                "seDate": se_date,
                "searchkey": searchkey,
            }
        )
        payload = post_json(
            client,
            _HISANNOUNCEMENT_URL,
            source="cninfo",
            rate=_CNINFO_RATE,
            data=data,
            headers=headers,
        )
        page_rows = payload.get("announcements") or []
        if not page_rows:
            break

        new_count = 0
        for announcement in page_rows:
            key = _announcement_key(announcement)
            if key in seen:
                continue
            seen.add(key)
            out.append(announcement)
            new_count += 1

        total = _payload_total(payload)
        if total is not None and page_num * _PAGE_SIZE >= total:
            break
        if len(page_rows) < _PAGE_SIZE:
            break
        if new_count == 0 and page_num > 1:
            break
        page_num += 1

    return out


def _announcement_window(period: Period) -> tuple[str, str]:
    """Filing typically posted within ~6 months of period end. Use generous window."""
    if period.type is PeriodType.ANNUAL:
        return f"{period.year + 1}-01-01", f"{period.year + 1}-09-30"
    if period.type is PeriodType.Q1:
        return f"{period.year}-04-01", f"{period.year}-07-31"
    if period.type is PeriodType.Q2:
        return f"{period.year}-07-01", f"{period.year}-10-31"
    if period.type is PeriodType.Q3:
        return f"{period.year}-10-01", f"{period.year + 1}-02-28"
    raise ValueError(f"Unknown period type {period.type}")


def _is_amendment(title: str) -> bool:
    bad_substrings = ("摘要", "更正", "已取消", "意见反馈")
    return any(s in title for s in bad_substrings)


def _is_ipo_amendment(title: str) -> bool:
    return any(keyword in title for keyword in _AMENDMENT_KEYWORDS)


def _is_pdf_announcement(announcement: dict) -> bool:
    adjunct_url = str(announcement.get("adjunctUrl") or "")
    return adjunct_url.lower().endswith(".pdf")


def _is_ipo_document(announcement: dict) -> bool:
    if not _is_pdf_announcement(announcement):
        return False
    title = _clean_title(announcement.get("announcementTitle") or "")
    if not title:
        return False
    if any(keyword in title for keyword in _IPO_EXCLUDE_KEYWORDS):
        return False
    return any(keyword in title for keyword in _IPO_PRIMARY_KEYWORDS)


def _fallback_safe_filename(value: str, *, fallback: str = "file", max_len: int = 120) -> str:
    cleaned = _BAD_FILENAME_CHARS.sub("_", value).strip(" ._")
    if not cleaned:
        cleaned = fallback
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" ._")
    return cleaned or fallback


def _ipo_output_path(
    output_root: Path,
    ticker: Ticker,
    *,
    sequence: int,
    label: str,
    filing_date: dt.date | None,
    is_amendment: bool,
) -> Path:
    if _report_output_path_for_filing is not None:
        return _report_output_path_for_filing(
            output_root,
            ticker,
            "ipo",
            sequence,
            label,
            ".pdf",
            filing_date=filing_date,
            is_amendment=is_amendment,
        )

    date_part = filing_date.isoformat() if filing_date else "undated"
    amend_part = "_amendment" if is_amendment else ""
    safe_label = _fallback_safe_filename(label, fallback="ipo_document")
    filename = _fallback_safe_filename(
        f"{ticker.exchange.value}_{ticker.code}_ipo_{sequence:03d}_{date_part}{amend_part}_{safe_label}.pdf",
        fallback=f"{ticker.exchange.value}_{ticker.code}_ipo_{sequence:03d}.pdf",
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


def _select_main_filing(announcements: list[dict], year: int, period: PeriodType) -> dict | None:
    """Pick the canonical filing for the period from candidate announcements."""
    if not announcements:
        return None

    year_str = str(year)
    keyword = {
        PeriodType.ANNUAL: "年度报告",
        PeriodType.Q1: "第一季度报告",
        PeriodType.Q2: "半年度报告",
        PeriodType.Q3: "第三季度报告",
    }[period]

    candidates = [
        a
        for a in announcements
        if keyword in (a.get("announcementTitle") or "")
        and year_str in (a.get("announcementTitle") or "")
        and not _is_amendment(a.get("announcementTitle") or "")
    ]
    if not candidates:
        # Fall back to anything mentioning the year
        candidates = [
            a
            for a in announcements
            if year_str in (a.get("announcementTitle") or "")
            and not _is_amendment(a.get("announcementTitle") or "")
        ]
    if not candidates:
        return None

    # Latest by announcementTime
    candidates.sort(key=lambda a: int(a.get("announcementTime") or 0), reverse=True)
    return candidates[0]


def _query_announcements(client, ticker: Ticker, period: Period) -> list[dict]:
    se_start, se_end = _announcement_window(period)
    return _query_cninfo_announcements(
        client,
        ticker,
        category=_CATEGORY[period.type],
        se_date=f"{se_start}~{se_end}",
    )


def _query_ipo_announcements(client, ticker: Ticker) -> list[dict]:
    seen: set[tuple[str, str, str]] = set()
    out: list[dict] = []
    for searchkey in _IPO_SEARCH_KEYS:
        try:
            rows = _query_cninfo_announcements(client, ticker, searchkey=searchkey)
        except Exception as exc:
            logger.warning(f"cninfo IPO search failed for {ticker.code} {searchkey}: {exc}")
            continue
        for announcement in rows:
            key = _announcement_key(announcement)
            if key in seen:
                continue
            seen.add(key)
            if _is_ipo_document(announcement):
                out.append(announcement)
    out.sort(key=lambda a: (_announcement_time(a), _announcement_source_id(a)))
    return out


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if not ticker.external_id:
        raise ValueError(f"A股下载需要 cninfo orgId（external_id 为空）：{ticker.code}")
    if period.type is PeriodType.IPO_PROSPECTUS:
        reports = download_ipo_documents(ticker, output_root)
        for report in reports:
            report.period = period
            report.kind = "ipo_prospectus"
        return reports

    with default_client(source="cninfo") as client:
        announcements = _query_announcements(client, ticker, period)
        main = _select_main_filing(announcements, period.year, period.type)
        if not main:
            logger.warning(f"[{ticker.code}] no {period.label()} filing found in cninfo")
            return []

        adjunct_url = main.get("adjunctUrl")
        if not adjunct_url:
            return []
        pdf_url = _pdf_url(str(adjunct_url))

        kind = _KIND[period.type]
        dest = report_output_path(output_root, ticker, period, kind, ".pdf")
        n_bytes = stream_to_file(
            client,
            pdf_url,
            dest,
            source="cninfo",
            rate=_CNINFO_RATE,
        )

        return [
            ReportFile(
                ticker=ticker,
                period=period,
                kind=kind,
                local_path=str(dest),
                source_url=pdf_url,
                title=main.get("announcementTitle"),
                file_size_bytes=n_bytes,
                source_format="pdf",
                output_format="pdf",
            )
        ]


def download_ipo_documents(ticker: Ticker, output_root: Path) -> list[ReportFile]:
    if not ticker.external_id:
        raise ValueError(f"A股 IPO 文档下载需要 cninfo orgId（external_id 为空）：{ticker.code}")

    with default_client(source="cninfo") as client:
        announcements = _query_ipo_announcements(client, ticker)
        if not announcements:
            logger.warning(f"[{ticker.code}] no cninfo IPO documents found")
            return []

        out: list[ReportFile] = []
        for sequence, announcement in enumerate(announcements, start=1):
            adjunct_url = str(announcement.get("adjunctUrl") or "")
            if not adjunct_url:
                continue

            title = _clean_title(announcement.get("announcementTitle") or "")
            filing_date = _announcement_date(announcement)
            is_amendment = _is_ipo_amendment(title)
            pdf_url = _pdf_url(adjunct_url)
            dest = _ipo_output_path(
                output_root,
                ticker,
                sequence=sequence,
                label=title or "ipo_document",
                filing_date=filing_date,
                is_amendment=is_amendment,
            )
            n_bytes = stream_to_file(
                client,
                pdf_url,
                dest,
                source="cninfo",
                rate=_CNINFO_RATE,
            )

            source_id = _announcement_source_id(announcement) or adjunct_url
            out.append(
                _make_report_file(
                    ticker=ticker,
                    period=None,
                    kind="ipo_document",
                    local_path=str(dest),
                    source_url=pdf_url,
                    title=title,
                    file_size_bytes=n_bytes,
                    form="IPO",
                    filing_date=filing_date,
                    report_date=None,
                    source_id=source_id,
                    accession_number=None,
                    is_amendment=is_amendment,
                    sequence=sequence,
                    source_format="pdf",
                    output_format="pdf",
                )
            )

        return out

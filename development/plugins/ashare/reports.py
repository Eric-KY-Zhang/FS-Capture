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

from pathlib import Path
from typing import Optional

from loguru import logger

from app.core.http import default_client, post_json, stream_to_file
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path


_HISANNOUNCEMENT_URL = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
_PDF_BASE = "http://static.cninfo.com.cn/"


_CATEGORY = {
    PeriodType.ANNUAL: "category_ndbg_szsh",
    PeriodType.Q1: "category_yjdbg_szsh",
    PeriodType.Q2: "category_bndbg_szsh",   # 半年报
    PeriodType.Q3: "category_sjdbg_szsh",
}


_KIND = {
    PeriodType.ANNUAL: "annual_report",
    PeriodType.Q1: "q1_report",
    PeriodType.Q2: "interim_report",
    PeriodType.Q3: "q3_report",
}


def _column_for(code: str) -> str:
    if code.startswith(("60", "688", "900")):
        return "sse"
    if code.startswith(("000", "001", "002", "200", "300", "301")):
        return "szse"
    return "bj"


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


def _select_main_filing(announcements: list[dict], year: int, period: PeriodType) -> Optional[dict]:
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
        a for a in announcements
        if keyword in (a.get("announcementTitle") or "")
        and year_str in (a.get("announcementTitle") or "")
        and not _is_amendment(a.get("announcementTitle") or "")
    ]
    if not candidates:
        # Fall back to anything mentioning the year
        candidates = [
            a for a in announcements
            if year_str in (a.get("announcementTitle") or "")
            and not _is_amendment(a.get("announcementTitle") or "")
        ]
    if not candidates:
        return None

    # Latest by announcementTime
    candidates.sort(key=lambda a: int(a.get("announcementTime") or 0), reverse=True)
    return candidates[0]


def _query_announcements(
    client, ticker: Ticker, period: Period
) -> list[dict]:
    se_start, se_end = _announcement_window(period)
    headers = {
        "Origin": "http://www.cninfo.com.cn",
        "Referer": (
            f"http://www.cninfo.com.cn/new/disclosure/stock?stockCode={ticker.code}"
        ),
        "Accept": "application/json, text/plain, */*",
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    }
    data = {
        "stock": f"{ticker.code},{ticker.external_id}",
        "tabName": "fulltext",
        "pageSize": 30,
        "pageNum": 1,
        "column": _column_for(ticker.code),
        "category": _CATEGORY[period.type],
        "seDate": f"{se_start}~{se_end}",
        "searchkey": "",
        "secid": "",
        "plate": "",
        "isHLtitle": "true",
    }
    payload = post_json(
        client, _HISANNOUNCEMENT_URL,
        source="cninfo", rate=5.0,
        data=data, headers=headers,
    )
    return payload.get("announcements") or []


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if not ticker.external_id:
        raise ValueError(f"A股下载需要 cninfo orgId（external_id 为空）：{ticker.code}")

    with default_client(source="cninfo") as client:
        announcements = _query_announcements(client, ticker, period)
        main = _select_main_filing(announcements, period.year, period.type)
        if not main:
            logger.warning(
                f"[{ticker.code}] no {period.label()} filing found in cninfo"
            )
            return []

        adjunct_url = main.get("adjunctUrl")
        if not adjunct_url:
            return []
        pdf_url = _PDF_BASE + adjunct_url

        kind = _KIND[period.type]
        dest = report_output_path(output_root, ticker, period, kind, ".pdf")
        n_bytes = stream_to_file(
            client, pdf_url, dest,
            source="cninfo", rate=5.0,
        )

        return [ReportFile(
            ticker=ticker,
            period=period,
            kind=kind,
            local_path=str(dest),
            source_url=pdf_url,
            title=main.get("announcementTitle"),
            file_size_bytes=n_bytes,
        )]

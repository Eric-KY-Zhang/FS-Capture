"""HK annual / interim / audit report download via HKEXnews titlesearch.

HKEXnews has no documented JSON API. Their `titlesearchservlet.do` endpoint
accepts a fixed parameter set and returns HTML. We parse it with selectolax to
recover document URLs.

For HK, audit reports are NOT separately filed — auditor's report is bundled
into the annual report PDF. We download annual reports for ANNUAL periods and
interim reports for Q2; Q1/Q3 generally have no HK filings.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from loguru import logger
from selectolax.parser import HTMLParser

from app.core.http import default_client, stream_to_file
from app.core.ratelimit import limiter
from app.core.models import Period, PeriodType, ReportFile, Ticker
from app.core.output_paths import report_output_path


_TITLESEARCH_URL = "https://www1.hkexnews.hk/search/titlesearch.xhtml"
_PDF_HOST = "https://www1.hkexnews.hk"

_NEGATIVE_TITLE_KEYWORDS = (
    "更正", "補充", "修訂", "澄清", "公告", "通函",
    "ESG", "環境", "社會", "管治", "Environmental", "Sustainability",
)


# HKEXnews internal document type codes (t1code / t2code).
# 40000 = 财务报告 (Financial Reports) > sub-categories.
_DOC_TYPE_BY_PERIOD = {
    PeriodType.ANNUAL: ("40000", "40100", "annual_report"),     # 年报
    PeriodType.Q2:     ("40000", "40200", "interim_report"),    # 中期/半年报
    PeriodType.Q3:     ("40000", "40300", "q3_report"),         # 第三季度报告
    PeriodType.Q1:     ("40000", "40400", "q1_report"),         # 第一季度报告
}


def _params(ticker: Ticker, period: Period) -> dict:
    t1, t2, _kind = _DOC_TYPE_BY_PERIOD[period.type]
    return {
        "lang": "ZH",
        "category": "0",
        "market": "SEHK",
        "searchType": "1",
        "documentType": "-1",
        "t1code": t1,
        "t2Gcode": "-2",
        "t2code": t2,
        "stockId": ticker.code.lstrip("0"),
        "from": f"{period.year}0101",
        "to": f"{period.year + 1}0630",
        "MB-LR": "M",
        "title": "",
    }


def _parse_results(html: str) -> list[dict]:
    """Extract {url, title, date} rows from titlesearch HTML."""
    tree = HTMLParser(html)
    rows: list[dict] = []
    for tr in tree.css("table.table-scroll tr"):
        cells = tr.css("td")
        if len(cells) < 4:
            continue
        date_txt = cells[0].text(strip=True)
        title_cell = cells[3]
        link = title_cell.css_first("a")
        if not link:
            continue
        href = link.attributes.get("href") or ""
        if not href:
            continue
        url = href if href.startswith("http") else _PDF_HOST + href
        rows.append({
            "url": url,
            "title": link.text(strip=True),
            "date": date_txt,
        })
    return rows


def _select_main(rows: list[dict], period: Period) -> Optional[dict]:
    if not rows:
        return None
    year_str = str(period.year)
    keywords = {
        PeriodType.ANNUAL: ("年報", "年度報告", "Annual Report"),
        PeriodType.Q1: ("第一季", "Q1"),
        PeriodType.Q2: ("中期", "半年", "Interim"),
        PeriodType.Q3: ("第三季", "Q3"),
    }[period.type]

    def matches(title: str) -> bool:
        return year_str in title and any(k.lower() in title.lower() for k in keywords)

    def score(row: dict) -> tuple:
        title = row["title"]
        negative = any(k.lower() in title.lower() for k in _NEGATIVE_TITLE_KEYWORDS)
        exact = title.strip().lower() in {
            f"年報 {year_str}".lower(),
            f"{year_str} 年報".lower(),
            f"annual report {year_str}".lower(),
            f"{year_str} annual report".lower(),
        }
        chinese_pdf = row["url"].lower().endswith("_c.pdf")
        return (
            negative,
            not exact,
            not chinese_pdf,
            len(title),
            row["date"],
        )

    candidates = [r for r in rows if matches(r["title"])]
    if not candidates:
        candidates = [r for r in rows if year_str in r["title"]]
    if not candidates:
        return None
    candidates.sort(key=score)
    return candidates[0]


def download(ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
    if period.type not in _DOC_TYPE_BY_PERIOD:
        return []

    params = _params(ticker, period)
    headers = {
        "Referer": "https://www1.hkexnews.hk/search/titlesearch.xhtml?lang=zh",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9",
    }

    with default_client(source="hkexnews") as client:
        limiter("hkexnews", 3.0).acquire_blocking()
        r = client.get(_TITLESEARCH_URL, params=params, headers=headers)
        if r.status_code != 200:
            logger.warning(f"HKEXnews search failed for {ticker.code}: HTTP {r.status_code}")
            return []
        rows = _parse_results(r.text)
        chosen = _select_main(rows, period)
        if not chosen:
            logger.warning(f"[{ticker.code}] no HK {period.label()} filing found")
            return []

        _, _, kind = _DOC_TYPE_BY_PERIOD[period.type]
        ext = ".pdf" if chosen["url"].lower().endswith(".pdf") else Path(chosen["url"]).suffix or ".htm"
        dest = report_output_path(output_root, ticker, period, kind, ext)
        n_bytes = stream_to_file(client, chosen["url"], dest, source="hkexnews", rate=3.0)

        return [ReportFile(
            ticker=ticker,
            period=period,
            kind=kind,
            local_path=str(dest),
            source_url=chosen["url"],
            title=chosen["title"],
            file_size_bytes=n_bytes,
        )]

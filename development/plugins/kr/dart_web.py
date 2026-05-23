"""DART public disclosure-page client for KR no-key fallback mode."""

from __future__ import annotations

import datetime as dt
import re
from typing import Any

from bs4 import BeautifulSoup
from loguru import logger

from app.core.ratelimit import limiter
from app.core.settings import load_settings

_DETAIL_SEARCH_URL = "https://dart.fss.or.kr/dsab007/detailSearch.ax"
_REFERER = "https://dart.fss.or.kr/dsab007/main.do?option=corp"

_SELECTORS = {
    "rows": "tbody#tbody tr",
    "corp_link": "a[href*='openCorpInfoNew']",
    "report_link": "a[href*='/dsaf001/main.do'], a[onclick*='openReportViewer']",
}

_CORP_CODE_RE = re.compile(r"openCorpInfoNew\(\s*['\"](?P<corp_code>\d+)['\"]")
_RCEPT_NO_RE = re.compile(
    r"(?:rcpNo=|openReportViewer\(\s*['\"])(?P<rcept_no>\d+)", re.IGNORECASE
)
_SPACE_RE = re.compile(r"\s+")


def _dart_web_rate() -> float:
    return load_settings().rate_limits.dart_web


def _dart_client_context():
    from .reports import _dart_client

    return _dart_client()


def _normalize_date(text: str) -> str:
    digits = re.sub(r"\D", "", str(text or ""))
    return digits if len(digits) == 8 else ""


def _normalize_text(text: str) -> str:
    return _SPACE_RE.sub(" ", str(text or "")).strip()


def _base_payload(
    *,
    text: str = "",
    corp_code: str = "",
    bgn_de: str = "",
    end_de: str = "",
    detail_type: str = "",
    final: bool = True,
    current_page: int = 1,
    max_results: int = 100,
) -> dict[str, str]:
    payload = {
        "currentPage": str(current_page),
        "maxResults": str(max_results),
        "maxLinks": "10",
        "sort": "date",
        "series": "desc",
        "option": "corp",
        "textCrpCik": corp_code,
        "textCrpNm": text,
        "textCrpNm2": text,
        "reportName": "",
        "reportName2": "",
        "tocSrch": "",
        "tocSrch2": "",
        "textPresenterNm": "",
        "startDate": _normalize_date(bgn_de),
        "endDate": _normalize_date(end_de),
        "decadeType": "",
        "businessCode": "all",
        "businessNm": "전체",
        "corporationType": "all",
        "closingAccountsMonth": "all",
        "lateKeyword": "",
        "keyword": "",
        "textkeyword": "",
        "reportNamePopYn": "",
        "autoSearch": "N",
        "autoSearchCorp": "Y",
    }
    if final:
        payload["finalReport"] = "recent"
    if detail_type:
        payload["publicType"] = detail_type
    return payload


def _detail_search_html(**payload_kwargs: Any) -> str:
    payload = _base_payload(**payload_kwargs)
    headers = {
        "Referer": _REFERER,
        "X-Requested-With": "XMLHttpRequest",
    }
    with _dart_client_context() as client:
        limiter("dart_web", _dart_web_rate()).acquire_blocking()
        response = client.post(_DETAIL_SEARCH_URL, data=payload, headers=headers, timeout=60.0)
        response.raise_for_status()
        return response.text


def _corp_code_from_href(href: str) -> str:
    match = _CORP_CODE_RE.search(href or "")
    return match.group("corp_code") if match else ""


def _rcept_no_from_link(link) -> str:
    href = link.get("href", "") if link else ""
    onclick = link.get("onclick", "") if link else ""
    match = _RCEPT_NO_RE.search(f"{href} {onclick}")
    return match.group("rcept_no") if match else ""


def _parse_detail_rows(html: str) -> list[dict[str, str]]:
    soup = BeautifulSoup(html or "", "lxml")
    rows: list[dict[str, str]] = []
    for tr in soup.select(_SELECTORS["rows"]):
        cells = tr.find_all("td")
        if len(cells) < 5:
            continue
        corp_link = cells[1].select_one(_SELECTORS["corp_link"])
        report_link = cells[2].select_one(_SELECTORS["report_link"])
        corp_code = _corp_code_from_href(corp_link.get("href", "") if corp_link else "")
        rcept_no = _rcept_no_from_link(report_link)
        if not corp_code or not rcept_no:
            continue
        rows.append(
            {
                "corp_code": corp_code,
                "corp_name": _normalize_text(corp_link.get_text(" ", strip=True)),
                "rcept_no": rcept_no,
                "report_nm": _normalize_text(report_link.get_text(" ", strip=True)),
                "rcept_dt": _normalize_date(cells[4].get_text(" ", strip=True)),
            }
        )
    return rows


def resolve_corp(stock_code: str) -> dict[str, str] | None:
    """Resolve a KR stock code to DART corp_code via public disclosure search."""
    norm = stock_code.strip().zfill(6)
    try:
        html = _detail_search_html(text=norm, max_results=15, final=True)
    except Exception as exc:
        logger.warning(f"DART public corp lookup failed for {norm}: {exc}")
        return None
    for row in _parse_detail_rows(html):
        return {"corp_code": row["corp_code"], "corp_name": row["corp_name"]}
    return None

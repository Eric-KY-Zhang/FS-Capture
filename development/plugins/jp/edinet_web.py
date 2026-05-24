"""EDINET public-web fallback for no-key JP filings.

Search uses Playwright because EDINET's GeneXus AJAX payload cannot be replayed
reliably with httpx. PDF download stays on the shared httpx path because the
public search result exposes a direct document id.
"""

from __future__ import annotations

import json
import re
import time
from datetime import date
from pathlib import Path
from typing import Any

from loguru import logger

from app.core.http import default_client, stream_to_file
from app.core.ratelimit import limiter
from app.core.settings import load_settings

_SEARCH_URL = "https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx"
_PDF_URL_TEMPLATE = "https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/{doc_id}.pdf"
_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/126.0.0.0 Safari/537.36"
)
_CONTEXT_MARKER = f"edinet_web:{_USER_AGENT}"
_SEARCH_TIMEOUT_SECONDS = 60.0

_SELECTORS = {
    "keyword": 'input[name="W0018vD_KEYWORD"]',
    "kikan": 'select[name="W0018vD_KIKAN"]',
    "syorui1": 'input[name="W0018vCHKSYORUI1"]',
    "search_button": 'input[name="W0018BTNBTN_SEARCH"]',
}

_PERIOD_RE = re.compile(
    r"[（(]\d{4}/\d{2}/\d{2}\s*[－\-ー〜~～]\s*(\d{4})/(\d{2})/(\d{2})[)）]"
)


def _edinet_web_rate() -> float:
    return load_settings().rate_limits.edinet_web


def _doc_id(row: dict[str, Any]) -> str:
    for key in ("SHORUI_KANRI_NO", "CODE_2", "CODE_3", "CODE_4"):
        value = str(row.get(key) or "").strip()
        if value:
            return value
    return ""


def _first_text(row: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = str(row.get(key) or "").strip()
        if value:
            return value
    return ""


def _normalize_doc_type(value: Any) -> str:
    return str(value or "").strip()[:3]


def _normalize_submit_datetime(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return text.replace("/", "-", 2)


def _normalize_sec_code(value: Any, fallback: str = "") -> str:
    sec_code = str(value or "").strip()
    if len(sec_code) == 5 and sec_code.endswith("0"):
        sec_code = sec_code[:4]
    if sec_code:
        return sec_code
    return str(fallback or "").strip()


def _period_end_from_title(title: str) -> str:
    match = _PERIOD_RE.search(title)
    if not match:
        return ""
    year, month, day = match.groups()
    return f"{year}-{month}-{day}"


def _pdf_flag(value: Any) -> str:
    return "1" if str(value or "").strip().lower() in {"1", "y", "yes", "true"} else ""


def _normalize_row(row: dict[str, Any], *, ticker_code: str = "") -> dict[str, Any]:
    title = _first_text(row, "SHORUI_NAME", "docDescription", "title")
    return {
        "doc_id": _doc_id(row),
        "doc_type_code": _normalize_doc_type(row.get("SYORUI_SB_CD_ID")),
        "submit_date_time": _normalize_submit_datetime(row.get("TEISHUTSU_NICHIJI")),
        "period_end": _period_end_from_title(title),
        "edinet_code": _first_text(row, "EDINET_CD", "edinetCode", "edinet_code"),
        "sec_code": _normalize_sec_code(row.get("SHOKEN_CD"), ticker_code),
        "jcn": _first_text(row, "JCN", "jcn"),
        "filer_name": _first_text(
            row,
            "TEISYUTUSYA_NAME",
            "TEISHUTSUSYA_NAME",
            "TEISHUTUSYA_NAME",
            "filerName",
            "filer_name",
        ),
        "title": title,
        "pdf_flag": _pdf_flag(row.get("PDFKBN") or row.get("pdfFlag") or row.get("pdf_flag")),
    }


def _extract_result_rows(response_json: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for value in response_json.get("gxValues") or []:
        if not isinstance(value, dict):
            continue
        raw = value.get("AV125W_RESULT_LIST_JSON")
        if not raw:
            continue
        try:
            parsed = json.loads(raw) if isinstance(raw, str) else raw
        except Exception as exc:
            logger.debug(f"EDINET result JSON parse failed: {exc}")
            continue
        if isinstance(parsed, list):
            rows.extend(row for row in parsed if isinstance(row, dict))
    return rows


def _dedupe_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[str] = set()
    result: list[dict[str, Any]] = []
    for row in rows:
        key = _doc_id(row) or json.dumps(row, ensure_ascii=False, sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def _is_search_response(response: Any) -> bool:
    request = getattr(response, "request", None)
    method = str(getattr(request, "method", "") or "").upper()
    if method and method != "POST":
        return False
    url = str(getattr(response, "url", "") or "")
    if "WEEE0030.aspx?" not in url and "GXAjaxRequest" not in url:
        return False
    return int(getattr(response, "status", 0) or 0) == 200


def _ensure_report_checkbox(page: Any) -> None:
    try:
        checked = page.is_checked(_SELECTORS["syorui1"])
    except Exception as exc:
        logger.debug(f"EDINET checkbox state probe failed, checking explicitly: {exc}")
        checked = False
    if not checked:
        page.check(_SELECTORS["syorui1"])


def _wait_for_search(page: Any, captured_rows: list[dict[str, Any]]) -> None:
    deadline = time.monotonic() + _SEARCH_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        if captured_rows:
            return
        page.wait_for_timeout(500)

    try:
        page.wait_for_function(
            "() => document.body.innerText.includes('レコードがありません。')",
            timeout=1_000,
        )
    except Exception as exc:
        logger.debug(f"EDINET no-record marker not observed after timeout: {exc}")


def _ensure_public_context() -> tuple[Any, Any]:
    from app.core import pdf_renderer

    state = pdf_renderer._ensure_state()
    if state.context is None or state.context_user_agent != _CONTEXT_MARKER:
        pdf_renderer._close_context(state)
        state.context = state.browser.new_context(
            user_agent=_USER_AGENT,
            locale="ja-JP",
            extra_http_headers={"Accept-Language": "ja,en;q=0.9"},
        )
        state.context_user_agent = _CONTEXT_MARKER
    return state.context, pdf_renderer


def _search_filings_once(ticker_code: str) -> list[dict[str, Any]]:
    limiter("edinet_web", _edinet_web_rate()).acquire_blocking()
    context, pdf_renderer = _ensure_public_context()
    page = None
    captured_rows: list[dict[str, Any]] = []

    try:
        page = context.new_page()

        def _on_response(response: Any) -> None:
            if not _is_search_response(response):
                return
            try:
                body = response.json()
            except Exception as exc:
                logger.debug(f"EDINET search response JSON decode failed: {exc}")
                return
            captured_rows.extend(_extract_result_rows(body))

        page.on("response", _on_response)
        page.goto(_SEARCH_URL, wait_until="networkidle", timeout=60_000)
        page.fill(_SELECTORS["keyword"], ticker_code)
        page.select_option(_SELECTORS["kikan"], "7")
        _ensure_report_checkbox(page)
        page.click(_SELECTORS["search_button"])
        _wait_for_search(page, captured_rows)
        return _dedupe_rows(captured_rows)
    finally:
        if page is not None:
            try:
                page.close()
            except Exception as exc:
                logger.warning(f"EDINET Playwright page close failed: {exc}")
        if not pdf_renderer._clear_context_state(context):
            logger.debug("EDINET Playwright context state was not fully cleared")


def search_filings_all(ticker_code: str) -> list[dict[str, Any]]:
    """Search EDINET public page once and return all normalized rows for a ticker."""
    normalized_ticker = str(ticker_code).strip()

    for attempt in range(2):
        try:
            raw_rows = _search_filings_once(normalized_ticker)
            return [
                _normalize_row(row, ticker_code=normalized_ticker)
                for row in raw_rows
                if isinstance(row, dict)
            ]
        except Exception as exc:
            logger.warning(f"EDINET public search failed for {normalized_ticker}: {exc}")
            if attempt == 0:
                from app.core import pdf_renderer

                pdf_renderer._shutdown_current_thread_renderer()
                continue
            raise
    return []


def search_filings(ticker_code: str, year: int) -> list[dict[str, Any]]:
    """Search EDINET public page once and return normalized rows for a filing year."""
    target_year = str(year)
    return [
        row
        for row in search_filings_all(ticker_code)
        if row["doc_id"] and row["submit_date_time"][:4] == target_year
    ]


def list_documents(
    submit_date: str | date | None = None,
    *,
    ticker: str | None = None,
    year: int | None = None,
) -> list[dict[str, Any]]:
    """Compatibility wrapper.

    Public EDINET no-key mode should use ``search_filings(ticker, year)`` so one
    Playwright search covers the full year. The per-date shape is retained for
    older call sites and returns no rows to avoid 365 browser searches.
    """
    if ticker and year:
        return search_filings(ticker, year)
    logger.warning(
        "EDINET public per-day list_documents is disabled; use search_filings(ticker, year)"
    )
    _ = submit_date
    return []


def download_document_pdf(doc_id: str, dest: Path) -> int:
    url = _PDF_URL_TEMPLATE.format(doc_id=doc_id)
    with default_client(source="edinet_web", timeout=120.0) as client:
        n_bytes = stream_to_file(
            client,
            url,
            dest,
            source="edinet_web",
            rate=_edinet_web_rate(),
            read_timeout=180.0,
        )
    with dest.open("rb") as handle:
        header = handle.read(4)
    if header != b"%PDF":
        dest.unlink(missing_ok=True)
        raise ValueError(f"EDINET public PDF is not valid: {url}")
    return n_bytes

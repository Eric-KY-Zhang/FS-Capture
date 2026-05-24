"""Spike EDINET public web search and PDF download behavior.

This is intentionally a one-off script for the v1.0 JP addendum. It verifies
the live WEEE0030 public search page before the real plugin implementation is
chosen.
"""

from __future__ import annotations

import argparse
import base64
import json
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import httpx
from bs4 import BeautifulSoup
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

SEARCH_URL = "https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx"
PDF_BASE_URL = "https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/126.0.0.0 Safari/537.36"
)

TICKERS = {
    "7203": "Toyota",
    "6758": "Sony",
    "9984": "SoftBank Group",
}


@dataclass
class FilingPick:
    ticker: str
    company: str
    submit_datetime: str
    title: str
    doc_id: str
    pdf_url: str
    pdf_path: Path
    pdf_size: int


def _headers(accept: str) -> dict[str, str]:
    return {
        "User-Agent": USER_AGENT,
        "Accept": accept,
        "Accept-Language": "ja,en;q=0.9",
    }


def inspect_initial_page() -> dict[str, Any]:
    """GET WEEE0030 and record the real form/state fields."""
    with httpx.Client(
        headers=_headers("text/html,application/xhtml+xml,*/*"),
        follow_redirects=True,
        timeout=60.0,
    ) as client:
        response = client.get(SEARCH_URL)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "lxml")
        fields = []
        for tag in soup.select("input, select, textarea"):
            name = tag.get("name")
            if not name:
                continue
            fields.append(
                {
                    "tag": tag.name,
                    "name": name,
                    "type": tag.get("type") or "",
                    "value_len": len(tag.get("value") or tag.text or ""),
                }
            )
        select = soup.select_one("select[name='W0018vD_KIKAN']")
        options = []
        if select:
            options = [
                {"value": option.get("value"), "label": option.get_text(" ", strip=True)}
                for option in select.select("option")
            ]
        return {
            "status": response.status_code,
            "cookies": sorted(cookie.name for cookie in client.cookies.jar),
            "has_viewstate": bool(soup.select_one("input[name='__VIEWSTATE']")),
            "has_eventvalidation": bool(soup.select_one("input[name='__EVENTVALIDATION']")),
            "fields": fields,
            "period_options": options,
        }


def httpx_direct_search_probe(ticker: str) -> dict[str, Any]:
    """Try the visible URL-query route; current EDINET returns no results."""
    query = (
        f"mul={ticker}&ctf=off&fls=on&lpr=off&rpr=off&oth=off"
        "&yer=&mon=&pfs=7&ser=1&pag=1&sor=2"
    )
    encoded = base64.b64encode(query.encode("utf-8")).decode("ascii")
    url = f"{SEARCH_URL}?{encoded}"
    with httpx.Client(
        headers=_headers("text/html,application/xhtml+xml,*/*"),
        follow_redirects=True,
        timeout=60.0,
    ) as client:
        response = client.get(url)
        response.raise_for_status()
    soup = BeautifulSoup(response.text, "lxml")
    gx_state = soup.select_one("input[name='GXState']")
    rows = []
    total = ""
    if gx_state:
        state = json.loads(gx_state["value"])
        total = str(state.get("vW_TOTALCOUNT") or "")
        rows = json.loads(state.get("vW_RESULT_LIST_JSON") or "[]")
    return {
        "url": url,
        "status": response.status_code,
        "total": total,
        "rows": len(rows),
        "contains_company_text": any(name in response.text for name in ("トヨタ", "ソニー", "ソフトバンク")),
    }


def _extract_result_rows(response_json: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for value in response_json.get("gxValues", []):
        raw = value.get("AV125W_RESULT_LIST_JSON")
        if raw:
            rows.extend(json.loads(raw))
    return rows


def search_with_playwright(ticker: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Use the real browser to execute GeneXus AJAX and capture JSON rows."""
    captured_rows: list[dict[str, Any]] = []
    ajax_events: list[dict[str, Any]] = []
    footer_text = ""

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=USER_AGENT, locale="ja-JP")

        def on_response(response) -> None:
            request = response.request
            if request.method != "POST" or "WEEE0030.aspx?" not in response.url:
                return
            ajax_events.append(
                {
                    "url": response.url,
                    "status": response.status,
                    "content_type": response.headers.get("content-type", ""),
                    "post_body_head": (request.post_data or "")[:500],
                }
            )
            try:
                data = response.json()
            except Exception:
                return
            captured_rows.extend(_extract_result_rows(data))

        page.on("response", on_response)
        page.goto(SEARCH_URL, wait_until="networkidle", timeout=60_000)
        page.fill("#W0018vD_KEYWORD", ticker)
        page.select_option("#W0018vD_KIKAN", "7")
        if not page.locator("#W0018vCHKSYORUI1").is_checked():
            page.check("#W0018vCHKSYORUI1")
        page.click("#W0018BTNBTN_SEARCH")
        deadline = time.monotonic() + 60
        while time.monotonic() < deadline:
            if captured_rows:
                break
            page.wait_for_timeout(500)
        if not captured_rows:
            try:
                page.wait_for_function(
                    """
                    () => document.body.innerText.includes("レコードがありません。")
                    """,
                    timeout=1_000,
                )
            except PlaywrightTimeoutError:
                pass
        if page.locator("#TBL_WEEE0030_FOOTER").count():
            footer_text = page.locator("#TBL_WEEE0030_FOOTER").inner_text()
        browser.close()

    # AJAX responses can be duplicated during event dispatch; de-dupe by doc id.
    seen: set[str] = set()
    rows: list[dict[str, Any]] = []
    for row in captured_rows:
        doc_id = _doc_id(row)
        key = doc_id or json.dumps(row, ensure_ascii=False, sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        rows.append(row)

    return rows, {
        "ajax_post_count": len(ajax_events),
        "ajax_events": ajax_events,
        "footer": footer_text,
    }


def _doc_id(row: dict[str, Any]) -> str:
    for key in ("SHORUI_KANRI_NO", "CODE_2", "CODE_3", "CODE_4"):
        value = str(row.get(key) or "").strip()
        if value:
            return value
    return ""


def select_annual_2024(rows: list[dict[str, Any]]) -> dict[str, Any]:
    candidates = [
        row
        for row in rows
        if str(row.get("TEISHUTSU_NICHIJI") or "").startswith("2024")
        and str(row.get("SYORUI_SB_CD_ID") or "").startswith("120")
        and "有価証券報告書" in str(row.get("SHORUI_NAME") or "")
    ]
    if not candidates:
        raise RuntimeError("No 2024 annual report row found")
    candidates.sort(key=lambda row: str(row.get("TEISHUTSU_NICHIJI") or ""))
    return candidates[-1]


def download_pdf(doc_id: str, dest: Path) -> int:
    url = f"{PDF_BASE_URL}/{doc_id}.pdf"
    dest.parent.mkdir(parents=True, exist_ok=True)
    with httpx.Client(
        headers=_headers("application/pdf,*/*"),
        follow_redirects=True,
        timeout=httpx.Timeout(connect=30.0, read=120.0, write=30.0, pool=30.0),
    ) as client:
        with client.stream("GET", url) as response:
            response.raise_for_status()
            with dest.open("wb") as handle:
                total = 0
                first = b""
                for chunk in response.iter_bytes(256 * 1024):
                    if not first:
                        first = chunk[:4]
                    handle.write(chunk)
                    total += len(chunk)
    if first != b"%PDF":
        raise RuntimeError(f"{url} did not return a PDF header")
    return total


def run(download_dir: Path) -> dict[str, Any]:
    page_probe = inspect_initial_page()
    direct_probe = httpx_direct_search_probe("7203")
    picks: list[FilingPick] = []
    search_meta: dict[str, Any] = {}

    for ticker, company in TICKERS.items():
        rows, meta = search_with_playwright(ticker)
        search_meta[ticker] = {
            "rows": len(rows),
            "footer": meta["footer"],
            "ajax_post_count": meta["ajax_post_count"],
        }
        selected = select_annual_2024(rows)
        doc_id = _doc_id(selected)
        pdf_url = f"{PDF_BASE_URL}/{doc_id}.pdf"
        pdf_path = download_dir / f"JP_{ticker}_{company}_2024_annual_spike.pdf"
        pdf_size = download_pdf(doc_id, pdf_path)
        picks.append(
            FilingPick(
                ticker=ticker,
                company=company,
                submit_datetime=str(selected.get("TEISHUTSU_NICHIJI") or ""),
                title=str(selected.get("SHORUI_NAME") or ""),
                doc_id=doc_id,
                pdf_url=pdf_url,
                pdf_path=pdf_path,
                pdf_size=pdf_size,
            )
        )
        time.sleep(1.0)

    return {
        "page_probe": page_probe,
        "direct_httpx_probe": direct_probe,
        "search_meta": search_meta,
        "picks": [pick.__dict__ | {"pdf_path": str(pick.pdf_path)} for pick in picks],
        "recommendation": "Plan B for search listing; httpx is safe for direct PDF download once doc_id is known.",
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--download-dir",
        type=Path,
        default=Path("output/jp_web_spike"),
        help="Ignored runtime output directory for downloaded PDFs.",
    )
    args = parser.parse_args()
    result = run(args.download_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

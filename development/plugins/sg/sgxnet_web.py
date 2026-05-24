"""SGXNet public API adapter."""

from __future__ import annotations

import html
import re
import threading
from pathlib import Path
from typing import Any

import httpx

from app.core.http import default_client, get_json, stream_to_file
from app.core.ratelimit import limiter
from app.core.settings import load_settings

CONFIG_URL = "https://www.sgx.com/config/appconfig.json?v=04c0b410"
_JSON_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://www.sgx.com/securities/company-announcements",
}
_HTML_HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://www.sgx.com/",
}

_LOCK = threading.Lock()
_CONFIG: dict[str, Any] | None = None
_TOKEN: str | None = None


def sgxnet_rate() -> float:
    return load_settings().rate_limits.sgxnet


def reset_sgxnet_cache() -> None:
    global _CONFIG, _TOKEN
    with _LOCK:
        _CONFIG = None
        _TOKEN = None


def _rot13(value: str) -> str:
    def convert(match: re.Match[str]) -> str:
        char = match.group(0)
        base = 65 if char <= "Z" else 97
        return chr((ord(char) + 13 - base) % 26 + base)

    return re.sub(r"[A-Za-z]", convert, value)


def _config(client: httpx.Client) -> dict[str, Any]:
    global _CONFIG
    if _CONFIG is not None:
        return _CONFIG
    with _LOCK:
        if _CONFIG is None:
            _CONFIG = get_json(
                client,
                CONFIG_URL,
                source="sgxnet",
                rate=sgxnet_rate(),
                headers=_JSON_HEADERS,
            )
        return _CONFIG


def _auth_headers(client: httpx.Client, *, refresh: bool = False) -> dict[str, str]:
    global _TOKEN
    cfg = _config(client)
    if _TOKEN is None or refresh:
        with _LOCK:
            if _TOKEN is None or refresh:
                payload = get_json(
                    client,
                    cfg["endpoints"]["CMS_API_URL"] + "/",
                    source="sgxnet",
                    rate=sgxnet_rate(),
                    params={"queryId": f"{cfg['CMS_VERSION']}:we_chat_qr_validator"},
                    headers=_JSON_HEADERS,
                )
                _TOKEN = _rot13(str(payload["data"]["qrValidator"]))
    return {**_JSON_HEADERS, "authorizationToken": str(_TOKEN)}


def _get_json_auth(
    client: httpx.Client,
    url: str,
    *,
    params: dict[str, Any] | None = None,
) -> Any:
    try:
        return get_json(
            client,
            url,
            source="sgxnet",
            rate=sgxnet_rate(),
            params=params,
            headers=_auth_headers(client),
        )
    except httpx.HTTPStatusError as exc:
        if exc.response.status_code not in {401, 403}:
            raise
        return get_json(
            client,
            url,
            source="sgxnet",
            rate=sgxnet_rate(),
            params=params,
            headers=_auth_headers(client, refresh=True),
        )


def search_announcements(
    symbol: str,
    *,
    start_year: int,
    end_year: int,
    page_size: int = 100,
    max_pages: int = 5,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with default_client(source="sgxnet", timeout=60.0) as client:
        url = _config(client)["endpoints"]["ANNOUNCEMENTS_API_URL"] + "securitycode"
        for page in range(max_pages):
            payload = _get_json_auth(
                client,
                url,
                params={
                    "value": symbol,
                    "pagestart": page,
                    "pagesize": page_size,
                    "periodstart": f"{start_year}0101_000000",
                    "periodend": f"{end_year}1231_235959",
                },
            )
            data = payload.get("data") or []
            rows.extend(data)
            if len(data) < page_size:
                break
    return rows


def list_financial_reports(*, page_size: int = 2000) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with default_client(source="sgxnet", timeout=60.0) as client:
        url = _config(client)["endpoints"]["FINANCIAL_REPORTS_API_URL"]
        params = {
            "pagestart": 0,
            "pagesize": page_size,
            "params": "id,companyName,documentDate,securityName,title,url",
        }
        first = _get_json_auth(client, url, params=params)
        rows.extend(first.get("data") or [])
        total_pages = int((first.get("meta") or {}).get("totalPages") or 1)
        for page in range(1, total_pages):
            params["pagestart"] = page
            rows.extend(_get_json_auth(client, url, params=params).get("data") or [])
    return rows


def list_ipo_prospectuses(*, page_size: int = 250) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with default_client(source="sgxnet", timeout=60.0) as client:
        base_url = _config(client)["endpoints"]["IPO_PROSPECTUS_API_URL"]
        count_payload = _get_json_auth(client, base_url + "count")
        count = int(count_payload.get("count") or 0)
        params = {
            "pagestart": 0,
            "pagesize": page_size,
            "params": "closing_date,name,id,modified_date,url,status",
        }
        for page in range((count + page_size - 1) // page_size):
            params["pagestart"] = page
            rows.extend(_get_json_auth(client, base_url, params=params).get("data") or [])
    return rows


def extract_pdf_links(page_url: str) -> list[str]:
    with default_client(source="sgxnet", timeout=60.0) as client:
        limiter("sgxnet", sgxnet_rate()).acquire_blocking()
        resp = client.get(page_url, headers=_HTML_HEADERS)
        resp.raise_for_status()
    links: list[str] = []
    for href in re.findall(r'href="([^"]+)"', resp.text, flags=re.I):
        decoded = html.unescape(href)
        if ".pdf" not in decoded.lower() and "FileOpen" not in decoded:
            continue
        links.append(str(httpx.URL(page_url).join(decoded)))
    return links


def download_pdf(pdf_url: str, dest: Path) -> int:
    with default_client(source="sgxnet", timeout=120.0) as client:
        n_bytes = stream_to_file(
            client,
            pdf_url,
            dest,
            source="sgxnet",
            rate=sgxnet_rate(),
            read_timeout=180.0,
        )
    if dest.read_bytes()[:4] != b"%PDF":
        dest.unlink(missing_ok=True)
        raise ValueError(f"SGXNet attachment is not a PDF: {pdf_url}")
    return n_bytes

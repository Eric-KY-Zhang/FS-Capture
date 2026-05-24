"""One-off SGXNet API spike for Sprint v1.0 batch 1.

Run from the repository root:
    python tmp_pytest_ok/sg_spike.py
"""
from __future__ import annotations

import datetime as dt
import html
import math
import re
import time
from dataclasses import dataclass
from typing import Any

import httpx


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
)
BASE_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://www.sgx.com/securities/company-announcements",
}
CONFIG_URL = "https://www.sgx.com/config/appconfig.json?v=04c0b410"


@dataclass(frozen=True)
class PdfCheck:
    label: str
    page_url: str
    pdf_url: str
    status_code: int
    content_type: str
    size_bytes: int
    first_bytes: bytes


def _rot13(value: str) -> str:
    def convert(match: re.Match[str]) -> str:
        char = match.group(0)
        base = 65 if char <= "Z" else 97
        return chr((ord(char) + 13 - base) % 26 + base)

    return re.sub(r"[A-Za-z]", convert, value)


def _date_from_ms(value: int | None) -> dt.date | None:
    if value is None:
        return None
    return dt.datetime.fromtimestamp(value / 1000, dt.UTC).date()


def _date_from_yyyymmdd(value: str | None) -> dt.date | None:
    if not value:
        return None
    return dt.datetime.strptime(value, "%Y%m%d").date()


def _normalize_ticker(value: str) -> str:
    return value.strip().upper().removesuffix(".SI")


def _get_config(client: httpx.Client) -> dict[str, Any]:
    resp = client.get(CONFIG_URL)
    resp.raise_for_status()
    return resp.json()


def _get_authorization_token(client: httpx.Client, config: dict[str, Any]) -> str:
    cms_url = config["endpoints"]["CMS_API_URL"]
    cms_version = config["CMS_VERSION"]
    resp = client.get(cms_url + "/", params={"queryId": f"{cms_version}:we_chat_qr_validator"})
    resp.raise_for_status()
    qr_validator = resp.json()["data"]["qrValidator"]
    return _rot13(qr_validator)


def _fetch_financial_reports(
    client: httpx.Client,
    config: dict[str, Any],
    headers: dict[str, str],
) -> list[dict[str, Any]]:
    base_url = config["endpoints"]["FINANCIAL_REPORTS_API_URL"]
    params = {
        "pagestart": 0,
        "pagesize": 2000,
        "params": "id,companyName,documentDate,securityName,title,url",
    }
    first = client.get(base_url, headers=headers, params=params).json()
    pages = first["meta"]["totalPages"]
    rows = list(first["data"])
    for page in range(1, pages):
        params["pagestart"] = page
        rows.extend(client.get(base_url, headers=headers, params=params).json()["data"])
    return rows


def _find_annual(
    rows: list[dict[str, Any]],
    *,
    security_name: str,
    report_year: int,
) -> dict[str, Any]:
    target = security_name.upper()
    candidates = []
    for row in rows:
        document_date = _date_from_ms(row.get("documentDate"))
        if not document_date or document_date.year != report_year:
            continue
        if row.get("securityName", "").upper() != target:
            continue
        if row.get("title", "").upper() != "ANNUAL REPORT":
            continue
        candidates.append(row)
    if not candidates:
        raise RuntimeError(f"No annual report found for {security_name} {report_year}")
    return candidates[0]


def _fetch_announcements(
    client: httpx.Client,
    config: dict[str, Any],
    headers: dict[str, str],
    ticker: str,
    start_year: int,
    end_year: int,
) -> list[dict[str, Any]]:
    base_url = config["endpoints"]["ANNOUNCEMENTS_API_URL"] + "securitycode"
    rows: list[dict[str, Any]] = []
    params = {
        "value": ticker,
        "pagestart": 0,
        "pagesize": 100,
        "periodstart": f"{start_year}0101_000000",
        "periodend": f"{end_year}1231_235959",
    }
    for page in range(5):
        params["pagestart"] = page
        payload = client.get(base_url, headers=headers, params=params).json()
        data = payload.get("data") or []
        rows.extend(data)
        if len(data) < params["pagesize"]:
            break
    return rows


def _find_h1(rows: list[dict[str, Any]], report_year: int) -> dict[str, Any]:
    candidates = []
    for row in rows:
        submission_date = _date_from_yyyymmdd(row.get("submission_date"))
        if not submission_date or submission_date.year != report_year:
            continue
        text = " ".join(
            str(row.get(key) or "")
            for key in ("category_name", "title", "issuer_name", "security_name")
        ).upper()
        if "FINANCIAL STATEMENTS" in text and "HALF YEARLY RESULTS" in text:
            candidates.append(row)
    if not candidates:
        raise RuntimeError(f"No H1 filing found for {report_year}")
    return candidates[0]


def _fetch_ipos(
    client: httpx.Client,
    config: dict[str, Any],
    headers: dict[str, str],
) -> list[dict[str, Any]]:
    base_url = config["endpoints"]["IPO_PROSPECTUS_API_URL"]
    count = client.get(base_url + "count", headers=headers).json()["count"]
    rows: list[dict[str, Any]] = []
    for page in range(math.ceil(count / 250)):
        payload = client.get(
            base_url,
            headers=headers,
            params={
                "pagestart": page,
                "pagesize": 250,
                "params": "closing_date,name,id,modified_date,url,status",
            },
        ).json()
        rows.extend(payload.get("data") or [])
    return rows


def _find_ipo(rows: list[dict[str, Any]], closing_year: int) -> dict[str, Any]:
    for row in rows:
        closing_date = _date_from_ms(row.get("closing_date"))
        if closing_date and closing_date.year == closing_year:
            return row
    raise RuntimeError(f"No IPO prospectus found for closing year {closing_year}")


def _extract_pdf_links(client: httpx.Client, page_url: str) -> list[str]:
    resp = client.get(page_url, headers={"Referer": "https://www.sgx.com/"})
    resp.raise_for_status()
    links: list[str] = []
    for href in re.findall(r'href="([^"]+)"', resp.text, flags=re.I):
        decoded = html.unescape(href)
        if ".pdf" not in decoded.lower() and "FileOpen" not in decoded:
            continue
        links.append(str(httpx.URL(page_url).join(decoded)))
    return links


def _choose_pdf(label: str, links: list[str]) -> str:
    lowered = [(link.lower(), link) for link in links]
    if "annual" in label:
        for text, link in lowered:
            if "annual%20report" in text or "annual report" in text:
                return link
    if "h1" in label:
        for text, link in lowered:
            if "interim%20financial%20statement" in text or "interim financial statement" in text:
                return link
    if "ipo" in label:
        for text, link in lowered:
            if "summary" not in text and "highlights" not in text:
                return link
    if not links:
        raise RuntimeError(f"No PDF links found for {label}")
    return links[0]


def _measure_pdf(client: httpx.Client, label: str, page_url: str, pdf_url: str) -> PdfCheck:
    size = 0
    first_bytes = b""
    with client.stream(
        "GET",
        pdf_url,
        headers={"Accept": "application/pdf,*/*", "Referer": page_url},
        timeout=120,
    ) as resp:
        for chunk in resp.iter_bytes(256 * 1024):
            if not first_bytes:
                first_bytes = chunk[:8]
            size += len(chunk)
        status_code = resp.status_code
        content_type = resp.headers.get("content-type", "")
    return PdfCheck(label, page_url, pdf_url, status_code, content_type, size, first_bytes)


def _anti_bot_probe(client: httpx.Client, config: dict[str, Any], headers: dict[str, str]) -> tuple[int, list[int]]:
    base_url = config["endpoints"]["ANNOUNCEMENTS_API_URL"] + "securitycode/count"
    statuses = []
    for _ in range(10):
        resp = client.get(
            base_url,
            headers=headers,
            params={
                "value": "D05",
                "periodstart": "20240101_000000",
                "periodend": "20241231_235959",
            },
        )
        statuses.append(resp.status_code)
        time.sleep(0.5)
    no_token = client.get(
        config["endpoints"]["ANNOUNCEMENTS_API_URL"] + "securitycode",
        params={
            "value": "D05",
            "periodstart": "20240101_000000",
            "periodend": "20241231_235959",
        },
    )
    return no_token.status_code, statuses


def main() -> None:
    variants = ["D05", "d05", "D05.SI", "d05.si"]
    print("Ticker normalization:", {value: _normalize_ticker(value) for value in variants})

    with httpx.Client(headers=BASE_HEADERS, timeout=60, follow_redirects=True) as client:
        config = _get_config(client)
        token = _get_authorization_token(client, config)
        api_headers = {**BASE_HEADERS, "authorizationToken": token}
        print("CMS_VERSION:", config["CMS_VERSION"])
        print("ANNOUNCEMENTS_API_URL:", config["endpoints"]["ANNOUNCEMENTS_API_URL"])
        print("FINANCIAL_REPORTS_API_URL:", config["endpoints"]["FINANCIAL_REPORTS_API_URL"])
        print("IPO_PROSPECTUS_API_URL:", config["endpoints"]["IPO_PROSPECTUS_API_URL"])
        print("authorizationToken length:", len(token))

        rows = _fetch_financial_reports(client, config, api_headers)
        annual_specs = [
            ("DBS 2024 annual", "DBS GROUP HOLDINGS LTD", 2024),
            ("UOB 2024 annual", "UNITED OVERSEAS BANK LIMITED", 2024),
            ("Singtel 2024 annual", "SINGTEL", 2024),
        ]
        checks: list[PdfCheck] = []
        for label, security_name, year in annual_specs:
            row = _find_annual(rows, security_name=security_name, report_year=year)
            links = _extract_pdf_links(client, row["url"])
            pdf_url = _choose_pdf("annual", links)
            checks.append(_measure_pdf(client, label, row["url"], pdf_url))

        h1_row = _find_h1(_fetch_announcements(client, config, api_headers, "U11", 2024, 2024), 2024)
        h1_pdf = _choose_pdf("h1", _extract_pdf_links(client, h1_row["url"]))
        checks.append(_measure_pdf(client, "UOB 2024 H1", h1_row["url"], h1_pdf))

        ipo_row = _find_ipo(_fetch_ipos(client, config, api_headers), 2024)
        ipo_pdf = _choose_pdf("ipo", _extract_pdf_links(client, ipo_row["url"]))
        checks.append(_measure_pdf(client, f"{ipo_row['name']} IPO", ipo_row["url"], ipo_pdf))

        no_token_status, burst_statuses = _anti_bot_probe(client, config, api_headers)
        print("No-token status:", no_token_status)
        print("10-request burst statuses:", burst_statuses)
        for check in checks:
            print(
                f"{check.label}: {check.status_code} {check.content_type} "
                f"{check.size_bytes} bytes first={check.first_bytes!r}"
            )
            if check.status_code != 200 or not check.content_type.startswith("application/pdf"):
                raise RuntimeError(f"PDF check failed for {check.label}: {check}")
            if not check.first_bytes.startswith(b"%PDF-"):
                raise RuntimeError(f"PDF signature check failed for {check.label}: {check.first_bytes!r}")


if __name__ == "__main__":
    main()

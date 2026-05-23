"""EDINET public-web fallback hooks for no-key JP mode.

The official EDINET API v2 currently returns an authentication error without a
subscription key. Keeping this module separate mirrors the KR public fallback
shape and gives the web scraper a single place to evolve if EDINET changes the
public search form again.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

from loguru import logger

from . import edinet_api

_PUBLIC_SEARCH_URL = "https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx"

_SELECTORS = {
    "keyword_input": "input[name='W0018vD_KEYWORD']",
    "search_button": "input[name='W0018BTNBTN_SEARCH']",
    "results_rows": "table.Grid tr",
    "pdf_link": "a[onclick*='PdfClick']",
}


def list_documents(submit_date: str | date) -> list[dict[str, Any]]:
    try:
        return edinet_api.list_documents(submit_date, api_key=None)
    except PermissionError as exc:
        logger.info(f"EDINET public list unavailable without API key: {exc}")
        return []


def download_document_pdf(doc_id: str, dest: Path) -> int:
    return edinet_api.download_document_pdf(doc_id, dest, api_key=None)

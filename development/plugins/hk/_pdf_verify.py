from __future__ import annotations

import io

import httpx
from pypdf import PdfReader

from app.core.http import default_client


_KEYWORDS_BY_KIND = {
    "annual": ("年報", "年度報告", "年报", "年度报告", "Annual Report"),
    "interim": ("中期", "半年", "Interim", "Half-Year", "Half Year"),
    "q1": ("第一季度", "First Quarter", "First Quarterly"),
    "q3": ("第三季度", "Third Quarter", "Third Quarterly"),
}


def verify_pdf_year_and_kind(url: str, target_year: int, target_kind: str) -> bool:
    """Check the first pages of an HKEX PDF for the expected year and report kind."""
    try:
        timeout = httpx.Timeout(connect=30.0, read=60.0, write=30.0, pool=30.0)
        with default_client(source="hkexnews") as client:
            response = client.get(url, timeout=timeout)
            response.raise_for_status()
            pdf_bytes = response.content[: 5 * 1024 * 1024]

        reader = PdfReader(io.BytesIO(pdf_bytes))
        texts: list[str] = []
        for idx in range(min(3, len(reader.pages))):
            texts.append(reader.pages[idx].extract_text() or "")
        text = "\n".join(texts)
    except Exception:
        return False

    keywords = _KEYWORDS_BY_KIND.get(target_kind, ())
    return str(target_year) in text and any(keyword in text for keyword in keywords)

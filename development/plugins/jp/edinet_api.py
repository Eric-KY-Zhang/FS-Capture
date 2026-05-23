"""EDINET official API v2 client for JP filings."""

from __future__ import annotations

import zipfile
from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Any

import httpx
from loguru import logger

from app.core.cache import cached_or_load
from app.core.http import default_client
from app.core.ratelimit import limiter
from app.core.settings import load_settings

_BASE_URL = "https://api.edinet-fsa.go.jp/api/v2"
_DOCUMENTS_URL = f"{_BASE_URL}/documents.json"
_DOCUMENT_URL = f"{_BASE_URL}/documents/{{doc_id}}"
_TTL = 24 * 3600


def _edinet_rate() -> float:
    return load_settings().rate_limits.edinet


def _headers(api_key: str | None) -> dict[str, str]:
    headers = {
        "Accept": "application/json,application/pdf,application/octet-stream,*/*",
    }
    if api_key:
        # EDINET documents specify Subscription-Key as a request parameter, but
        # keeping the header too makes the branch explicit and easy to assert.
        headers["Subscription-Key"] = api_key
    return headers


def _params(*, api_key: str | None, request_type: str) -> dict[str, str]:
    params = {"type": request_type}
    if api_key:
        params["Subscription-Key"] = api_key
    return params


def _normalize_sec_code(value: Any) -> str:
    digits = "".join(ch for ch in str(value or "") if ch.isdigit())
    if len(digits) == 5 and digits.endswith("0"):
        return digits[:4]
    return digits


def _normalize_row(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "doc_id": str(row.get("docID") or row.get("doc_id") or "").strip(),
        "doc_type_code": str(row.get("docTypeCode") or row.get("doc_type_code") or "").strip(),
        "submit_date_time": str(
            row.get("submitDateTime") or row.get("submit_date_time") or ""
        ).strip(),
        "period_end": str(row.get("periodEnd") or row.get("period_end") or "").strip(),
        "edinet_code": str(row.get("edinetCode") or row.get("edinet_code") or "").strip(),
        "sec_code": _normalize_sec_code(row.get("secCode") or row.get("sec_code")),
        "jcn": str(row.get("JCN") or row.get("jcn") or "").strip(),
        "filer_name": str(row.get("filerName") or row.get("filer_name") or "").strip(),
        "title": str(row.get("docDescription") or row.get("title") or "").strip(),
        "pdf_flag": str(row.get("pdfFlag") or row.get("pdf_flag") or "").strip(),
    }


def _raise_for_payload(payload: Any, *, context: str) -> None:
    if not isinstance(payload, dict):
        return
    status = str(
        payload.get("StatusCode")
        or payload.get("statusCode")
        or (payload.get("metadata") or {}).get("status")
        or ""
    )
    if status and status not in {"200", "OK"}:
        message = str(payload.get("message") or (payload.get("metadata") or {}).get("message") or "")
        if status in {"401", "403"}:
            raise PermissionError(f"EDINET {context} requires a valid API key: {message}")
        raise RuntimeError(f"EDINET {context} failed with status {status}: {message}")


def list_documents(submit_date: str | date, *, api_key: str | None = None) -> list[dict[str, Any]]:
    day = submit_date.isoformat() if isinstance(submit_date, date) else str(submit_date)
    cache_key = f"jp:edinet:documents:v1:{day}:{'key' if api_key else 'public'}"

    def _fetch() -> list[dict[str, Any]]:
        params = {"date": day, **_params(api_key=api_key, request_type="2")}
        with default_client(source="edinet", timeout=60.0) as client:
            limiter("edinet", _edinet_rate()).acquire_blocking()
            response = client.get(_DOCUMENTS_URL, params=params, headers=_headers(api_key))
            response.raise_for_status()
            payload = response.json()
        _raise_for_payload(payload, context="documents list")
        rows = payload.get("results") if isinstance(payload, dict) else []
        if not isinstance(rows, list):
            rows = []
        return [_normalize_row(row) for row in rows if isinstance(row, dict)]

    return cached_or_load(cache_key, _fetch, expire=_TTL)


def _write_pdf_or_zip(content: bytes, dest: Path) -> int:
    if content.startswith(b"%PDF"):
        dest.write_bytes(content)
        return dest.stat().st_size
    if content.startswith(b"PK"):
        with zipfile.ZipFile(BytesIO(content)) as archive:
            for name in archive.namelist():
                if name.lower().endswith(".pdf"):
                    dest.write_bytes(archive.read(name))
                    return dest.stat().st_size
    raise ValueError("EDINET document response was neither PDF nor ZIP-with-PDF")


def download_document_pdf(doc_id: str, dest: Path, *, api_key: str | None = None) -> int:
    url = _DOCUMENT_URL.format(doc_id=doc_id)
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")
    tmp_dest.unlink(missing_ok=True)
    params = _params(api_key=api_key, request_type="2")
    try:
        with default_client(source="edinet", timeout=httpx.Timeout(120.0, read=180.0)) as client:
            limiter("edinet", _edinet_rate()).acquire_blocking()
            response = client.get(url, params=params, headers=_headers(api_key))
            response.raise_for_status()
            content_type = response.headers.get("content-type", "")
            if "application/json" in content_type:
                payload = response.json()
                _raise_for_payload(payload, context=f"document {doc_id}")
                raise RuntimeError(f"EDINET document {doc_id} returned JSON instead of PDF")
            n_bytes = _write_pdf_or_zip(response.content, tmp_dest)
        tmp_dest.replace(dest)
        return n_bytes
    except Exception as exc:
        tmp_dest.unlink(missing_ok=True)
        logger.warning(f"EDINET PDF download failed for {doc_id}: {exc}")
        raise

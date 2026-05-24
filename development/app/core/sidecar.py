from __future__ import annotations

import hashlib
import json
import uuid
from collections.abc import Iterator
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from loguru import logger

from app.core.models import Exchange, ReportFile
from app.core.settings import load_settings


def _sidecar_root(cache_root: Path | None = None) -> Path:
    root = cache_root if cache_root is not None else load_settings().cache_path()
    return root / "sidecars"


def _stem_for(report: ReportFile) -> str:
    return (report.accession_number or Path(report.local_path).stem).strip()


def sidecar_path(report: ReportFile, cache_root: Path | None = None) -> Path:
    return (
        _sidecar_root(cache_root)
        / report.ticker.exchange.value
        / f"{_stem_for(report)}.meta.json"
    )


def write_sidecar(report: ReportFile, cache_root: Path | None = None) -> Path:
    path = sidecar_path(report, cache_root)
    path.parent.mkdir(parents=True, exist_ok=True)

    pdf = Path(report.local_path)
    file_size = pdf.stat().st_size if pdf.exists() else report.file_size_bytes
    sha256 = hashlib.sha256(pdf.read_bytes()).hexdigest() if pdf.exists() else ""
    period = report.period

    payload: dict[str, Any] = {
        "exchange": report.ticker.exchange.value,
        "ticker_code": report.ticker.code,
        "ticker_name": report.ticker.name,
        "period_year": period.year if period else None,
        "period_type": period.type.value if period else None,
        "kind": report.kind,
        "title": report.title,
        "source_url": report.source_url,
        "downloaded_at": datetime.now(UTC).isoformat(),
        "file_size_bytes": file_size,
        "sha256": sha256,
    }
    tmp_path = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        tmp_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        tmp_path.replace(path)
    finally:
        tmp_path.unlink(missing_ok=True)
    return path


def read_sidecar(
    stem: str,
    exchange: Exchange,
    cache_root: Path | None = None,
) -> dict[str, Any] | None:
    path = _sidecar_root(cache_root) / exchange.value / f"{stem}.meta.json"
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def iter_sidecars(
    exchange: Exchange | None = None,
    cache_root: Path | None = None,
) -> Iterator[dict[str, Any]]:
    root = _sidecar_root(cache_root)
    glob_target = root / exchange.value if exchange else root
    if not glob_target.exists():
        return
    for path in glob_target.rglob("*.meta.json"):
        try:
            yield json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            logger.warning(f"sidecar read failed for {path}: {exc}")


def migrate_legacy_sidecars(output_root: Path, cache_root: Path | None = None) -> int:
    """Move legacy output-adjacent sidecars into cache/sidecars."""
    moved = 0
    if not output_root.exists():
        return moved

    for old in output_root.rglob("*.meta.json"):
        try:
            payload = json.loads(old.read_text(encoding="utf-8"))
            exchange_code = payload.get("exchange")
            if not exchange_code:
                logger.warning(f"legacy sidecar missing exchange, skipped: {old}")
                continue

            name = (
                f"{old.name.removesuffix('.pdf.meta.json')}.meta.json"
                if old.name.endswith(".pdf.meta.json")
                else old.name
            )
            new_path = _sidecar_root(cache_root) / str(exchange_code) / name
            new_path.parent.mkdir(parents=True, exist_ok=True)
            old.replace(new_path)
            moved += 1
        except (OSError, json.JSONDecodeError) as exc:
            logger.warning(f"sidecar migration failed for {old}: {exc}")
    return moved

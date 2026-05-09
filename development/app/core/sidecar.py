from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from app.core.models import ReportFile


def write_sidecar(report: ReportFile) -> Path:
    """Write a JSON metadata sidecar next to a downloaded disclosure file."""
    pdf_path = Path(report.local_path)
    sidecar = pdf_path.with_suffix(pdf_path.suffix + ".meta.json")

    file_size = pdf_path.stat().st_size
    sha256 = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
    period = report.period

    meta: dict[str, Any] = {
        "exchange": report.ticker.exchange.value,
        "ticker_code": report.ticker.code,
        "ticker_name": report.ticker.name,
        "period_year": period.year if period else None,
        "period_type": period.type.value if period else None,
        "kind": report.kind,
        "title": report.title,
        "source_url": report.source_url,
        "downloaded_at": datetime.now(timezone.utc).isoformat(),
        "file_size_bytes": file_size,
        "sha256": sha256,
    }
    sidecar.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return sidecar

from __future__ import annotations

import re
from pathlib import Path

from .models import Period, Ticker


_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def safe_filename(value: str, *, fallback: str = "file") -> str:
    cleaned = _BAD_FILENAME_CHARS.sub("_", value).strip(" ._")
    return cleaned or fallback


def report_output_path(output_root: Path, ticker: Ticker, period: Period, kind: str, suffix: str) -> Path:
    """Flat report path under output_root with enough metadata in the filename."""
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    name = safe_filename(
        f"{ticker.exchange.value}_{ticker.code}_{period.year}_{period.type.value}_{kind}{suffix}"
    )
    return output_root / name

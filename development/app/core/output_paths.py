from __future__ import annotations

import re
from pathlib import Path

from .models import Period, Ticker

_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def safe_filename(value: str, *, fallback: str = "file") -> str:
    cleaned = _BAD_FILENAME_CHARS.sub("_", value).strip(" ._")
    return cleaned or fallback


def report_output_path(
    output_root: Path, ticker: Ticker, period: Period, kind: str, suffix: str
) -> Path:
    """Flat report path under output_root with enough metadata in the filename."""
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    name = safe_filename(
        f"{ticker.exchange.value}_{ticker.code}_{period.year}_{period.type.value}_{kind}{suffix}"
    )
    return output_root / name


def report_output_path_for_filing(
    output_root: Path,
    ticker: Ticker,
    family: str,
    sequence: int | None,
    label: str,
    suffix: str,
    *,
    filing_date: str | None = None,
    is_amendment: bool = False,
    source_id: str | None = None,
) -> Path:
    """Flat path for non-periodic disclosure documents such as IPO prospectuses."""
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    seq = f"{sequence:03d}" if sequence is not None else "000"
    parts = [ticker.exchange.value, ticker.code, family]
    if filing_date:
        parts.append(filing_date)
    parts.append(seq)
    if is_amendment:
        parts.append("amendment")
    parts.append(label)
    if source_id:
        parts.append(source_id)
    name = safe_filename("_".join(str(p) for p in parts if p) + suffix)
    return output_root / name

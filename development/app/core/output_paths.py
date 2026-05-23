from __future__ import annotations

import re
import time
from pathlib import Path

from .models import Period, PeriodType, Ticker

_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Map internal kind strings → 中文 label used in filenames. Anything not listed
# falls back to the raw kind string (already filename-safe for ASCII labels).
_KIND_ZH = {
    "annual_report": "年报",
    "audit_report": "审计报告",
    "interim_report": "半年报",
    "q1_report": "一季报",
    "q3_report": "三季报",
    "ipo_prospectus": "招股书",
    "ipo_document": "招股书",
}

# Map PeriodType → 中文 label (used when kind is not in the table above, e.g.
# unrecognised plugin-specific kinds where we still want a Chinese period hint).
_PERIOD_ZH = {
    PeriodType.ANNUAL: "年报",
    PeriodType.Q1: "一季报",
    PeriodType.Q2: "半年报",
    PeriodType.Q3: "三季报",
    PeriodType.IPO_PROSPECTUS: "招股书",
}


def safe_filename(value: str, *, fallback: str = "file") -> str:
    cleaned = _BAD_FILENAME_CHARS.sub("_", value).strip(" ._")
    return cleaned or fallback


def _short_name(ticker: Ticker) -> str:
    """Company short name for filename. Falls back to '' if unresolved.

    The name was captured during ``resolve_name`` (the user-blocking 确认 step
    in the UI), so it's expected to be present whenever a report download fires.
    """
    if ticker.name:
        return safe_filename(ticker.name, fallback="")
    return ""


def _kind_label(kind: str, period: Period) -> str:
    """Resolve the trailing report-type segment.

    Prefer the Chinese mapping for known kinds; otherwise fall back to the
    period-type label, then to the raw kind string.
    """
    label = _KIND_ZH.get(kind)
    if label:
        return label
    label = _PERIOD_ZH.get(period.type)
    if label:
        return label
    return safe_filename(kind, fallback="report")


def report_output_path(
    output_root: Path, ticker: Ticker, period: Period, kind: str, suffix: str
) -> Path:
    """Flat report path under output_root.

    Format: ``{market}_{code}_{name}_{year}_{kind_zh}{suffix}``
    Example: ``A_600519_贵州茅台_2024_年报.pdf``
    """
    suffix = suffix if suffix.startswith(".") else f".{suffix}"
    name_seg = _short_name(ticker)
    kind_seg = _kind_label(kind, period)

    parts = [ticker.exchange.value, ticker.code]
    if name_seg:
        parts.append(name_seg)
    parts.extend([str(period.year), kind_seg])

    return output_root / safe_filename("_".join(parts) + suffix)


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
    """Flat path for non-periodic disclosure documents (e.g. IPO prospectuses).

    Format: ``{market}_{code}_{name}_{kind_zh}_{filing_date}_{seq}[_补充]{suffix}``
    Example: ``A_600519_贵州茅台_招股书_2024-04-01_001.pdf``

    ``family`` ("ipo" / etc.) and ``label`` are still accepted for backward
    compatibility but rolled into the Chinese kind segment. We treat any
    family containing ``ipo`` (or any label mapped via ``_KIND_ZH``) as
    招股书; otherwise fall back to ``family``.
    """
    suffix = suffix if suffix.startswith(".") else f".{suffix}"

    # Resolve the report-type segment.
    kind_seg = (
        _KIND_ZH.get(label)
        or _KIND_ZH.get(family)
        or (
            "招股书"
            if "ipo" in (family or "").lower() or "ipo" in (label or "").lower()
            else safe_filename(label or family or "report", fallback="report")
        )
    )

    name_seg = _short_name(ticker)
    parts = [ticker.exchange.value, ticker.code]
    if name_seg:
        parts.append(name_seg)
    parts.append(kind_seg)
    if filing_date:
        parts.append(filing_date)
    if sequence is not None:
        parts.append(f"{sequence:03d}")
    if is_amendment:
        parts.append("补充")
    if source_id:
        # Optional disambiguation tail, e.g. the raw filename/accession number.
        parts.append(source_id)

    return output_root / safe_filename("_".join(str(p) for p in parts if p) + suffix)


def cleanup_stale_parts(output_root: Path, max_age_days: int = 7) -> int:
    """Delete orphan ``.part`` download files older than ``max_age_days``."""
    if not output_root.exists():
        return 0
    cutoff = time.time() - max_age_days * 24 * 3600
    removed = 0
    for path in output_root.rglob("*.part"):
        try:
            if path.is_file() and path.stat().st_mtime < cutoff:
                path.unlink()
                removed += 1
        except OSError:
            continue
    return removed

"""
Phase 4f Step 6: compare 原币 vs 统一RMB workbook dumps.

Expected use:
  py tools/diff_phase4f_rmb.py samples/regression_phase4f_yuanbi.json samples/regression_phase4f_rmb.json

The dump format is a JSON object of:
  {
    "SheetName": [[cell_value, ...], ...],
    ...
  }

Expected interpretation for the Phase 4f six-company scenario:
  - A股 sheets should be byte-identical because reporting currency is RMB.
  - 港股 02519 should also be byte-identical because reporting currency is RMB.
  - Non-RMB markets should differ by a stable FX ratio for numeric cells.
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any


IDENTITY_PREFIXES = ("A股_",)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("yuanbi", type=Path)
    parser.add_argument("rmb", type=Path)
    parser.add_argument("--identity-sheet", action="append", default=[])
    parser.add_argument("--tolerance", type=float, default=1e-9)
    return parser.parse_args()


def normalize(value: Any) -> Any:
    if isinstance(value, str):
        value = value.strip()
        return None if value == "" else value
    return value


def close(left: Any, right: Any, tolerance: float) -> bool:
    left = normalize(left)
    right = normalize(right)
    if left == right:
        return True
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(float(left), float(right), rel_tol=0, abs_tol=tolerance)
    return False


def numeric_ratio(left: Any, right: Any) -> float | None:
    left = normalize(left)
    right = normalize(right)
    if not isinstance(left, (int, float)) or not isinstance(right, (int, float)):
        return None
    if float(left) == 0:
        return None
    return float(right) / float(left)


def compare_sheet(name: str, left_rows: list[list[Any]], right_rows: list[list[Any]], tolerance: float) -> dict[str, Any]:
    mismatches: list[tuple[str, Any, Any]] = []
    ratios: list[float] = []
    row_count = max(len(left_rows), len(right_rows))
    for r in range(row_count):
        left_row = left_rows[r] if r < len(left_rows) else []
        right_row = right_rows[r] if r < len(right_rows) else []
        col_count = max(len(left_row), len(right_row))
        for c in range(col_count):
            left = left_row[c] if c < len(left_row) else None
            right = right_row[c] if c < len(right_row) else None
            if not close(left, right, tolerance):
                mismatches.append((f"R{r + 1}C{c + 1}", normalize(left), normalize(right)))
                ratio = numeric_ratio(left, right)
                if ratio is not None:
                    ratios.append(ratio)

    out: dict[str, Any] = {
        "mismatches": len(mismatches),
        "examples": mismatches[:10],
    }
    if ratios:
        out["numeric_ratio_min"] = min(ratios)
        out["numeric_ratio_max"] = max(ratios)
        out["numeric_ratio_avg"] = sum(ratios) / len(ratios)
    return out


def main() -> int:
    args = parse_args()
    yuanbi = json.loads(args.yuanbi.read_text(encoding="utf-8"))
    rmb = json.loads(args.rmb.read_text(encoding="utf-8"))

    total_failures = 0
    sheets = sorted(set(yuanbi) | set(rmb))
    identity_sheets = set(args.identity_sheet)
    for sheet in sheets:
        left_rows = yuanbi.get(sheet, [])
        right_rows = rmb.get(sheet, [])
        result = compare_sheet(sheet, left_rows, right_rows, args.tolerance)
        expected_identity = sheet.startswith(IDENTITY_PREFIXES) or sheet in identity_sheets
        mismatches = int(result["mismatches"])
        status = "OK"
        if expected_identity and mismatches:
            status = "FAIL"
            total_failures += 1
        print(f"{sheet}: {mismatches} mismatches [{status}]")
        if result.get("numeric_ratio_avg") is not None:
            print(
                "  numeric ratio rmb/yuanbi: "
                f"min={result['numeric_ratio_min']:.8g}, "
                f"max={result['numeric_ratio_max']:.8g}, "
                f"avg={result['numeric_ratio_avg']:.8g}"
            )
        for addr, left, right in result["examples"]:
            print(f"  {addr}: 原币={left!r}, RMB={right!r}")

    if total_failures:
        print(f"TOTAL: {total_failures} identity-sheet failures")
        return 1
    print("TOTAL: identity-sheet checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

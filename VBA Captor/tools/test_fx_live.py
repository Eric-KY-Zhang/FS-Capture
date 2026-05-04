"""
Phase 4f Step 2 live verification: install fresh modules + invoke
EnsureFxRateCached / GetFxRate over the network, report the results.

Pre-conditions:
  - 上市公司财务数据查询.xlsm exists in repo root (built by build_template.py
    + install_modules.py at least once)
  - The workbook is NOT currently open in Excel (script aborts if ~$ lock file
    detected — close Excel and rerun)
  - Trust access to VBA project model is enabled in Excel
  - Internet works; xueqiu.com reachable

Usage:
  cd .../VBA Captor
  py tools/test_fx_live.py [--skip-install]

Output: prints the FX sheet rows written by EnsureFxRateCached for
USD/HKD/KRW × multiple periodEnds, plus GetFxRate() round-trip.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import time
from pathlib import Path

import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
XLSM = ROOT / "上市公司财务数据查询.xlsm"
LOCK = ROOT / f"~${XLSM.name}"

# Fallback source: when running from a worktree, the user's xlsm lives in the
# main repo at .claude/worktrees/<wt>/VBA Captor/  ←→  VBA Captor/ (3 levels up)
SOURCE_CANDIDATES = [
    ROOT.parent.parent.parent / "VBA Captor" / "上市公司财务数据查询.xlsm",
    Path("E:/Claude+CODEX Project/FS Capture/VBA Captor/上市公司财务数据查询.xlsm"),
]


# (curCode, periodEnd, expected_eop_range, expected_avg_range)
TEST_CASES = [
    ("USD", "2024-12-31", (6.8, 7.5), (6.8, 7.5)),
    ("HKD", "2024-12-31", (0.85, 0.95), (0.85, 0.95)),
    ("KRW", "2024-12-31", (0.0040, 0.0060), (0.0040, 0.0060)),
    ("USD", "2023-12-31", (6.8, 7.5), (6.8, 7.5)),
    ("HKD", "2023-12-31", (0.85, 0.95), (0.85, 0.95)),
]

FX_COL_MAP = {  # (eop_col, avg_col), 1-based
    "USD": (2, 3),
    "HKD": (4, 5),
    "KRW": (6, 7),
}


def find_fx_row(ws_fx, period_end: str) -> int:
    last = ws_fx.Cells(ws_fx.Rows.Count, 1).End(-4162).Row  # xlUp
    for r in range(2, max(last, 2) + 1):
        v = ws_fx.Cells(r, 1).Value
        if v is None:
            continue
        if str(v).strip() == period_end:
            return r
    return 0


def in_range(val, rng) -> bool:
    if val is None:
        return False
    try:
        f = float(val)
    except (TypeError, ValueError):
        return False
    return rng[0] <= f <= rng[1]


def run_install():
    print("=" * 70)
    print("STEP 1: installing modules into workbook")
    print("=" * 70)
    sys.path.insert(0, str(ROOT / "tools"))
    import install_modules  # noqa: E402

    install_modules.main()


def run_live_test():
    print("\n" + "=" * 70)
    print("STEP 2: opening workbook for live FX test")
    print("=" * 70)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    failed_cases = []
    passed_cases = []
    try:
        wb = excel.Workbooks.Open(str(XLSM))
        try:
            # Cookie + currency state
            try:
                pool = wb.Sheets("样本池")
                cookie = str(pool.Range("B5").Value or "")
                b6 = pool.Range("B6").Value
                print(f"\n样本池 B5 cookie length = {len(cookie)}")
                if cookie:
                    print(f"  cookie head = {cookie[:60]!r}...")
                else:
                    print("  WARN: cookie empty — FX call may still work via warmup")
                print(f"样本池 B6 显示币种 = {b6!r}")
            except Exception as e:
                print(f"样本池 read failed: {e}")

            # FX sheet
            sheet_names = [wb.Sheets.Item(i).Name for i in range(1, wb.Sheets.Count + 1)]
            if "汇率" not in sheet_names:
                print(f"\nFATAL: 汇率 sheet missing. Available sheets: {sheet_names}")
                return False
            ws_fx = wb.Sheets("汇率")

            print(f"\n汇率 sheet header (Row 1):")
            for c in range(1, 9):
                v = ws_fx.Cells(1, c).Value
                print(f"  Col {c}: {v!r}")

            initial_last = ws_fx.Cells(ws_fx.Rows.Count, 1).End(-4162).Row
            print(f"\n汇率 sheet existing data rows: rows 2..{initial_last}")
            for r in range(2, max(initial_last, 1) + 1):
                row_vals = [ws_fx.Cells(r, c).Value for c in range(1, 9)]
                print(f"  Row {r}: {row_vals}")

            # Run cases
            print("\n" + "-" * 70)
            print("Invoking EnsureFxRateCached for each (cur, period)")
            print("-" * 70)
            for cur, period, eop_rng, avg_rng in TEST_CASES:
                tag = f"{cur}@{period}"
                t0 = time.time()
                try:
                    ok = excel.Run("模块_抓汇率.EnsureFxRateCached", period, cur)
                except Exception as e:
                    print(f"\n[{tag}] EnsureFxRateCached RAISED: {e}")
                    failed_cases.append((tag, f"RAISED: {e}"))
                    continue
                elapsed = time.time() - t0
                print(f"\n[{tag}] returned {ok!r} in {elapsed:.2f}s")

                # Read back
                row = find_fx_row(ws_fx, period)
                if row == 0:
                    print(f"  ! No 汇率 row found for periodEnd={period}")
                    failed_cases.append((tag, f"row missing for {period}"))
                    continue
                eop_col, avg_col = FX_COL_MAP[cur]
                eop_val = ws_fx.Cells(row, eop_col).Value
                avg_val = ws_fx.Cells(row, avg_col).Value
                print(f"  Row {row}: A={ws_fx.Cells(row, 1).Value!r}  "
                      f"EOP[col{eop_col}]={eop_val!r}  AVG[col{avg_col}]={avg_val!r}")

                eop_ok = in_range(eop_val, eop_rng)
                avg_ok = in_range(avg_val, avg_rng)
                print(f"  EOP in {eop_rng}: {eop_ok}  |  AVG in {avg_rng}: {avg_ok}")
                if eop_ok and avg_ok:
                    passed_cases.append(tag)
                else:
                    failed_cases.append(
                        (tag, f"EOP={eop_val!r} avg={avg_val!r} ranges {eop_rng}/{avg_rng}")
                    )

            # GetFxRate round-trip
            print("\n" + "-" * 70)
            print("Round-trip via GetFxRate (utility wrapper)")
            print("-" * 70)
            for cur, period, *_ in TEST_CASES:
                try:
                    eop = excel.Run("模块_工具函数.GetFxRate", cur, period, True)
                    avg = excel.Run("模块_工具函数.GetFxRate", cur, period, False)
                    print(f"  GetFxRate({cur!r}, {period!r}, eop=T) = {eop!r}")
                    print(f"  GetFxRate({cur!r}, {period!r}, eop=F) = {avg!r}")
                except Exception as e:
                    print(f"  GetFxRate({cur!r}, {period!r}) RAISED: {e}")

            # RMB/CNY short-circuit
            print("\nRMB short-circuit check:")
            for cur in ("RMB", "CNY"):
                rate = excel.Run("模块_工具函数.GetFxRate", cur, "2024-12-31", True)
                print(f"  GetFxRate({cur!r}, '2024-12-31', eop=T) = {rate!r}  (expect 1.0)")

            # Cache hit check (call USD@2024-12-31 again, should be fast)
            print("\nCache hit check (re-call USD@2024-12-31):")
            t0 = time.time()
            ok = excel.Run("模块_抓汇率.EnsureFxRateCached", "2024-12-31", "USD")
            elapsed = time.time() - t0
            print(f"  returned {ok!r} in {elapsed:.2f}s  (expect <0.5s = cache hit no PS)")

            # Save so the cache persists for the user
            try:
                wb.Save()
                print("\n+ Workbook saved with new FX cache rows")
            except Exception as e:
                print(f"\n! Save failed: {e}")
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"Passed: {len(passed_cases)}/{len(TEST_CASES)}")
    for c in passed_cases:
        print(f"  + {c}")
    if failed_cases:
        print(f"Failed: {len(failed_cases)}")
        for tag, why in failed_cases:
            print(f"  - {tag}: {why}")
    return len(failed_cases) == 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-install", action="store_true", help="skip install_modules step")
    args = ap.parse_args()

    if not XLSM.exists():
        # In a worktree the xlsm lives in the main repo. Copy a fresh snapshot
        # so install + the live test stay isolated from the user's open file.
        for src in SOURCE_CANDIDATES:
            if src.exists():
                print(f"+ Copying {src} -> {XLSM}")
                shutil.copy2(src, XLSM)
                break
        else:
            print(f"FATAL: no source xlsm found. Tried {SOURCE_CANDIDATES}")
            sys.exit(1)

    if LOCK.exists():
        print(f"FATAL: lock file detected next to worktree xlsm: {LOCK.name}")
        print(f"Close Excel and re-run.")
        sys.exit(1)

    if not args.skip_install:
        run_install()

    ok = run_live_test()
    sys.exit(0 if ok else 2)


if __name__ == "__main__":
    main()

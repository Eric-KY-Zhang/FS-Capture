"""
Phase 4f Step 3 lite reviewer test:
  Local A股-shaped smoke test for WriteWideTable.
  B6=原币 vs B6=统一RMB, expect byte-identical because A股 reports RMB.

This avoids live Sina/Xueqiu/StockAnalysis fetches. Progress is printed to stderr
(always unbuffered) so we can see where COM calls fail.
"""
from __future__ import annotations
import sys
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
XLSM = ROOT / "上市公司财务数据查询.xlsm"


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def dump_sheet(ws):
    used = ws.UsedRange
    if used.Rows.Count == 1 and used.Columns.Count == 1:
        return [[used.Value]]
    raw = used.Value
    return [list(r) for r in raw]


def main():
    log("=== open Excel ===")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    try:
        log(f"=== open workbook {XLSM.name} ===")
        wb = excel.Workbooks.Open(str(XLSM))
        try:
            log("=== verify win32com Optional Boolean argument ===")
            bool_echo = excel.Run("模块_测试.TestOptionalBool", True)
            log(f"=== TestOptionalBool(True) -> {bool_echo!r} ===")

            log("=== invoking 模块_测试.TestStep3Smoke ===")
            excel.Run("模块_测试.TestStep3Smoke")
            log("=== smoke macro done ===")

            log("=== invoking 模块_测试.TestStep45Smoke ===")
            excel.Run("模块_测试.TestStep45Smoke")
            log("=== Step 4/5 smoke macro done ===")

            ws_yuanbi = wb.Sheets("_phase4f_step3_yuanbi")
            ws_rmb = wb.Sheets("_phase4f_step3_rmb")
            a = dump_sheet(ws_yuanbi)
            b = dump_sheet(ws_rmb)
            log(f"=== dumped 原币 scratch: {len(a)}r x {len(a[0]) if a else 0}c ===")
            log(f"=== dumped 统一RMB scratch: {len(b)}r x {len(b[0]) if b else 0}c ===")
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        log("=== Excel quit ===")

    log("\n=== DIFF ===")
    diffs = []
    if len(a) != len(b):
        diffs.append(("rows", len(a), len(b)))
    for r in range(min(len(a), len(b))):
        if len(a[r]) != len(b[r]):
            diffs.append((f"r{r+1} cols", len(a[r]), len(b[r])))
            continue
        for c in range(len(a[r])):
            if a[r][c] != b[r][c]:
                if isinstance(a[r][c], (int, float)) and isinstance(b[r][c], (int, float)):
                    if abs(a[r][c] - b[r][c]) < 1e-9:
                        continue
                diffs.append((f"r{r+1}c{c+1}", a[r][c], b[r][c]))
    log(f"\n=== A股_资产负债表 diff: {len(diffs)} mismatches ===")
    for d in diffs[:20]:
        log(f"  {d}")
    if len(diffs) == 0:
        log("\n*** PASS: A股 toggle byte-identical (proves displayMode short-circuit works) ***")
        return 0
    log("\n*** FAIL: A股 toggle differs - investigate Codex hook logic ***")
    return 1


if __name__ == "__main__":
    sys.exit(main())

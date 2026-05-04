"""Phase 4h workbook state inspector.

# Phase 4j.1: 移除已删除的跨市场 BS/IS/CF sheet + 跨市场对比按钮检查
"""

from __future__ import annotations

import sys
from pathlib import Path

import win32com.client as win32


ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "上市公司财务数据查询.xlsm"
US_MODULE = ROOT / "modules" / "模块_抓美股财报.bas"


def cell_text(ws, row: int, col: int) -> str:
    return str(ws.Cells(row, col).Text)


def main() -> int:
    print("=== Phase 4h state inspect ===", flush=True)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    failures: list[str] = []

    try:
        print("\n[0] stockanalysis fallback auto path", flush=True)
        us_code = US_MODULE.read_text(encoding="utf-8")
        for ticker in ("BABA", "JD", "PDD"):
            if f'"{ticker}"' not in us_code:
                failures.append(f"stockanalysis fallback whitelist missing {ticker}")
        if "ReadStockAnalysisFallbackEnabled" in us_code:
            failures.append("US fallback still depends on manual toggle helper")
        print("stockanalysis fallback auto path whitelist includes BABA/JD/PDD; EDGAR/雪球 failure should auto fallback")

        wb = excel.Workbooks.Open(str(BOOK))
        excel.Run("模块_工具函数.SetSilentMode", True)

        print("\n[1] cross-market sheets", flush=True)
        sheet_specs = [
            ("跨市场_指标表", 4),
        ]
        for sheet_name, data_col in sheet_specs:
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                failures.append(f"missing sheet: {sheet_name}")
                continue
            used = ws.UsedRange
            r1 = [cell_text(ws, 1, c) for c in range(1, min(12, used.Columns.Count) + 1)]
            r2 = [cell_text(ws, 2, c) for c in range(1, min(12, used.Columns.Count) + 1)]
            formulas = []
            for r in range(3, min(5, used.Rows.Count) + 1):
                formulas.append((cell_text(ws, r, 1), cell_text(ws, r, 2), str(ws.Cells(r, data_col).Formula), ws.Cells(r, data_col).Value))
            print(f"{sheet_name}: rows={used.Rows.Count}, cols={used.Columns.Count}")
            print(f"  R1={r1}")
            print(f"  R2={r2}")
            print(f"  rows3_5={formulas}")
            if used.Rows.Count < 3 or used.Columns.Count < data_col:
                failures.append(f"{sheet_name} used range too small")
            if not str(ws.Cells(3, data_col).Formula).startswith("="):
                failures.append(f"{sheet_name} row3 data cell is not a formula")

        print("\n[2] buttons and toggles", flush=True)
        ws_pool = wb.Worksheets("样本池")
        expected_buttons = {
            "BtnClearCache": ("O11:P12", "清空 HTTP 缓存"),
        }
        for name, (addr, caption) in expected_buttons.items():
            try:
                shape = ws_pool.Shapes(name)
                top_left = str(shape.TopLeftCell.Address).replace("$", "")
                bottom_right = str(shape.BottomRightCell.Address).replace("$", "")
                print(f"{name}: {top_left}:{bottom_right}, caption={shape.TextFrame2.TextRange.Text}, action={shape.OnAction}")
                if str(shape.TextFrame2.TextRange.Text) != caption:
                    failures.append(f"{name} caption mismatch")
            except Exception as exc:
                failures.append(f"missing button {name}: {exc}")
        removed_buttons = {
            "BtnBuildCrossAll", "BtnHideCrossMarket",
            "BtnBuildCrossInd", "BtnBuildCrossBS", "BtnBuildCrossIS", "BtnBuildCrossCF",
            "BtnRunBalance", "BtnRunProfit", "BtnRunCash", "BtnRunInd",
            "BtnRunUSBalance", "BtnRunUSProfit", "BtnRunUSCash", "BtnRunUSInd",
            "BtnRunHKBalance", "BtnRunHKProfit", "BtnRunHKCash", "BtnRunHKInd",
            "BtnRunKRBalance", "BtnRunKRProfit", "BtnRunKRCash", "BtnRunKRInd",
        }
        for name in sorted(removed_buttons):
            try:
                ws_pool.Shapes(name)
                failures.append(f"removed button still present: {name}")
            except Exception:
                pass
        print("manual fallback toggle removed; old cross-market and single-table buttons absent")

        print("\n[3] E6 realtime toggle smoke", flush=True)
        saved_b6 = ws_pool.Range("E6").Value
        excel.Run("模块_测试.TestPhase4hToggleSmoke")
        ws_smoke = wb.Worksheets("_phase4h_toggle_smoke")
        ws_pool.Range("E6").Value = "原币"
        excel.Calculate()
        raw_val = ws_smoke.Range("C3").Value
        ws_pool.Range("E6").Value = "统一RMB"
        excel.Calculate()
        rmb_val = ws_smoke.Range("C3").Value
        ws_pool.Range("E6").Value = saved_b6
        print(f"_phase4h_toggle_smoke C3 raw={raw_val}, rmb={rmb_val}")
        if not (isinstance(raw_val, (int, float)) and isinstance(rmb_val, (int, float)) and rmb_val > raw_val * 5):
            failures.append("E6 realtime toggle did not change USD numeric value")

        print("\n[4] local cache smoke", flush=True)
        excel.Run("模块_工具函数.ClearLocalCache")
        miss = excel.Run("模块_工具函数.ReadLocalHttpCache", "phase4h_inspect_key")
        excel.Run("模块_工具函数.WriteLocalHttpCache", "phase4h_inspect_key", '{"ok":true}')
        hit = excel.Run("模块_工具函数.ReadLocalHttpCache", "phase4h_inspect_key")
        print(f"cache before write={miss!r}, after write={hit!r}")
        if str(miss) != "" or '"ok":true' not in str(hit):
            failures.append("local cache read/write smoke failed")

        print("\n[5] stockanalysis fallback smoke", flush=True)
        ws_pool.Range("E13:F80").ClearContents()
        ws_pool.Range("E13").Value = "BABA"
        ws_pool.Range("F13").Value = "阿里巴巴"
        ws_pool.Range("E3").Value = 2025
        ws_pool.Range("E4").Value = "全部"
        ws_pool.Range("E5").Value = "invalid_cookie_for_phase4h"
        excel.Run("模块_抓美股资产负债表.Main")
        ws_diag = wb.Worksheets("美股_抓取诊断")
        fallback_rows = 0
        for r in range(3, ws_diag.UsedRange.Rows.Count + 1):
            src = cell_text(ws_diag, r, 5)
            if "stockanalysis" in src and "fallback" in src:
                fallback_rows += 1
        print(f"stockanalysis fallback rows={fallback_rows}")
        if fallback_rows == 0:
            failures.append("stockanalysis fallback did not write diagnostic rows")

        excel.Run("模块_工具函数.SetSilentMode", False)
        wb.Close(False)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass

    if failures:
        print("\n*** FAIL ***")
        for item in failures:
            print(f"- {item}")
        return 1

    print("\n*** PASS: Phase 4h workbook state checks passed ***")
    return 0


if __name__ == "__main__":
    sys.exit(main())

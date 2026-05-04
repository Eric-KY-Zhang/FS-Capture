from __future__ import annotations

import sys
from pathlib import Path

import win32com.client as win32

XLSM = Path(r"E:\Claude+CODEX Project\FS Capture\VBA Captor\上市公司财务数据查询.xlsm")


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def shape_text(shape) -> str:
    try:
        return shape.TextFrame2.TextRange.Text.strip()
    except Exception:
        return ""


def clean_addr(addr: str) -> str:
    return str(addr).replace("$", "")


def main() -> None:
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(XLSM))
        try:
            log("=== Phase 4g state inspect ===")

            excel.Run("模块_工具函数.BuildCrossMarketIndicatorSheet")
            ws = wb.Sheets("跨市场_指标表")
            log("\n[1] 跨市场_指标表 headers")
            r1 = [ws.Cells(1, c).Value for c in range(1, 13)]
            r2 = [ws.Cells(2, c).Value for c in range(1, 13)]
            log(f"R1 C1:C12 = {r1}")
            log(f"R2 C1:C12 = {r2}")

            log("\n[2] Row 3-5 formulas and values")
            for row in range(3, 6):
                label = [ws.Cells(row, c).Value for c in range(1, 4)]
                cells = []
                for col in range(4, 9):
                    cell = ws.Cells(row, col)
                    cells.append(
                        {
                            "addr": clean_addr(cell.Address),
                            "value": cell.Value,
                            "formula": cell.Formula,
                        }
                    )
                log(f"R{row} label={label}")
                for item in cells:
                    log(f"  {item['addr']}: value={item['value']!r}, formula={item['formula']!r}")

            pool = wb.Sheets("样本池")
            log("\n[3] hide-tab buttons")
            for name in ("BtnHideA", "BtnHideUS", "BtnHideHK", "BtnHideKR", "BtnHideAll"):
                shape = pool.Shapes.Item(name)
                log(
                    f"{name}: {clean_addr(shape.TopLeftCell.Address)}:"
                    f"{clean_addr(shape.BottomRightCell.Address)}, "
                    f"caption={shape_text(shape)!r}, action={shape.OnAction}"
                )

            log("\n[4] market sheet visibility before toggle")
            market_prefixes = ("A股_", "美股_", "港股_", "韩股_")
            market_sheets = []
            for i in range(1, wb.Sheets.Count + 1):
                sheet = wb.Sheets(i)
                if sheet.Name.startswith(market_prefixes):
                    market_sheets.append(sheet)
                    log(f"{sheet.Name}: Visible={sheet.Visible}")

            log("\n[5] global toggle check")
            pool.Activate()
            excel.Run("模块_总入口.切换所有分市场tabs")
            hidden_count = sum(1 for sheet in market_sheets if sheet.Visible == 0)
            log(f"Hidden after first global toggle: {hidden_count}/{len(market_sheets)}")
            excel.Run("模块_总入口.切换所有分市场tabs")
            visible_count = sum(1 for sheet in market_sheets if sheet.Visible == -1)
            log(f"Visible after second global toggle: {visible_count}/{len(market_sheets)}")

            log("\n[6] diagnostic headers")
            for name in ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"):
                diag = wb.Sheets(name)
                headers = [diag.Cells(2, c).Value for c in range(1, 12)]
                log(f"{name}: headers={headers}, K2={diag.Cells(2, 11).Value!r}")

        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()

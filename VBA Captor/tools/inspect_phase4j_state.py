from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

import win32com.client as win32


ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "上市公司财务数据查询.xlsm"


def rgb_long(hex_str: str) -> int:
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r


def cell_text(ws, row: int, col: int) -> str:
    return str(ws.Cells(row, col).Text or "").strip()


def shape_caption(ws, name: str) -> str:
    return str(ws.Shapes(name).TextFrame2.TextRange.Text or "").strip()


def main() -> int:
    print("=== Phase 4j state inspect ===", flush=True)
    failures: list[str] = []
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None

    try:
        wb = excel.Workbooks.Open(str(BOOK))
        excel.Run("模块_工具函数.BuildAllCrossMarketSheets")

        print("\n[1] 字段映射 sheet", flush=True)
        ws_map = wb.Worksheets("字段映射")
        headers = [cell_text(ws_map, 1, c) for c in range(1, 9)]
        expected_headers = ["报表类型", "标准字段名", "A股原名", "美股原名", "港股原名", "韩股原名", "显示顺序", "备注"]
        print(f"headers={headers}")
        if headers != expected_headers:
            failures.append("字段映射 headers mismatch")
        last_map_row = ws_map.Cells(ws_map.Rows.Count, 2).End(-4162).Row
        counts = Counter(cell_text(ws_map, r, 1) for r in range(2, last_map_row + 1))
        reviewer_flags = sum("[需 reviewer 确认]" in cell_text(ws_map, r, 8) for r in range(2, last_map_row + 1))
        print(f"mapping rows={last_map_row - 1}, counts={dict(counts)}, reviewer_flags={reviewer_flags}")
        if not (60 <= last_map_row - 1 <= 80):
            failures.append("字段映射 row count outside expected range")

        print("\n[2] 跨市场 P3 mapped/unmapped layout", flush=True)
        specs = [
            ("BS", "跨市场_资产负债表"),
            ("IS", "跨市场_利润表"),
            ("CF", "跨市场_现金流量表"),
        ]
        hit_rates: dict[str, tuple[int, int, float]] = {}
        for code, sheet_name in specs:
            ws = wb.Worksheets(sheet_name)
            mapping_names = [cell_text(ws_map, r, 2) for r in range(2, last_map_row + 1) if cell_text(ws_map, r, 1) == code]
            sep_row = 0
            last_row = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
            for r in range(3, last_row + 1):
                if "未映射字段" in cell_text(ws, r, 1):
                    sep_row = r
                    break
            if sep_row == 0:
                failures.append(f"{sheet_name} missing unmapped separator")
                continue
            mapped_rows = [cell_text(ws, r, 2) for r in range(3, sep_row) if cell_text(ws, r, 2)]
            matched = sum(1 for name in mapping_names if name in mapped_rows)
            rate = matched / max(1, len(mapping_names))
            hit_rates[code] = (matched, len(mapping_names), rate)
            print(f"{sheet_name}: mapped_rows={len(mapped_rows)}, mapping_hit={matched}/{len(mapping_names)} ({rate:.1%}), sep_row={sep_row}")
            if len(mapped_rows) < 5:
                failures.append(f"{sheet_name} mapped upper section < 5 rows")
            grouped_rows = 0
            hidden_group_rows = 0
            for r in range(sep_row + 1, min(last_row, sep_row + 120) + 1):
                try:
                    if ws.Rows(r).OutlineLevel > 1:
                        grouped_rows += 1
                        if ws.Rows(r).Hidden:
                            hidden_group_rows += 1
                except Exception:
                    pass
            print(f"  grouped_rows={grouped_rows}, hidden_group_rows={hidden_group_rows}")
            if grouped_rows == 0:
                failures.append(f"{sheet_name} outline grouping not detected")

        print("\n[3] visual system samples", flush=True)
        ws_pool = wb.Worksheets("样本池")
        checks = {
            "样本池 Q7 fill": (int(ws_pool.Range("Q7").Interior.Color), rgb_long("ED7D31")),
            "样本池 Q8 caption": (shape_caption(ws_pool, "BtnBuildCrossAll"), "一键跨市场对比"),
            "样本池 T9 caption": (shape_caption(ws_pool, "BtnClearCache"), "清空 HTTP 缓存"),
            "跨市场 tab color": (int(wb.Worksheets("跨市场_资产负债表").Tab.Color), rgb_long("ED7D31")),
            "使用说明 A1 font": (int(wb.Worksheets("使用说明").Range("A1").Font.Size), 24),
            "汇率 J10 title": (cell_text(wb.Worksheets("汇率"), 10, 10), "汇率数据来源与折算说明"),
        }
        for name, (got, want) in checks.items():
            print(f"{name}: got={got!r}, want={want!r}")
            if got != want:
                failures.append(f"{name} mismatch")

        print("\n[4] 使用说明 7 sections", flush=True)
        ws_intro = wb.Worksheets("使用说明")
        used = ws_intro.UsedRange
        parts = []
        for r in range(1, used.Rows.Count + 1):
            for c in range(1, used.Columns.Count + 1):
                parts.append(str(used.Cells(r, c).Value or ""))
        intro_text = "\n".join(parts)
        for title in ("§ 1 项目概览", "§ 2 快速开始", "§ 3 输出 sheet 说明", "§ 4 数据源声明", "§ 5 汇率换算说明", "§ 6 常见问题", "§ 7 版本历史"):
            if title not in intro_text:
                failures.append(f"missing intro section {title}")
        print(f"intro_chars={len(intro_text)}")
        if len(intro_text) < 1200:
            failures.append("使用说明内容过短")

        print("\n[5] tab visibility behavior", flush=True)
        us_official = ["美股_资产负债表", "美股_利润表", "美股_现金流量表", "美股_指标表"]
        for name in us_official:
            wb.Worksheets(name).Visible = 0
        wb.Worksheets("美股_抓取诊断").Visible = 0
        excel.Run("模块_总入口.切换美股tabs")
        official_visible = [wb.Worksheets(name).Visible for name in us_official]
        diag_visible = wb.Worksheets("美股_抓取诊断").Visible
        print(f"after toggle US official={official_visible}, diag={diag_visible}")
        if any(v != -1 for v in official_visible) or diag_visible != 0:
            failures.append("market toggle did not exclude diagnostic sheet")
        for name in us_official:
            wb.Worksheets(name).Visible = 0
        excel.Run("模块_总入口.UnhideMarketTabs", "US")
        official_visible = [wb.Worksheets(name).Visible for name in us_official]
        diag_visible = wb.Worksheets("美股_抓取诊断").Visible
        print(f"after UnhideMarketTabs US official={official_visible}, diag={diag_visible}")
        if any(v != -1 for v in official_visible) or diag_visible != 0:
            failures.append("UnhideMarketTabs did not preserve diagnostic hidden")

        for prefix in ("A股_", "美股_", "港股_", "韩股_", "跨市场_"):
            for idx in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(idx)
                if ws.Name.startswith(prefix) and "抓取诊断" not in ws.Name:
                    ws.Visible = -1
        for name in ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"):
            wb.Worksheets(name).Visible = 0

        print("\n[6] mapping hit rates", flush=True)
        for code, (matched, total, rate) in hit_rates.items():
            print(f"{code}: {matched}/{total} = {rate:.1%}")
            if rate < 0.5:
                failures.append(f"{code} mapping hit rate below 50%")

        wb.Close(SaveChanges=False)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()

    if failures:
        print("\n*** FAIL ***")
        for item in failures:
            print(f"- {item}")
        return 1

    print("\n*** PASS: Phase 4j workbook state checks passed ***")
    return 0


if __name__ == "__main__":
    sys.exit(main())

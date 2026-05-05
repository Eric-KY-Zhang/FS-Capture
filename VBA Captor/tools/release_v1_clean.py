"""
v1.0 release 清理(无 macro 调用,直接 cell-level 清空,避开 CleanReleaseWorkbook macro hang)
"""
import sys
from pathlib import Path
import zipfile
import win32com.client as win32

ROOT = Path(r"E:\Claude+CODEX Project\FS Capture\VBA Captor")
RELEASE_FILE = ROOT / "release" / "上市公司财务数据查询v1.0_release.xlsm"


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def main() -> None:
    if not RELEASE_FILE.exists():
        raise SystemExit(f"Release file missing: {RELEASE_FILE}")

    log("Opening release xlsm")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(RELEASE_FILE))
        try:
            # 1. 清 cookie / fallback 相关 cell (Phase 4i 后 cookie 在 E5; 老版 B5 兼容)
            try:
                pool = wb.Sheets("样本池")
                for addr in ("B5", "B7", "B8", "O5", "O6"):
                    try:
                        pool.Range(addr).Value = ""
                    except Exception:
                        pass
                # E5 是合并 (E5:M5), 用 unmerge + 写空
                try:
                    rng = pool.Range("E5:M5")
                    is_merged = bool(rng.MergeCells)
                    if is_merged:
                        rng.UnMerge()
                    pool.Range("E5").Value = ""
                    if is_merged:
                        rng.Merge()
                except Exception as e:
                    log(f"  ! E5 clear: {e}")
                log("  + cookie / fallback cells cleared")
            except Exception as e:
                log(f"  ! pool cells: {e}")

            # 2. 清 3 张诊断 sheet 历史 R3+
            for diag_name in ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"):
                try:
                    ws = wb.Sheets(diag_name)
                    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                    if last_row >= 3:
                        ws.Range(ws.Cells(3, 1), ws.Cells(last_row, 17)).Clear()
                    log(f"  + {diag_name} R3:R{last_row} cleared (was {last_row - 2} rows)")
                except Exception as e:
                    log(f"  ! {diag_name}: {e}")

            # 3. 清 16 张分市场报表数据 + 跨市场指标表(让 release 是干净空模板)
            sheet_names = []
            for market in ("A股", "美股", "港股", "韩股"):
                for kind in ("资产负债表", "利润表", "现金流量表", "指标表"):
                    sheet_names.append(f"{market}_{kind}")
            sheet_names.append("跨市场_指标表")
            for name in sheet_names:
                try:
                    ws = wb.Sheets(name)
                    used = ws.UsedRange
                    if used.Rows.Count > 2:
                        # R3 起的数据区清空; R1-R2 静态表头保留
                        try:
                            ws.Range(ws.Cells(3, 1), ws.Cells(used.Rows.Count, used.Columns.Count)).Clear()
                        except Exception:
                            pass
                except Exception as e:
                    log(f"  ! clear {name}: {e}")
            log(f"  + {len(sheet_names)} 报表 sheet R3+ data cleared (header preserved)")

            # 4. BuiltinDocumentProperties
            try:
                props = wb.BuiltinDocumentProperties
                props.Item("Author").Value = "Eric Zhang"
                try: props.Item("Last Author").Value = "Eric Zhang"
                except Exception: pass
                try: props.Item("Company").Value = ""
                except Exception: pass
                try: props.Item("Manager").Value = ""
                except Exception: pass
                try: props.Item("Comments").Value = "v1.0 release - 上市公司财务数据查询 - 联系: 214978902@qq.com"
                except Exception: pass
                try: props.Item("Title").Value = "上市公司财务数据查询 v1.0"
                except Exception: pass
                log("  + BuiltinDocumentProperties: Author=Eric Zhang, Title=上市公司财务数据查询 v1.0")
            except Exception as e:
                log(f"  ! props: {e}")

            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    log(f"  + saved (size={RELEASE_FILE.stat().st_size:,} bytes)")

    # 5. Strip xl/webextensions/ via zipfile
    log("Strip xl/webextensions/ from release zip")
    tmp = RELEASE_FILE.with_suffix(".tmp")
    removed = 0
    try:
        with zipfile.ZipFile(RELEASE_FILE, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if "webextension" in item.filename.lower():
                        removed += 1
                        continue
                    zout.writestr(item, zin.read(item.filename))
        tmp.replace(RELEASE_FILE)
        log(f"  + removed {removed} webextension entries")
    except Exception as e:
        if tmp.exists():
            tmp.unlink()
        log(f"  ! strip webextensions: {e}")

    log(f"\nFinal size: {RELEASE_FILE.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()

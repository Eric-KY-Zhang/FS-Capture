"""
v1.0 发布版构建脚本(一次性,Phase 4n 收尾用):
  1. 把当前 xlsm 复制成 release/上市公司财务数据查询v1.0_source.xlsm(开发版)
  2. 把当前 xlsm 复制成 release/上市公司财务数据查询v1.0_release.xlsm
  3. 在 release 副本上跑 CleanReleaseWorkbook 宏(清 cookie / 诊断历史 / cache)
  4. 改 BuiltinDocumentProperties (Author=Eric Zhang, 删除其他个人化属性)
  5. 移除 xl/webextensions/ (claude.fileId 残留)
"""
from __future__ import annotations
import shutil
import sys
import zipfile
from pathlib import Path
import win32com.client as win32

ROOT = Path(r"E:\Claude+CODEX Project\FS Capture\VBA Captor")
SRC_XLSM = ROOT / "上市公司财务数据查询.xlsm"
RELEASE_DIR = ROOT / "release"
SOURCE_FILE = RELEASE_DIR / "上市公司财务数据查询v1.0_source.xlsm"
RELEASE_FILE = RELEASE_DIR / "上市公司财务数据查询v1.0_release.xlsm"


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def step1_copy_source() -> None:
    log("[1] Copy current xlsm -> v1.0_source.xlsm (preserve debug)")
    shutil.copy2(SRC_XLSM, SOURCE_FILE)
    log(f"    + {SOURCE_FILE.name} ({SOURCE_FILE.stat().st_size:,} bytes)")


def step2_copy_release() -> None:
    log("[2] Copy current xlsm -> v1.0_release.xlsm (will be cleaned)")
    shutil.copy2(SRC_XLSM, RELEASE_FILE)
    log(f"    + {RELEASE_FILE.name} ({RELEASE_FILE.stat().st_size:,} bytes)")


def step3_clean_via_macro() -> None:
    log("[3] Run CleanReleaseWorkbook macro on release file")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(RELEASE_FILE))
        try:
            try:
                excel.Run("模块_工具函数.CleanReleaseWorkbook")
                log("    + macro completed")
            except Exception as e:
                log(f"    ! macro raise: {e}")

            # set document properties
            try:
                props = wb.BuiltinDocumentProperties
                props.Item("Author").Value = "Eric Zhang"
                props.Item("Last Author").Value = "Eric Zhang"
                try:
                    props.Item("Company").Value = ""
                except Exception:
                    pass
                try:
                    props.Item("Manager").Value = ""
                except Exception:
                    pass
                try:
                    props.Item("Comments").Value = "v1.0 release - 上市公司财务数据查询 - 联系: 214978902@qq.com"
                except Exception:
                    pass
                try:
                    props.Item("Title").Value = "上市公司财务数据查询 v1.0"
                except Exception:
                    pass
                log("    + BuiltinDocumentProperties updated (Author=Eric Zhang)")
            except Exception as e:
                log(f"    ! props update: {e}")

            # also ensure samples sheet B5/E5/B8/O5/O6 are cleared (CleanReleaseWorkbook should do this, defensive double-check)
            try:
                pool = wb.Sheets("样本池")
                for addr in ("B5", "E5", "B8", "O5", "O6"):
                    try:
                        pool.Range(addr).Value = ""
                    except Exception:
                        pass
                log("    + cookie / fallback cells double-cleared")
            except Exception as e:
                log(f"    ! cookie cell clear: {e}")

            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    log(f"    + saved (size={RELEASE_FILE.stat().st_size:,} bytes)")


def step4_strip_webextensions() -> None:
    log("[4] Strip xl/webextensions/ from release zip (claude.fileId etc.)")
    tmp = RELEASE_FILE.with_suffix(".tmp")
    removed = 0
    try:
        with zipfile.ZipFile(RELEASE_FILE, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename.startswith("xl/webextensions/"):
                        removed += 1
                        continue
                    if "webExtensionLink" in item.filename:
                        removed += 1
                        continue
                    if item.filename == "xl/webExtensions/webExtension1.xml":
                        removed += 1
                        continue
                    zout.writestr(item, zin.read(item.filename))
        tmp.replace(RELEASE_FILE)
        log(f"    + removed {removed} webextension-related entries")
    except Exception as e:
        if tmp.exists():
            tmp.unlink()
        log(f"    ! strip webextensions: {e}")


def step5_summary() -> None:
    log("\n[Summary]")
    for p in (SOURCE_FILE, RELEASE_FILE):
        if p.exists():
            log(f"  + {p.name}: {p.stat().st_size:,} bytes")
        else:
            log(f"  ! MISSING: {p.name}")


def main() -> None:
    if not SRC_XLSM.exists():
        raise SystemExit(f"Source xlsm missing: {SRC_XLSM}")
    if not RELEASE_DIR.exists():
        RELEASE_DIR.mkdir(parents=True, exist_ok=True)
    step1_copy_source()
    step2_copy_release()
    step3_clean_via_macro()
    step4_strip_webextensions()
    step5_summary()


if __name__ == "__main__":
    main()

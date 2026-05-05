from __future__ import annotations

from pathlib import Path
import re

import win32com.client as win32


ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "上市公司财务数据查询.xlsm"
FIXTURE_DIR = ROOT / "tests" / "fixtures"
MODULE_TEST = ROOT / "modules" / "模块_测试.bas"
MODULE_UTIL = ROOT / "modules" / "模块_工具函数.bas"


EXPECTED_FIXTURES = {
    "sec_aapl_companyfacts.json",
    "xueqiu_hk_00700_balance.json",
    "stockanalysis_kr_005930_income.html",
    "fx_usdcny_kline.json",
    "http_429_response.json",
    "malformed_xueqiu.txt",
    "missing_fields_edgar.json",
}

EXPECTED_MACROS = {
    "Test_Offline_US_Edgar_AAPL",
    "Test_Offline_HK_Xueqiu_Tencent",
    "Test_Offline_KR_StockAnalysis_Samsung",
    "Test_Offline_FX_Missing_DoesNotFallbackToOne",
    "Test_Offline_Diagnostic_Score_NotDate",
    "Test_Offline_Cache_HitMissExpired",
    "Test_Offline_AppState_RestoreAfterError",
    "Test_Offline_DataQuality_BS_Imbalance_Detection",
}


def log(message: str = "") -> None:
    print(message, flush=True)


def cell_text(ws, row: int, col: int) -> str:
    return str(ws.Cells(row, col).Text)


def main() -> int:
    log("=== Phase 4m state inspect ===")
    failures: list[str] = []

    log("\n[1] offline fixture files")
    present = {p.name for p in FIXTURE_DIR.iterdir()} if FIXTURE_DIR.exists() else set()
    missing = sorted(EXPECTED_FIXTURES - present)
    log(f"present fixtures={sorted(EXPECTED_FIXTURES & present)}")
    if missing:
        failures.append("missing offline fixtures: " + ", ".join(missing))

    log("\n[2] offline test macro inventory")
    test_text = MODULE_TEST.read_text(encoding="utf-8")
    macros = sorted(set(re.findall(r"Public Sub (Test_Offline_[A-Za-z0-9_]+)\(", test_text)))
    log(f"offline macros={macros}")
    if set(macros) != EXPECTED_MACROS:
        failures.append("offline macro set does not match the Phase 4m plan")

    util_text = MODULE_UTIL.read_text(encoding="utf-8")
    for needle in ("RunDataQualityChecks", "AddDiagnosticQARow", "GetTtlHoursForSource"):
        if needle not in util_text:
            failures.append(f"missing utility implementation: {needle}")
    if not re.search(r"Sub BuildCrossMarketIndicatorSheet\(\)[\s\S]*?RunDataQualityChecks[\s\S]*?End Sub", util_text):
        failures.append("BuildCrossMarketIndicatorSheet is not wired to RunDataQualityChecks")

    ttl_call_files = [
        ROOT / "modules" / "模块_工具函数.bas",
        ROOT / "modules" / "模块_抓美股财报.bas",
        ROOT / "modules" / "模块_抓韩股财报.bas",
    ]
    for path in ttl_call_files:
        for line_no, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if "RunCachedHttpGet(" in line and "Public Function RunCachedHttpGet" not in line:
                log(f"TTL call {path.name}:{line_no}: {line.strip()}")
                if "GetTtlHoursForSource(" not in line:
                    failures.append(f"RunCachedHttpGet caller still has non-source-aware TTL: {path.name}:{line_no}")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(BOOK))
        excel.Run("模块_工具函数.SetSilentMode", True)

        log("\n[3] QA diagnostic rows")
        excel.Run("模块_工具函数.RunDataQualityChecks")
        ws_diag = wb.Worksheets("美股_抓取诊断")
        last_row = ws_diag.Cells(ws_diag.Rows.Count, 1).End(-4162).Row  # xlUp
        qa_rows: list[tuple[int, str, str, str]] = []
        for row in range(3, last_row + 1):
            if cell_text(ws_diag, row, 1) == "GLOBAL_QA":
                qa_rows.append((row, cell_text(ws_diag, row, 2), cell_text(ws_diag, row, 4), cell_text(ws_diag, row, 10)))
        log(f"GLOBAL_QA rows={qa_rows}")
        qa_codes = {code for _, code, _, _ in qa_rows}
        if {"BS_BALANCE", "FX_MISSING", "KEY_FIELDS"} - qa_codes:
            failures.append("GLOBAL_QA rows do not include BS_BALANCE / FX_MISSING / KEY_FIELDS")

        log("\n[4] per-source TTL map")
        ttl_sec = int(excel.Run("模块_工具函数.GetTtlHoursForSource", "SEC_TICKER_MAP"))
        ttl_edgar = int(excel.Run("模块_工具函数.GetTtlHoursForSource", "EDGAR"))
        ttl_xueqiu = int(excel.Run("模块_工具函数.GetTtlHoursForSource", "XUEQIU"))
        ttl_unknown = int(excel.Run("模块_工具函数.GetTtlHoursForSource", "unknown"))
        log(f"SEC_TICKER_MAP={ttl_sec}, EDGAR={ttl_edgar}, XUEQIU={ttl_xueqiu}, unknown={ttl_unknown}")
        if ttl_sec != 168 or ttl_edgar != 24 or ttl_xueqiu != 12 or ttl_unknown != 24:
            failures.append("GetTtlHoursForSource returned unexpected TTL values")

        excel.Run("模块_工具函数.SetSilentMode", False)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()

    if failures:
        log("\n*** FAIL: Phase 4m state checks failed ***")
        for item in failures:
            log(f"  - {item}")
        return 1

    log("\n*** PASS: Phase 4m workbook state checks passed ***")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

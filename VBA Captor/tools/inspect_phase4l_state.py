from __future__ import annotations

from pathlib import Path

import win32com.client as win32


ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "上市公司财务数据查询.xlsm"


EXPECTED_DIAG_HEADERS = [
    "公司", "报表", "输出指标", "状态", "数据源",
    "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注", "FX_Rate",
    "CacheStatus", "CacheAgeHours", "HTTPStatus", "ElapsedMs", "RetryCount", "ErrorStage",
]


def log(msg: str = "") -> None:
    print(msg, flush=True)


def cell_text(ws, addr: str) -> str:
    return str(ws.Range(addr).Text)


def main() -> int:
    log("=== Phase 4l state inspect ===")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    failures: list[str] = []

    try:
        wb = excel.Workbooks.Open(str(BOOK))
        excel.Run("模块_工具函数.SetSilentMode", True)

        log("\n[1] diagnostic telemetry headers")
        for sheet_name in ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"):
            ws = wb.Worksheets(sheet_name)
            headers = [str(ws.Cells(2, c).Value) for c in range(1, 18)]
            fmt_lq = str(ws.Range("L3:Q3").NumberFormat)
            log(f"{sheet_name}: headers={headers}")
            log(f"{sheet_name}: L3:Q3 NumberFormat={fmt_lq!r}")
            if headers != EXPECTED_DIAG_HEADERS:
                failures.append(f"{sheet_name} diagnostic headers are not 17-column Phase 4l headers")
            if fmt_lq != "@":
                failures.append(f"{sheet_name} telemetry columns L:Q are not text formatted")

        log("\n[2] HTTP/cache telemetry MISS -> HIT")
        excel.Run("模块_测试.TestPhase4lHttpMissHitSmoke")
        ws_http = wb.Worksheets("_phase4l_http_smoke")
        first_cache = cell_text(ws_http, "A1")
        first_status = int(ws_http.Range("B1").Value or 0)
        first_elapsed = int(ws_http.Range("C1").Value or 0)
        second_cache = cell_text(ws_http, "D1")
        second_status = int(ws_http.Range("E1").Value or 0)
        second_elapsed = int(ws_http.Range("F1").Value or 0)
        first_len = int(ws_http.Range("G1").Value or 0)
        second_len = int(ws_http.Range("H1").Value or 0)
        log(
            f"AAPL SEC first={first_cache}/{first_status}/{first_elapsed}ms/{first_len} bytes; "
            f"second={second_cache}/{second_status}/{second_elapsed}ms/{second_len} bytes"
        )
        if first_cache != "MISS" or first_status != 200 or first_len <= 0:
            failures.append("AAPL SEC first request was not a successful cache MISS")
        if second_cache != "HIT" or second_status != 0 or second_len <= 0:
            failures.append("AAPL SEC second request was not a cache HIT")

        ws_diag = wb.Worksheets("美股_抓取诊断")
        diag_tail = [str(ws_diag.Cells(3, c).Text) for c in range(12, 18)]
        log(f"diagnostic row L:Q={diag_tail}")
        if diag_tail[0] != "HIT" or diag_tail[2] != "0":
            failures.append("diagnostic telemetry row did not capture cache HIT / HTTPStatus 0")

        log("\n[3] SEC rate limit")
        excel.Run("模块_测试.TestPhase4lSecRateSmoke")
        ws_sec = wb.Worksheets("_phase4l_sec_smoke")
        sec_interval = float(ws_sec.Range("A1").Value or 0)
        sec_statuses = (int(ws_sec.Range("B1").Value or 0), int(ws_sec.Range("C1").Value or 0))
        sec_caches = (cell_text(ws_sec, "D1"), cell_text(ws_sec, "E1"))
        log(f"SEC interval={sec_interval:.1f}ms, statuses={sec_statuses}, cache={sec_caches}")
        if sec_interval < 110:
            failures.append(f"SEC request interval below 110ms: {sec_interval:.1f}ms")
        if sec_statuses != (200, 200) or sec_caches != ("MISS", "MISS"):
            failures.append("SEC rate smoke did not perform two successful MISS requests")

        log("\n[4] CleanReleaseWorkbook")
        excel.Run("模块_测试.TestPhase4lCleanReleaseSmoke")
        ws_release = wb.Worksheets("_phase4l_release_smoke")
        e5_after = cell_text(ws_release, "A1")
        b5_after = cell_text(ws_release, "B1")
        cache_after = cell_text(ws_release, "C1")
        diag_after = cell_text(ws_release, "D1")
        cache_dir = ROOT / ".cache"
        cache_files = list(cache_dir.glob("*")) if cache_dir.exists() else []
        log(f"E5={e5_after!r}, B5={b5_after!r}, cache={cache_after!r}, diag_A3={diag_after!r}, cache_files={len(cache_files)}")
        if e5_after or b5_after:
            failures.append("CleanReleaseWorkbook did not clear E5/B5 cookie cells")
        if cache_after or cache_files:
            failures.append("CleanReleaseWorkbook did not clear local HTTP cache")
        if diag_after:
            failures.append("CleanReleaseWorkbook did not clear diagnostic history")

        excel.Run("模块_工具函数.SetSilentMode", False)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()

    if failures:
        log("\n*** FAIL: Phase 4l state checks failed ***")
        for item in failures:
            log(f"  - {item}")
        return 1

    log("\n*** PASS: Phase 4l workbook state checks passed ***")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

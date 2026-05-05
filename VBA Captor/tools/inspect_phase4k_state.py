from __future__ import annotations

from pathlib import Path

import win32com.client as win32


ROOT = Path(__file__).resolve().parents[1]


def log(msg: str = "") -> None:
    print(msg, flush=True)


def norm_addr(shape) -> str:
    top_left = str(shape.TopLeftCell.Address).replace("$", "")
    bottom_right = str(shape.BottomRightCell.Address).replace("$", "")
    return f"{top_left}:{bottom_right}"


def main() -> int:
    books = list(ROOT.glob("*.xlsm"))
    if not books:
        raise SystemExit("no xlsm workbook found")
    book = books[0]

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    failures: list[str] = []
    try:
        wb = excel.Workbooks.Open(str(book))

        log("=== Phase 4k state inspect ===")

        log("\n[1] diagnostic Score column text format")
        for sheet_name in ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"):
            ws = wb.Worksheets(sheet_name)
            fmt = str(ws.Range("I3").NumberFormat)
            log(f"{sheet_name} I3 NumberFormat={fmt!r}")
            if fmt != "@":
                failures.append(f"{sheet_name} Score column is not text formatted")

        excel.Run("模块_测试.TestPhase4kScoreSmoke")
        ws_kr_diag = wb.Worksheets("韩股_抓取诊断")
        score_text = str(ws_kr_diag.Range("I3").Text)
        log(f"score smoke I3 Text={score_text!r}, Value={ws_kr_diag.Range('I3').Value!r}")
        if score_text != "1/1":
            failures.append("Score text smoke did not preserve 1/1")

        log("\n[2] AppStateGuard module")
        try:
            comp = wb.VBProject.VBComponents("模块_AppStateGuard")
            code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            has_begin = "BeginAppState" in code
            has_end = "EndAppState" in code
            log(f"模块_AppStateGuard found: Begin={has_begin}, End={has_end}")
            if not (has_begin and has_end):
                failures.append("模块_AppStateGuard missing BeginAppState/EndAppState")
        except Exception as exc:
            failures.append(f"模块_AppStateGuard missing: {exc}")

        log("\n[3] FX_MISSING corrupt smoke")
        excel.Run("模块_测试.TestPhase4kFxMissingSmoke")
        ws_missing = wb.Worksheets("_phase4k_fx_missing_smoke")
        blank_status = str(ws_missing.Range("ZZ1").Value)
        missing_count = int(ws_missing.Range("ZZ2").Value or 0)
        log(f"fx missing output={blank_status!r}, diagnostic_count={missing_count}")
        if blank_status != "BLANK":
            failures.append("KRW missing FX smoke output cell was not blank")
        if missing_count < 1:
            failures.append("KRW missing FX smoke did not write FX_MISSING diagnostic row")

        first_missing_row = None
        last_row = ws_kr_diag.Cells(ws_kr_diag.Rows.Count, 1).End(-4162).Row
        for row in range(3, last_row + 1):
            if str(ws_kr_diag.Cells(row, 4).Value) == "FX_MISSING":
                first_missing_row = row
                break
        log(f"first FX_MISSING diagnostic row={first_missing_row}")

        log("\n[4] live FX UDF formula + response time")
        excel.Run("模块_测试.TestPhase4kLiveFxSmoke")
        ws_live = wb.Worksheets("_phase4k_live_fx_smoke")
        formula = str(ws_live.Range("C3").Formula)
        before_val = ws_live.Range("ZZ1").Value
        after_val = ws_live.Range("ZZ2").Value
        elapsed = float(ws_live.Range("ZZ3").Value or 0)
        log(f"C3 Formula={formula}")
        log(f"before={before_val}, after={after_val}, elapsed={elapsed:.3f}s")
        if "GetFxFromSheet" not in formula:
            failures.append("live FX formula does not call GetFxFromSheet")
        if not (isinstance(before_val, (int, float)) and isinstance(after_val, (int, float)) and after_val != before_val):
            failures.append("live FX value did not change after rate edit")
        if elapsed > 5:
            failures.append(f"live FX response exceeded 5s: {elapsed:.3f}s")

    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()

    if failures:
        log("\n*** FAIL: Phase 4k state checks failed ***")
        for item in failures:
            log(f"  - {item}")
        return 1

    log("\n*** PASS: Phase 4k workbook state checks passed ***")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

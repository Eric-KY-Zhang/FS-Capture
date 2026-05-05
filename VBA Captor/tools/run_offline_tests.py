from __future__ import annotations

import sys
import time
import json
from pathlib import Path

import win32com.client as win32


ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "上市公司财务数据查询.xlsm"
LOCK = ROOT / f"~${BOOK.name}"
FIXTURE_DIR = ROOT / "tests" / "fixtures"
TRACE = ROOT / "test_outputs" / "phase4m_offline_trace.txt"


MOCK_FIXTURES = {
    "xueqiu_hk_00700_balance.json": (
        '{"data":{"quote_name":"Tencent","currency":"CNY","currency_name":"RMB",'
        '"list":[{"report_date":"2024-12-31","ta":1000,"tlia":400,"teqy":600}]}}\n'
    ),
    "stockanalysis_kr_005930_income.html": (
        "<html><head><title>Samsung Electronics Co. Ltd. Income Statement</title></head>"
        "<body><script>var financialData={\"revenue\":[1000,900],\"opinc\":[200,180],"
        "\"netinc\":[150,120]};</script>005930</body></html>\n"
    ),
    "fx_usdcny_kline.json": (
        '{"data":{"symbol":"USDCNY.FX","item":[{"timestamp":1735603200000,'
        '"close":7.2,"volume":0}]}}\n'
    ),
    "http_429_response.json": '{\n  "status": 429,\n  "statusText": "Too Many Requests",\n  "retryAfter": 1\n}\n',
    "malformed_xueqiu.txt": '{"data": {"list": [ {"report_date": "2024-12-31", "ta": 1000 }',
    "missing_fields_edgar.json": (
        '{\n'
        '  "cik": 320193,\n'
        '  "entityName": "Offline Missing Fields Inc.",\n'
        '  "facts": {\n'
        '    "us-gaap": {\n'
        '      "Assets": {"units": {"USD": [{"fy": 2024, "fp": "FY", "val": 1000}]}}\n'
        '    }\n'
        '  }\n'
        '}\n'
    ),
}

TEST_MACROS = [
    "Test_Offline_US_Edgar_AAPL",
    "Test_Offline_HK_Xueqiu_Tencent",
    "Test_Offline_KR_StockAnalysis_Samsung",
    "Test_Offline_FX_Missing_DoesNotFallbackToOne",
    "Test_Offline_Diagnostic_Score_NotDate",
    "Test_Offline_Cache_HitMissExpired",
    "Test_Offline_AppState_RestoreAfterError",
    "Test_Offline_DataQuality_BS_Imbalance_Detection",
]


def log(message: str = "") -> None:
    print(message, flush=True)
    TRACE.parent.mkdir(parents=True, exist_ok=True)
    with TRACE.open("a", encoding="utf-8") as fh:
        fh.write(message + "\n")


def ensure_fixtures() -> list[Path]:
    FIXTURE_DIR.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []

    sec_source = ROOT / "samples" / "AAPL_edgar.json"
    if not sec_source.exists():
        raise FileNotFoundError(f"missing source sample: {sec_source}")
    sec_target = FIXTURE_DIR / "sec_aapl_companyfacts.json"
    sec_data = json.loads(sec_source.read_text(encoding="utf-8"))
    us_gaap = sec_data.get("facts", {}).get("us-gaap", {})
    compact = {
        "cik": sec_data.get("cik", 320193),
        "entityName": sec_data.get("entityName", "Apple Inc."),
        "facts": {
            "us-gaap": {
                name: us_gaap[name]
                for name in ("Revenues", "Assets")
                if name in us_gaap
            }
        },
    }
    sec_target.write_text(json.dumps(compact, ensure_ascii=False), encoding="utf-8")
    written.append(sec_target)

    for name, body in MOCK_FIXTURES.items():
        target = FIXTURE_DIR / name
        target.write_text(body, encoding="utf-8")
        written.append(target)

    return written


def main() -> int:
    TRACE.parent.mkdir(parents=True, exist_ok=True)
    TRACE.write_text(f"Phase 4m offline trace start {time.strftime('%Y-%m-%d %H:%M:%S')}\n", encoding="utf-8")
    only = set()
    if "--only" in sys.argv:
        idx = sys.argv.index("--only")
        if idx + 1 < len(sys.argv):
            only = {name.strip() for name in sys.argv[idx + 1].split(",") if name.strip()}

    log("=== Phase 4m offline fixture tests ===")
    if LOCK.exists():
        log(f"FATAL: workbook lock file exists: {LOCK}")
        log("Close the workbook in Excel and rerun.")
        return 2

    fixtures = ensure_fixtures()
    log(f"fixtures ready: {len(fixtures)} files")
    for path in sorted(fixtures):
        log(f"  - {path.relative_to(ROOT)}")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    passed = 0
    failed: list[str] = []

    try:
        wb = excel.Workbooks.Open(str(BOOK))
        excel.Run("模块_工具函数.SetSilentMode", True)
        for macro in TEST_MACROS:
            if only and macro not in only:
                continue
            try:
                log(f"  > {macro}: START")
                result = str(excel.Run("模块_测试.RunOfflineTest", macro))
                if result == "PASS":
                    log(f"  + {macro}: PASS")
                    passed += 1
                else:
                    log(f"  ! {macro}: {result}")
                    failed.append(macro)
            except Exception as exc:  # COM exceptions need to keep moving for summary.
                log(f"  ! {macro}: FAIL - {exc}")
                failed.append(macro)
        excel.Run("模块_工具函数.SetSilentMode", False)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()

    log(f"\nSUMMARY: {passed}/{len(TEST_MACROS)} PASS")
    if failed:
        log("FAILED: " + ", ".join(failed))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# Phase 4m: 优化 Sprint 3 — 离线测试 fixture + 数据质量 QA + cache 分源 TTL

> **版本**: v2(2026-05-05,Phase 4m 优化 Sprint 3 闭环)
> **状态**: ✅ Phase 4m 全期闭环
> **作者**: Claude(planner) + Codex(generator)
> **背景**: Phase 4k 修核心数据 bug + UX live FX,Phase 4l 给 HTTP/cache 层加可观测性 + 重试稳定 + 发布清理。Phase 4m 继续优化 backlog 第三批 — 给项目加**离线 unit 测试**(让重构时有信心)+ 数据质量自动 QA(BS 平衡 / FX missing / 关键字段缺失)+ cache 分源 TTL(避免 24h 一刀切)。

## 项目语境(给 Generator 的 anchor 段)

本期 3 项工作**全部在现有数据流内**,不引入新数据源 / 新抓数频次 / 新并发。Step 1 (P2-03) 是**纯本地工作**(读 samples/ 已有 fixture 跑解析 unit test,不发任何 HTTP),Step 2 (P1-06) 在跨市场指标表生成时跑 3 条 QA 规则写诊断,Step 3 (P1-02) 把现有 cache TTL 从单一 24h 改成 source-aware map。

## 3 项任务清单

| # | Task | 价值 | 工作量 |
|---|---|---|---|
| 1 | **P2-03** 离线 fixture 测试集 | 重构信心 + 字段 mapping bug 早发现 | 4-5h |
| 2 | **P1-06** 数据质量 QA(精简 3 条)| 用户拿到错数据前自动告警 | 1.5h |
| 3 | **P1-02** cache 分源 TTL | 抓数效率(SEC ticker map 168h vs 雪球 12h)| 1.5h |

总耗时 ~7-8h Codex,~1h Reviewer。

## Step 总览

| Step | 内容 | 估时 | 阻塞依赖 |
|---|---|---|---|
| 1 | P2-03:8 个离线测试 fixture + Test 宏 + Python runner | 4-5h | 无 |
| 2 | P1-06:3 条 QA 规则 + 写诊断 + 跨市场指标表生成时自动跑 | 1.5h | 无 |
| 3 | P1-02:cache TTL map per source(`RunCachedHttpGet` 接受 ttl 参数,callers 按数据源传)| 1.5h | 无 |
| 4 | 端到端回归 + 新增 `inspect_phase4m_state.py` + STATUS §DD 收口 | 0.5h | 1-3 |

**Codex 工作流建议**:**1 round 端到端**(跟 4j/4k/4l 一致),commit `Phase 4m optimization sprint 3 (offline fixtures + data quality QA + per-source cache TTL)`。

---

## Step 1 — P2-03 离线 fixture 测试集

### 背景

当前 6 张 frozen 回归驱动:
- `test_fx_live.py`(联网,真打 FX 抓数)
- `diff_phase4f_step3_lite.py`(联网,抓真 A 股)
- `inspect_phase4g/4h/4k/4l_state.py`(state inspection,部分要联网)

**全部是 happy path + integration 级**,缺**unit 级覆盖**:
- 解析逻辑(EDGAR JSON / 雪球 finance / stockanalysis HTML)能否正确提取字段?
- cache HIT/MISS/EXPIRED 边界对吗?
- retry/backoff 在 429 / 5xx 时真的退避?
- malformed JSON / missing fields 时不崩溃?

→ 加 8 个**离线测试**(读本地 fixture,不发 HTTP),让 Codex 重构时跑这些就知道有没有破坏字段映射。

### 实现

**子任务 1A — Fixture 目录结构**:

```
VBA Captor/
└── tests/
    └── fixtures/
        ├── sec_aapl_companyfacts.json       # SEC EDGAR AAPL real response (~150KB)
        ├── xueqiu_hk_00700_balance.json     # 雪球 HK 腾讯 BS (Phase 4h 已有 sample)
        ├── stockanalysis_kr_005930_income.html  # KRX 三星 IS (Phase 4d 已有 sample)
        ├── fx_usdcny_kline.json             # 雪球 USDCNY K 线 (Phase 4f 已有 sample)
        ├── http_429_response.json           # mock 429 response
        ├── malformed_xueqiu.txt             # 故意不完整的 JSON
        ├── missing_fields_edgar.json        # EDGAR JSON 但缺 us-gaap.Revenues
        └── README.md                         # fixture 用途说明
```

**复用现有 samples/**:Phase 4d/4f/4h 已经收集了一些 sample(eg `xueqiu_HK_00700_balance.json`、`AAPL_edgar.json`、`xueqiu_kline_USDCNY.json`、`stockanalysis_KR_005930_income.html` 等)— **优先 copy 到 tests/fixtures/ 而不是重新抓取**。

**新建的 mock fixtures**(http_429_response.json / malformed_xueqiu.txt / missing_fields_edgar.json)手工构造或从现有 sample 改造。

**子任务 1B — VBA `模块_测试.bas` 新增 fixture loader**:

```vba
' Phase 4m Step 1: 读 tests/fixtures/<name> 文件返回 String
'   不发 HTTP, 让测试可以离线跑
'   path 相对于 xlsm 所在目录
Public Function LoadFixture(ByVal fileName As String) As String
    Dim fixturePath As String
    fixturePath = ThisWorkbook.Path & "\tests\fixtures\" & fileName
    
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fixturePath) Then
        Err.Raise vbObjectError + 700, "LoadFixture", _
            "Fixture not found: " & fixturePath
    End If
    Dim ts As Object: Set ts = fso.OpenTextFile(fixturePath, 1, False, -1) ' ForReading, Unicode
    LoadFixture = ts.ReadAll
    ts.Close
    Exit Function
EH:
    LoadFixture = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
```

**子任务 1C — 8 个测试 Sub**:

`模块_测试.bas` 新增以下 Public Sub(命名前缀 `Test_Offline_`):

```vba
Public Sub Test_Offline_US_Edgar_AAPL()
    ' 读 sec_aapl_companyfacts.json fixture
    ' 调 ParseEdgarCompanyFacts 解析 (现有 helper)
    ' 断言: us-gaap.Revenues 存在 + 单位为 USD + 至少 5 个 fiscal periods
    
    Dim body As String: body = LoadFixture("sec_aapl_companyfacts.json")
    Dim dictData As Object: Set dictData = ... ' parse
    
    AssertTrue dictData.Exists("Revenues"), "Revenues field missing"
    AssertTrue dictData("Revenues").Count >= 5, "Revenues should have >= 5 periods"
    Debug.Print "Test_Offline_US_Edgar_AAPL: PASS"
End Sub

Public Sub Test_Offline_HK_Xueqiu_Tencent()
    ' 读 xueqiu_hk_00700_balance.json
    ' 断言: tta (Total Assets) > 0 + reporting currency = HKD or CNY
    ...
End Sub

Public Sub Test_Offline_KR_StockAnalysis_Samsung()
    ' 读 stockanalysis_kr_005930_income.html
    ' 断言: 解析出 Revenue / Operating Income / Net Income
    ...
End Sub

Public Sub Test_FX_Missing_DoesNotFallbackToOne()
    ' 调 GetFxRateStatus("USD", "2099-01-01" 不存在的日期, True)
    ' 断言: status = "FX_MISSING" + outRate = 0 (Phase 4k 行为)
    ...
End Sub

Public Sub Test_Diagnostic_Score_NotDate()
    ' 写 "1/1" 到诊断 sheet I 列
    ' 断言: 读出来仍是 "1/1" 文本, 不是 46023 (Phase 4k 行为)
    ...
End Sub

Public Sub Test_Cache_HitMissExpired()
    ' 1. ClearLocalCache
    ' 2. WriteLocalHttpCache "test_key", "test_body"
    ' 3. 调 ReadLocalHttpCacheWithAge "test_key", 24, age → 期望 age ~ 0, body = "test_body"
    ' 4. 断言: cache HIT
    ...
End Sub

Public Sub Test_AppState_RestoreAfterError()
    ' 1. BeginAppState (改 Calculation 到 manual)
    ' 2. 故意 raise 错误
    ' 3. EndAppState
    ' 4. 断言: Application.Calculation 恢复到 automatic
    ...
End Sub

Public Sub Test_DataQuality_BS_Imbalance_Detection()
    ' 用 fixture: missing_fields_edgar.json (人造 BS 不平衡数据)
    ' 调 RunDataQualityChecks (Step 2 新增)
    ' 断言: 诊断行出现 "BS_NOT_BALANCED"
    ...
End Sub
```

**子任务 1D — 测试 runner**:

新建 `tools/run_offline_tests.py`:

```python
"""
Phase 4m Step 1: 离线测试 runner
逐个调 模块_测试.Test_Offline_* macros, 捕获 PASS/FAIL, 报告总数
"""
import sys
from pathlib import Path
import win32com.client as win32

XLSM = Path(r"E:\...\上市公司财务数据查询.xlsm")

TEST_MACROS = [
    "Test_Offline_US_Edgar_AAPL",
    "Test_Offline_HK_Xueqiu_Tencent",
    "Test_Offline_KR_StockAnalysis_Samsung",
    "Test_FX_Missing_DoesNotFallbackToOne",
    "Test_Diagnostic_Score_NotDate",
    "Test_Cache_HitMissExpired",
    "Test_AppState_RestoreAfterError",
    "Test_DataQuality_BS_Imbalance_Detection",
]

def main():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(str(XLSM))
    
    passed, failed = 0, 0
    for macro in TEST_MACROS:
        try:
            excel.Run(f"模块_测试.{macro}")
            print(f"  + {macro}: PASS", flush=True)
            passed += 1
        except Exception as e:
            print(f"  ! {macro}: FAIL — {e}", flush=True)
            failed += 1
    
    wb.Close(SaveChanges=False)
    excel.Quit()
    
    print(f"\nSUMMARY: {passed}/{len(TEST_MACROS)} PASS")
    if failed > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
```

### Assert helpers

`模块_测试.bas` 新增私有 helpers:

```vba
Private Sub AssertTrue(ByVal cond As Boolean, ByVal msg As String)
    If Not cond Then
        Err.Raise vbObjectError + 701, "AssertTrue", "Assertion failed: " & msg
    End If
End Sub

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal msg As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 702, "AssertEquals", _
            msg & " | expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub
```

### 验证

- 跑 `py tools/run_offline_tests.py` → **8/8 PASS**
- 故意 corrupt 一个 fixture(eg 把 sec_aapl_companyfacts.json 删掉 Revenues 字段)→ Test_Offline_US_Edgar_AAPL FAIL,error message 清晰
- 不需要联网

### Generator 不要做

- ❌ 不要把 fixture 提交到 git(`tests/fixtures/` 加 .gitignore 规则,sample 数据可能含个人化信息)
- ❌ 不要让 fixture 测试调用真 HTTP(本质是 unit test,不能依赖网络)
- ❌ 不要在测试里改 xlsm 任何 sheet(测试无副作用)

---

## Step 2 — P1-06 数据质量 QA(精简 3 条)

### 背景

GPT 5.5 Pro 提议 8 条 QA,Reviewer 精简到 3 条最稳的(避免 false positive)。这 3 条都是**会计常识级别的硬规则**:

1. **BS 平衡**:`|Assets - Liabilities - Equity| / Assets > 1%` → 字段映射可能错
2. **FX 缺失**:跨市场指标表 cell 出现 `#N/A` 或空(Phase 4k 已有 FX_MISSING 诊断,这条是聚合统计)
3. **关键字段缺失**:8 个核心字段任一公司缺失 → KEY_FIELD_MISSING

### 实现

**子任务 2A — `模块_工具函数.bas` 新增 `RunDataQualityChecks`**:

```vba
' Phase 4m Step 2: 跨市场指标表生成完毕后自动跑 3 条 QA 规则
'   写到诊断 sheet (A 股诊断 sheet 不存在所以改写"美股_抓取诊断" or 新增 "数据质量诊断" sheet)
' 选择: 写到 美股_抓取诊断 sheet 末尾追加, 用 "QA_" 前缀的 公司名 = "GLOBAL"

Public Sub RunDataQualityChecks()
    Dim diagWs As Worksheet
    Set diagWs = ThisWorkbook.Sheets("美股_抓取诊断")  ' 借用美股诊断作为 QA 输出
    
    ' Rule 1: BS 平衡 (跨 4 市场 BS 数据)
    Dim bsViolations As Long
    bsViolations = CheckBalanceSheetEquation()
    
    ' Rule 2: FX missing 聚合
    Dim fxMissingCount As Long
    fxMissingCount = CountFxMissingDiagnostics()
    
    ' Rule 3: 关键字段缺失 (8 项: Revenue / Total Assets / Total Liab / Equity / Net Income / CFO / EPS Basic / Cash)
    Dim keyMissing As Long
    keyMissing = CheckKeyFieldsPresence()
    
    ' 写 QA 汇总行 到诊断 sheet
    AddDiagnosticQARow "BS_BALANCE", "BS 平衡检查", _
        IIf(bsViolations = 0, "OK", "FAIL"), _
        bsViolations & " companies with |A - L - E| / A > 1%"
    AddDiagnosticQARow "FX_MISSING", "汇率缺失统计", _
        IIf(fxMissingCount = 0, "OK", "WARN"), _
        fxMissingCount & " cells with FX_MISSING"
    AddDiagnosticQARow "KEY_FIELDS", "关键字段覆盖", _
        IIf(keyMissing = 0, "OK", "WARN"), _
        keyMissing & " missing key fields across 4 markets"
End Sub

Private Function CheckBalanceSheetEquation() As Long
    ' 遍历 4 张 BS sheet, 找每家公司每期的 Total Assets / Total Liab / Total Equity
    ' 计算 |A - L - E| / A, 返回 > 1% 的违规数
    Dim violations As Long
    Dim markets As Variant: markets = Array("A股", "美股", "港股", "韩股")
    Dim m As Variant
    For Each m In markets
        ' ... 遍历 X股_资产负债表, 找 Total Assets / Total Liab / Total Equity 行
        ' 累加 violations
    Next m
    CheckBalanceSheetEquation = violations
End Function

Private Function CountFxMissingDiagnostics() As Long
    ' 扫描 3 张诊断 sheet 的 状态 列, 计 "FX_MISSING" 出现次数
    ...
End Function

Private Function CheckKeyFieldsPresence() As Long
    ' 8 个关键字段 (Revenue / Total Assets / etc.)
    ' 遍历 4 张 BS / IS / CF, 找哪些公司哪个字段为空
    ...
End Function

Private Sub AddDiagnosticQARow(ByVal qaCode As String, ByVal qaName As String, _
                                ByVal status As String, ByVal detail As String)
    ' 写一行到诊断 sheet
    '   公司 = "GLOBAL_QA"
    '   报表 = qaCode
    '   输出指标 = qaName
    '   状态 = OK / WARN / FAIL
    '   匹配方式+备注 = detail
    ...
End Sub
```

**子任务 2B — `BuildCrossMarketIndicatorSheet` 末尾自动调**:

`模块_工具函数.bas` `BuildCrossMarketIndicatorSheet` 函数末尾(在 `Application.DisplayAlerts = True` 之前)追加:

```vba
' Phase 4m Step 2: 自动跑 3 条 QA 规则
On Error Resume Next
RunDataQualityChecks
Err.Clear
On Error GoTo 0
```

**子任务 2C — 跨市场指标表 A1 注释**:

更新 `BuildCrossMarketIndicatorSheet` 的 A1 cell comment,加一行:
```
"数据质量检查结果在 美股_抓取诊断 sheet 末尾的 GLOBAL_QA 行查看"
```

### 验证

- 跑 `一键全抓` → 跨市场指标表生成完后,美股_抓取诊断 sheet 出现 3 行 GLOBAL_QA(BS_BALANCE / FX_MISSING / KEY_FIELDS)
- 故意改 A股_资产负债表 某公司 Total Assets 为 0 → BS_BALANCE 状态变 FAIL,detail 显示 1 violation

### Generator 不要做

- ❌ 不要新建 `数据质量检查` sheet(用现有诊断 sheet 末尾追加,避免 sheet 数爆炸)
- ❌ 不要把 QA 规则做成阻塞型(eg "BS 不平衡时禁止生成跨市场表")— QA 是提示,不是 gating
- ❌ 不要为 互联网公司 这种特殊业态 加 sector-specific 例外(本期 3 条都是会计硬规则,不分 sector)

---

## Step 3 — P1-02 cache 分源 TTL

### 背景

Phase 4l 引入 `RunCachedHttpGet(url, cacheKey, source, ttlHours, ByRef result)` 接受 ttlHours 参数,但**所有 caller 都传 24h**,等于变相 24h 一刀切。

不同数据源更新频率差异巨大:
- SEC ticker 映射表(几乎不变)→ 168h(1 周)
- EDGAR companyfacts(季报后更新)→ 24h(每日刷新即可)
- 雪球 HK 财报(实时变化少)→ 12h
- stockanalysis HTML(更新慢)→ 24h
- FX K 线(每日 close 后稳定)→ 24h
- 雪球 quote API(实时数据,本工具不缓存)→ N/A

### 实现

**子任务 3A — `模块_工具函数.bas` 新增 TTL map**:

```vba
' Phase 4m Step 3: source-aware cache TTL
' 不同数据源不同 TTL, 避免 24h 一刀切
Public Function GetTtlHoursForSource(ByVal source As String) As Long
    Select Case UCase$(Trim$(source))
        Case "SEC_TICKER_MAP":      GetTtlHoursForSource = 168    ' 1 周, 几乎不变
        Case "EDGAR_COMPANYFACTS":  GetTtlHoursForSource = 24     ' 每日刷新
        Case "EDGAR":               GetTtlHoursForSource = 24     ' 同上 (兼容老 source 名)
        Case "XUEQIU":              GetTtlHoursForSource = 12     ' 12h
        Case "XUEQIU_HK":           GetTtlHoursForSource = 12
        Case "STOCKANALYSIS_KR":    GetTtlHoursForSource = 24
        Case "STOCKANALYSIS_US":    GetTtlHoursForSource = 24
        Case "FX_KLINE":            GetTtlHoursForSource = 24
        Case Else:                  GetTtlHoursForSource = 24     ' 默认 24h
    End Select
End Function
```

**子任务 3B — Caller 改用 GetTtlHoursForSource**:

grep 所有 `RunCachedHttpGet(...)` 调用点(应该在 模块_抓美股财报 / 模块_抓港股财报 / 模块_工具函数 / 等),把 hardcoded `24` 参数改成 `GetTtlHoursForSource(source)`。

**子任务 3C — 诊断 sheet 显示实际 TTL**:

诊断 sheet `CacheStatus` 列下的备注或 `CacheAgeHours` 列旁边可以加显示当时用的 TTL — 但 Phase 4l 17 列已经够紧了,**本期不加列**。改成:`CacheStatus` 文本里 append TTL,eg `"HIT(24h)"` / `"HIT(168h)"`。

或者 — **不改 CacheStatus 显示**,让 TTL 只在 `RunCachedHttpGet` 内部生效,reviewer 想知道 source TTL 看 GetTtlHoursForSource 函数定义即可。**推荐这个**(简单)。

### 验证

- VBE 调 `?GetTtlHoursForSource("SEC_TICKER_MAP")` → 返回 168
- 跑 `一键美股` → 诊断 sheet 看到 EDGAR companyfacts cache 24h(MISS 后 HIT,跟 Phase 4l 一样)
- (无法直接观察:SEC ticker map 真的 168h cache,因为短时测不出 24h vs 168h 差别。**通过代码 review 验证**即可)

### Generator 不要做

- ❌ 不要把 TTL hardcode 在 caller 里(eg 美股 fetch 写 `RunCachedHttpGet(..., 24, ...)`)— 必须调 `GetTtlHoursForSource`
- ❌ 不要把 TTL > 1 周(168h)(企业财报更新有滞后,1 周已是上限)
- ❌ 不要给 cookie / 失败响应 设 TTL(Phase 4l 已规定不缓存)

---

## Step 4 — 端到端回归 + STATUS §DD 收口

### 4A. 跑 frozen 6 张

```bash
py tools/test_fx_live.py --skip-install
py -u tools/diff_phase4f_step3_lite.py
py -u tools/inspect_phase4g_state.py
py -u tools/inspect_phase4h_state.py
py -u tools/inspect_phase4k_state.py
py -u tools/inspect_phase4l_state.py
```

任一退化立即停下。

### 4B. 跑新增 8 个离线测试

```bash
py tools/run_offline_tests.py
```

期望:**8/8 PASS**。

### 4C. 新增 `tools/inspect_phase4m_state.py`

检查项:
1. `tests/fixtures/` 目录存在 + 至少 7 个 fixture 文件
2. `模块_测试.bas` 含 8 个 Test_Offline_* 函数定义(grep 验证)
3. `RunDataQualityChecks` 函数存在;跑 `BuildCrossMarketIndicatorSheet` 后,美股_抓取诊断 sheet 末尾出现 3 行 `公司=GLOBAL_QA`
4. `GetTtlHoursForSource` 函数存在 + 返回值正确(eg `SEC_TICKER_MAP=168, EDGAR=24, XUEQIU=12`)

### 4D. STATUS §DD 收口

模仿 §CC 格式追加:

```markdown
## DD. Phase 4m 收口: 优化 Sprint 3 — 离线测试 + 数据质量 QA + cache 分源 TTL

执行依据: PHASE_4M_PLAN.md v1。状态: ✅ Codex 已实现并通过 6 张 frozen 回归 + 8 离线测试 + 新增 inspect。

### DD.1 已完成
- [Step 1 P2-03] 8 个离线 fixture + Test_Offline_* 宏 + tools/run_offline_tests.py runner
- [Step 2 P1-06] 3 条 QA 规则 (BS 平衡 / FX missing 统计 / 关键字段覆盖) 自动写诊断 GLOBAL_QA 行
- [Step 3 P1-02] cache 分源 TTL: SEC ticker map 168h / EDGAR 24h / 雪球 12h / stockanalysis 24h / FX 24h

### DD.2 验证结果
[4A 回归 + 4B 离线 8 测试 + 4C inspect 结果]

### DD.3 已知边界
- 离线 fixture 用 samples/ 已有数据 + 3 个 mock fixture (429/malformed/missing); 不含真 cookie/token
- QA 写到现有诊断 sheet 末尾, 用 公司=GLOBAL_QA 标识; 不另起 sheet
- TTL map hardcode 在 GetTtlHoursForSource, 改 TTL 需要重装 xlsm
```

PHASE_4M_PLAN.md v1 → v2,标记 ✅ Phase 4m 全期闭环。

---

## ⚠️ 全 Phase 严禁动

| 文件/区域 | 原因 |
|---|---|
| `模块_抓汇率.bas` / 旧 `GetFxRate` 系签名 | Phase 4f / 4k frozen |
| 核心 `EdgarHttpGet / XueqiuHttpGet / StockAnalysisHttpGet` 签名 | Phase 4l frozen |
| `RunCachedHttpGet / HttpGetWithRetry / EnforceSecRateLimit / THttpResult` | Phase 4l frozen,本期只调用不改 |
| `CleanReleaseWorkbook` | Phase 4l frozen |
| `汇率` sheet 8 列结构 + 4 市场 fetch 字段映射 | 长期 frozen |
| 16 张分市场 sheet + 跨市场指标表 内容 | Phase 4j frozen |
| 6 张 frozen 回归驱动核心断言 | Phase 4f-4l 验证基线 |
| 样本池 R14+ 用户数据 | 数据安全 |

## ⚠️ 联系 Planner 触发条件

- Step 1 离线 fixture 解析失败(eg AAPL companyfacts 字段路径变了)→ 说明 fixture 旧 / 解析逻辑有 bug,先排查
- Step 2 BS 平衡 QA 误报率 > 50%(eg 真实公司也被标 FAIL)→ 说明 1% 阈值太严,调宽
- Step 3 改 TTL map 后任一 frozen 回归退化(说明 cache 行为改变)
- 任一 frozen 回归 PASS → FAIL
- 离线测试 8/8 < 6/8(超过 25% 失败说明 fixture 设计有问题)

## State-bound inspect 同步规则

本期不改诊断 sheet 列结构(只 append `公司=GLOBAL_QA` 行,跟现有诊断行同列结构),所以 `inspect_phase4h/4k/4l_state.py` **不需要改动**。

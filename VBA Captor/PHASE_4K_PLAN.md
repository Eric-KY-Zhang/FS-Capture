# Phase 4k: 优化 Sprint 1 — 数据准确性 + UX live FX + 状态守护

> **版本**: v2(2026-05-05,Phase 4k 优化 Sprint 1 闭环)
> **状态**: ✅ Phase 4k 全期闭环
> **作者**: Claude(planner) + Codex(generator)
> **背景**: Phase 4f-4j 累计 9 phase / 17 commit 完成 RMB 换算 / 跨市场 / 缓存 / fallback / UX 全套核心开发。GPT 5.5 Pro 静态审阅工具后给出 14 项优化 backlog(`codex_vba_tool_optimization_plan.md`,已读完后删除)。Phase 4k 取其中 4 项最高 ROI 的修 bug + UX 改进,作为优化阶段 Sprint 1。

## 项目语境(给 Generator 的 anchor 段)

本期 4 项工作**全部在现有数据流内**,不引入新数据源 / 新抓数频次 / 新并发。**唯一带"行为变化"的是** Step 3:汇率缺失时不再静默 fallback 到 1,而是写空值 + 诊断标记 — 这是修一个长期潜在的数据准确性 bug(KRW 汇率缺失 → 韩股按 1:1 算 → 数值差 ~200x),不是新增功能。其他 3 个 Step 都是局部 fix 或 UX 改进。

## 4 项任务清单(基于 GPT 5.5 Pro 审阅,Reviewer 调整优先级)

| # | Task | 严重性 | 工作量 |
|---|---|---|---|
| 1 | **P0-02** 韩股诊断 Score 列日期化 bug 修复 | 真 bug(诊断信息丢失)| 0.5h |
| 2 | **P0-03** AppStateGuard 状态守护(1 个入口示范)| 防御性(避免宏报错后 Excel 残留状态)| 1.5h |
| 3 | **P0-05** 汇率缺失不再 fallback 1 | **严重数据准确性 bug**(韩股可能差 ~200x)| 3h |
| 4 | **P1-05** 报表公式 live ref 汇率表 | UX 大改进(汇率手改实时刷新)| 2.5h |

总耗时 ~7.5h Codex,~1h Reviewer。

## Step 总览

| Step | 内容 | 估时 | 阻塞依赖 |
|---|---|---|---|
| 1 | P0-02:韩股诊断 Score 列改文本格式 | 0.5h | 无 |
| 2 | P0-03:新建 `模块_AppStateGuard.bas` + 给 `一键全抓` 入口示范使用 | 1.5h | 无 |
| 3 | P0-05:`GetFxRate` 增加 status 返回,缺失不再 fallback 1 + 诊断写 `FX_MISSING` | 3h | 无 |
| 4 | P1-05:`WriteWideTable` 公式从 baked-in `* 7.1234` 改成 `* 汇率!$X$Y` cell-ref | 2.5h | Step 3(共享 GetFxRate 改动)|
| 5 | 端到端回归 4 张 + 新增 `inspect_phase4k_state.py` + STATUS §BB 收口 | 0.5h | 1-4 |

**Codex 工作流建议(2 round)**:
- **Round 1** = Step 1 + Step 2(独立小 task 先暖身,确认 Codex 跟新 plan 节奏对齐)
- **Round 2** = Step 3 + Step 4 + Step 5(FX 重构 + 公式 live ref + 收口,这两个 step 共享 GetFxRate 改动,一起做减少 merge 冲突)

如果 Codex 倾向端到端 1 round 也可以(用户最近偏好 single commit),只是 review 范围会大一些。

---

## Step 1 — P0-02 韩股诊断 Score 列日期化修复

### 背景

`模块_抓韩股财报.bas` 的 `AddOrUpdateKRMatch` 写诊断行 Score 列时用 `CStr(candIdx) & "/" & CStr(totalCand)`(eg `"1/1"`、`"2/3"`)。Excel 把 `1/1` 自动识别为日期 → 显示 `46023`(2026-01-01 的序列值),诊断信息**完全失真**,reviewer 看不出字段命中质量。

### 修复方案(任选其一,推荐 A)

**A. 改成数值 Score + 拆分 Rank/Total**(GPT 推荐)
- 把"匹配方式"语义化为数值:exact match=100 / alias match=85 / fuzzy match=70 等
- 单独列 `CandidateRank` / `CandidateTotal`(整数列)
- 诊断 sheet 加新列(K11+ 或在现有 11 列基础上扩展)

**B. 保留文本 Score,强制文本格式**(简单)
- 写表前 `ws.Columns("I:I").NumberFormat = "@"`(I 列 = Score 列,Phase 4f 已按 11 列设计)
- 写入值前加 apostrophe:`arr(r, 9) = "'" & scoreText`
- 不增加列,不破坏现有诊断结构

### 推荐 B(简单 + 不改 schema)

**子任务 1A**:
- `tools/install_modules.py` `_make_diagnostic_sheet` 函数:I 列(Score 列)新增 `NumberFormat = "@"`
- `tools/install_modules.py` `_refresh_diagnostic_headers` 函数:同样设置 I 列文本格式
- `模块_工具函数.bas` 写诊断行的 helper(grep `arr(.*, 9)` 或 `Score`):写入前加 `"'"` 前缀
- 韩股 / 港股 / 美股 3 张诊断 sheet 全覆盖(不只韩股,3 个市场都用同一 helper)

### 验证

- 跑一次 `一键韩股`(任意ticker eg 005930),打开 `韩股_抓取诊断` sheet → I 列 Score 显示 `1/1`、`2/3` 文本而非 `46023` 序列值
- 如果之前已经被日期化的旧诊断行,可以手工清空诊断 sheet 然后重抓覆盖

### Generator 不要做

- ❌ 不要改诊断 sheet 列数(11 列 frozen)
- ❌ 不要改 Score 内容语义(eg 改成 100/85 数值 — 那是方案 A,本期选 B)

---

## Step 2 — P0-03 AppStateGuard(1 入口示范)

### 背景

VBA 入口宏(eg `一键全抓`)中途出错时,如果没恢复 `Application.ScreenUpdating / DisplayAlerts / Calculation / EnableEvents`,用户的 Excel 会留在"屏幕不刷新 / 弹窗禁用 / 公式不重算"的脏状态,要重启 Excel 才能恢复。当前代码部分恢复了 `ScreenUpdating + StatusBar`,但其他属性(Calculation / EnableEvents / DisplayAlerts)漏了。

### 实现

**子任务 2A — 新建 `modules/模块_AppStateGuard.bas`**:

```vba
Option Explicit

' Phase 4k Step 2: 全 Excel 状态守护,任何入口宏报错时一次性恢复
'   用法:
'     Public Sub 一键全抓()
'         Dim st As TAppState
'         On Error GoTo EH
'         st = BeginAppState("正在全量抓取...")
'         ' ... main logic ...
'     CleanExit:
'         EndAppState st
'         Exit Sub
'     EH:
'         Application.StatusBar = "一键全抓出错: " & Err.Description
'         Resume CleanExit
'     End Sub

Public Type TAppState
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    DisplayStatusBar As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    StatusBarValue As Variant
    DisplayPageBreaks As Boolean
    HasActiveSheet As Boolean
End Type

Public Function BeginAppState(Optional ByVal statusText As String = "") As TAppState
    Dim st As TAppState
    st.ScreenUpdating = Application.ScreenUpdating
    st.DisplayAlerts = Application.DisplayAlerts
    st.DisplayStatusBar = Application.DisplayStatusBar
    st.EnableEvents = Application.EnableEvents
    st.Calculation = Application.Calculation
    st.StatusBarValue = Application.StatusBar
    st.HasActiveSheet = Not (ActiveSheet Is Nothing)
    If st.HasActiveSheet Then st.DisplayPageBreaks = ActiveSheet.DisplayPageBreaks
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    If st.HasActiveSheet Then ActiveSheet.DisplayPageBreaks = False
    If Len(statusText) > 0 Then Application.StatusBar = statusText
    
    BeginAppState = st
End Function

Public Sub EndAppState(ByRef st As TAppState)
    On Error Resume Next
    Application.Calculation = st.Calculation
    Application.EnableEvents = st.EnableEvents
    Application.DisplayStatusBar = st.DisplayStatusBar
    Application.DisplayAlerts = st.DisplayAlerts
    Application.ScreenUpdating = st.ScreenUpdating
    Application.StatusBar = st.StatusBarValue
    If st.HasActiveSheet Then
        On Error Resume Next
        ActiveSheet.DisplayPageBreaks = st.DisplayPageBreaks
    End If
    Err.Clear
    On Error GoTo 0
End Sub
```

**子任务 2B — 改造 `模块_总入口.bas` 的 `一键全抓` 作为示范**:

把现有 `一键全抓` Sub 的开头/结尾包成 `BeginAppState/EndAppState` pattern,保留所有现有抓数 + 末尾 `BuildCrossMarketIndicatorSheet` + `UnhideMarketTabs` 调用。

**子任务 2C — Python 装表同步**:
- `tools/install_modules.py` 把 `模块_AppStateGuard.bas` 加入 `MODULE_FILES` 列表(让装表脚本自动注入)

### 不要做(本 step 范围控制)

- ❌ **不要**给所有 7 个入口宏(`一键A股 / 一键美股 / 一键港股 / 一键韩股 / 一键全抓 / 切换X股tabs ×4 / 显示/隐藏所有市场数据 等)全部改造**
- 只改 `一键全抓` 1 个作为示范。**其他入口下个 sprint 增量做**(避免一次性改太多引入回归)
- ❌ 不要修改 `一键X股`(单市场)入口

### 验证

- VBE 编译 `Debug > Compile VBAProject` 通过
- 故意制造一个错误(eg 把 `BuildCrossMarketIndicatorSheet` 改成不存在的名字),跑 `一键全抓` → 看 Excel `Application.Calculation` 在错误后能恢复 `xlCalculationAutomatic`

### Generator 不要做

- ❌ 不要改其他 6 个入口宏
- ❌ 不要扩大 `On Error Resume Next` 范围(只在 `EndAppState` 内部用)

---

## Step 3 — P0-05 汇率缺失不再 fallback 1

### 背景

**这是 Phase 4k 最重要的 fix**。当前 `WriteWideTable` line 1417-1421 区域:

```vba
fx = GetFxRate(curCode, strPeriod, useEopForBS)
If fx > 0 Then
    writeVal = CDbl(rawVal) * fx
Else
    writeVal = rawVal    ' ← 这里是 bug:fx <= 0 时 writeVal = 原币 raw,但单元格按 RMB 显示
End If
```

实际行为:如果 KRW 汇率缺失(eg 用户没跑过 `EnsureFxRateCached`),`GetFxRate` 返回 0,写入单元格的是**韩元原值**,但 sheet 标题显示"统一RMB" → 用户以为这是 RMB 数值,实际差 ~200x(KRW:RMB ≈ 200:1)。

**审计/财务专业人士看到这种数值会做错决策**。

### 实现

**子任务 3A — `GetFxRate` 改进 status 返回**:

`模块_工具函数.bas` line ~690 区域 `GetFxRate` 函数:
- **保留旧签名**:`Public Function GetFxRate(curCode As String, periodKey As String, useEop As Boolean) As Double` 不动(回归驱动 `test_fx_live.py` 调它)
- **新增 helper**:`Public Function GetFxRateStatus(curCode As String, periodKey As String, useEop As Boolean, ByRef outRate As Double) As String`
  - 返回值是 status 字符串:
    - `"OK"` — 找到汇率
    - `"RMB_BASE"` — 报告币种 = RMB / CNY,outRate = 1.0
    - `"FX_MISSING"` — 汇率 sheet 无对应行 / 列,outRate = 0
    - `"FX_FETCH_FAILED"` — 触发 EnsureFxRateCached 但失败,outRate = 0
  - 旧 `GetFxRate` 内部调用 `GetFxRateStatus`,只返回 outRate(向后兼容)

**子任务 3B — `WriteWideTable` 改用新 helper**:

把 line 1417-1421 改成:

```vba
Dim fx As Double, fxStatus As String
fxStatus = GetFxRateStatus(curCode, strPeriod, useEopForBS, fx)
Select Case fxStatus
    Case "OK", "RMB_BASE":
        writeVal = CDbl(rawVal) * fx    ' RMB_BASE 时 fx=1.0,等价短路
    Case "FX_MISSING", "FX_FETCH_FAILED":
        writeVal = ""                   ' 写空值,不要静默写原币
        ' 同步写诊断行
        AddDiagnosticFxMissing strCode, strInd, strPeriod, curCode, fxStatus
    Case Else:
        writeVal = ""
End Select
```

**子任务 3C — 诊断 sheet 加 FX 状态行**:

新增 helper `AddDiagnosticFxMissing(code, indicator, period, currency, status)`,在对应市场的诊断 sheet 追加 1 行:
- 公司 = code
- 报表 = "FX_CONVERSION"
- 输出指标 = indicator
- 状态 = status(eg `"FX_MISSING"`)
- 数据源 = "FX_Sheet"
- Taxonomy = (空)
- 命中字段 = currency
- Unit = (空)
- Score = (空)
- 匹配方式+备注 = "汇率缺失,统一RMB 模式下该 cell 留空,请检查汇率 sheet 或重跑 EnsureFxRateCached"
- FX_Rate = 0

**子任务 3D — RMB / CNY 短路保留**:

确保 `GetFxRateStatus` 收到 `curCode = "RMB"` 或 `"CNY"` 时直接返回 `"RMB_BASE"` + outRate=1.0,不查汇率 sheet。这是 Phase 4f 已有行为,不要破坏。

### 验证

- 跑 `test_fx_live.py --skip-install` → **5/5 PASS**(USD/HKD/KRW 都有 cache,不触发 missing)
- 手工 corrupt 测试:打开 xlsm,把 汇率 sheet 的 `KRWCNY期均` 列(D 列?eg D2)清空,B6 设"统一RMB",跑 `一键韩股`,验证韩股报表 cell 留空(不是数字),诊断 sheet 出现 `FX_MISSING` 行
- 恢复 `KRWCNY期均` 数值,重跑,数值正常

### Generator 不要做

- ❌ 不要改 `EnsureFxRateCached` 的抓汇率逻辑(那是 Phase 4f frozen)
- ❌ 不要改 `汇率` sheet 的 8 列结构
- ❌ 不要修改旧 `GetFxRate` 签名(`test_fx_live.py` 调用方依赖旧签名)
- ❌ 不要在 `原币` 模式下 trigger FX_MISSING 诊断(原币模式下根本不调 GetFxRate)

---

## Step 4 — P1-05 报表公式 live ref 汇率表

### 背景

Phase 4h §W.3 已落账"汇率手改不会反向刷",当前 `WriteWideTable` 写出的公式是:

```excel
=IF(样本池!$B$6="统一RMB", H100*7.1234, H100)
```

`7.1234` 是写表时 Codex 计算 `GetFxRate(...)` baked into formula。用户改 `汇率!B2`(USDCNY期末)从 7.30 → 7.50 后,这个公式不会重算 — 必须重跑 `一键X股`。

**用户期望**:改 `汇率!B2` 后立即看到所有报表 cell 重算成新汇率值。

### 实现

**子任务 4A — `WriteWideTable` 公式改用 cell-ref**:

把当前 baked-in `* fx_value` 改成:

```excel
=IF(样本池!$B$6="统一RMB", H100*GetFxFromSheet("USD",$E$2,"AVG"), H100)
```

其中 `GetFxFromSheet` 是新 UDF(VBA 自定义函数):

```vba
' Phase 4k Step 4: Live FX lookup from 汇率 sheet
'   currencyCode: "USD" / "HKD" / "KRW" / "RMB" / "CNY"
'   periodEnd: 报告期 cell 值 (Date or yyyy-mm-dd string)
'   rateKind: "EOP" or "AVG"
' 返回 Double or #N/A 错误
Public Function GetFxFromSheet(ByVal currencyCode As String, _
                                ByVal periodEnd As Variant, _
                                ByVal rateKind As String) As Variant
    On Error GoTo EH
    Dim curCode As String: curCode = UCase$(Trim$(CStr(currencyCode)))
    If curCode = "RMB" Or curCode = "CNY" Then
        GetFxFromSheet = 1#
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("汇率")
    Dim periodKey As String: periodKey = Format(CDate(periodEnd), "yyyy-mm-dd")
    
    ' 找到 periodKey 对应的 row (汇率 sheet R2+ 报告期列 A)
    Dim row As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For row = 2 To lastRow
        If Format(ws.Cells(row, 1).Value, "yyyy-mm-dd") = periodKey Then
            Dim col As Long
            Select Case curCode & "_" & UCase$(rateKind)
                Case "USD_EOP": col = 2
                Case "USD_AVG": col = 3
                Case "HKD_EOP": col = 4
                Case "HKD_AVG": col = 5
                Case "KRW_EOP": col = 6
                Case "KRW_AVG": col = 7
                Case Else:
                    GetFxFromSheet = CVErr(xlErrNA)
                    Exit Function
            End Select
            ' 优先 H 列 override(Phase 4f 设计)
            If IsNumeric(ws.Cells(row, 8).Value) And ws.Cells(row, 8).Value > 0 Then
                ' Override 列复杂(可能多个 override 字段),保留为 future enhancement
            End If
            If IsNumeric(ws.Cells(row, col).Value) And ws.Cells(row, col).Value > 0 Then
                GetFxFromSheet = CDbl(ws.Cells(row, col).Value)
            Else
                GetFxFromSheet = CVErr(xlErrNA)
            End If
            Exit Function
        End If
    Next row
    GetFxFromSheet = CVErr(xlErrNA)
    Exit Function
EH:
    GetFxFromSheet = CVErr(xlErrNA)
End Function
```

**子任务 4B — 改写 `WriteWideTable` formula 生成**:

当前 baked-in:
```vba
formulaCur = ... & "*" & fxString & ")"   ' fxString = "7.1234"
```

改成:
```vba
formulaCur = ... & "*GetFxFromSheet(""" & curCode & """,样本池!$E$2,""" & rateKind & """))"
```

注意 双引号 escape:`""USD""` 在 VBA 字符串里是 1 个 `"USD"`。

**子任务 4C — 性能 budget**:

UDF 比 baked-in 慢(每次 recalc 触发函数调用)。10 公司 × 5 期 × 18 指标 × 4 报表 = 3600 个 UDF 调用。如果每次 < 1ms,total ~3.6s。可接受。

如果实测 > 5s,降级到 baked-in 但加 `Worksheet_Change` 监听 `汇率` sheet 改动,改动后主动重写 baked-in。

### 验证

- 跑 `一键全抓 4 市场` → 数值正常显示 RMB 等价值
- 改 `汇率!B2`(USDCNY期末)从 7.30 → 8.00,**不点任何按钮** → 所有美股 BS cell 立即按新汇率刷新
- 跑 `diff_phase4f_step3_lite.py` → A 股 byte-identical(因为 A 股是 RMB,UDF 短路返回 1.0,跟 baked-in `*1` 等价)

### Generator 不要做

- ❌ 不要改 `汇率` sheet 8 列结构
- ❌ 不要改 H 列 override 语义(本 step 暂不实现 override 优先,留 future)
- ❌ 不要把 UDF 名定为非中文(VBA UDF 中文名容易跟 sheet 名混淆,用 `GetFxFromSheet` 英文)

---

## Step 5 — 端到端回归 + STATUS §BB 收口

### 5A. 跑 frozen 4 张

```bash
py tools/test_fx_live.py --skip-install      # GetFxRate 旧签名保留, 必 PASS
py -u tools/diff_phase4f_step3_lite.py        # A 股 byte-identical (RMB 短路)
py -u tools/inspect_phase4g_state.py          # state 不变
py -u tools/inspect_phase4h_state.py          # state 不变, 但 fallback 状态检查可能要微调
```

任一退化立即停下。

### 5B. 新增 `tools/inspect_phase4k_state.py`

检查 4 件事:
- 韩股诊断 Score 列 NumberFormat = "@"(Step 1)
- `模块_AppStateGuard` 模块存在 + `BeginAppState / EndAppState` 函数定义(Step 2)
- 故意 corrupt 汇率 sheet KRWCNY 后跑韩股 fetch(用 smoke macro),验证韩股报表 cell 留空 + 诊断行有 `FX_MISSING`(Step 3)— 然后恢复
- 改汇率 sheet USDCNY 数值后,美股报表 cell 数值变化(Step 4)

### 5C. STATUS §BB 收口

模仿 §AA 格式追加:

```markdown
## BB. Phase 4k 收口: 优化 Sprint 1 — 数据准确性 + UX live FX + 状态守护

执行依据: `PHASE_4K_PLAN.md` v1。状态: ✅ Codex 已实现并通过 4 张 frozen 回归 + 新增 inspect。

### BB.1 已完成
- [Step 1 P0-02] 韩股诊断 Score 列改文本格式, 不再被自动转日期
- [Step 2 P0-03] 新增 `模块_AppStateGuard.bas`, `一键全抓` 入口示范使用 `BeginAppState/EndAppState`
- [Step 3 P0-05] `GetFxRateStatus` 新 helper, 汇率缺失不再 fallback 1, 写空值 + 诊断 `FX_MISSING`
- [Step 4 P1-05] 报表公式 baked-in `*7.1234` 改成 `*GetFxFromSheet("USD",...)`, 改汇率 sheet 自动刷

### BB.2 验证结果
[5A 回归 + 5B inspect + 性能数据]

### BB.3 已知边界
- AppStateGuard 只 apply 到 `一键全抓` 入口, 其他 6 个入口下个 sprint 增量做
- FX_MISSING 时报表 cell 显示空, 不显示 #N/A(避免 Excel 错误传播)
- GetFxFromSheet UDF 不读 H 列 override, 留 future enhancement
```

PHASE_4K_PLAN.md v1 → v2,标记 ✅ Phase 4k 全期闭环。

---

## ⚠️ 全 Phase 严禁动

| 文件/区域 | 原因 |
|---|---|
| `模块_抓汇率.bas` | Phase 4f Step 2 frozen,只用其 cache,不动抓数逻辑 |
| `汇率` sheet 8 列结构 | Phase 4f frozen,只读不改 |
| 旧 `GetFxRate` 函数签名 | `test_fx_live.py` 依赖,本期保留兼容 |
| 4 市场 fetch 模块的字段映射 | Phase 4c-4h frozen |
| 16 张分市场 sheet 内容 | Phase 4j frozen |
| 跨市场指标表 | Phase 4g/4j frozen |
| 4 张 frozen 回归驱动核心断言 | Phase 4f-4h 验证基线 |
| 样本池 R14+ 用户数据 | 数据安全 |

## ⚠️ 联系 Planner 触发条件

- Step 3 改 `WriteWideTable` 后 `test_fx_live.py` 退化(说明 GetFxRate 兼容性破了)
- Step 4 UDF 实测性能 > 5s(需要决定降级到 Worksheet_Change 路径)
- Step 4 改公式后 `diff_phase4f_step3_lite.py` 退化(说明 A 股 RMB 短路在 UDF 里没正确实现)
- 任一 frozen 回归 PASS → FAIL
- VBE Compile 失败(语法错误)

## State-bound inspect 同步规则(Phase 4j.1 经验)

- `inspect_phase4h_state.py` 如果检查 fallback toggle / B6 行为 / 诊断列数 等, 在 Step 3 后可能需要微调断言(eg 诊断 sheet 多了 FX_MISSING 行类型)— **state-bound 同步,不算违反 frozen**
- 顶部加注释 `# Phase 4k: 同步 FX_MISSING 诊断行检查` 标识改动

---

## Codex 工作流建议(2 round 或 1 round 端到端)

### 选项 A — 2 round(分开)

- Round 1 = Step 1 + Step 2(独立小 task)
  - commit `Phase 4k Round 1 P0-02 KR Score + P0-03 AppStateGuard`
- Round 2 = Step 3 + Step 4 + Step 5(FX 重构 + UX live ref + 收口)
  - commit `Phase 4k Round 2 P0-05 FX missing handling + P1-05 live FX ref + closure`

### 选项 B — 1 round 端到端(用户最近偏好)

- 全部 5 step 一次 commit:`Phase 4k optimization sprint 1 (FX accuracy + live ref + state guard + KR score fix)`
- review 范围大(~1.5h Reviewer 工作),但用户上下文切换少
- 推荐选 B

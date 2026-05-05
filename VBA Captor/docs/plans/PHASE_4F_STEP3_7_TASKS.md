# Phase 4f Step 3-7 — Generator 细化 task list

> **作者**: Planner (Opus 4.7 1M)
> **日期**: 2026-05-04
> **状态**: READY FOR GENERATOR
> **范围**: Plan §Step 3 → §Step 7(全部本地 VBA 重构 + 模板调整 + 回归验证;无新增网络代码)
> **依据**:
> - `PHASE_4F_RMB_PLAN.md` v3 修订纪要(Step 2 联网验证通过后锁定)
> - `STATUS.md §T`(Step 2 收口章节,含 5/5 测试结果 + 架构决策留痕)
> - 现有 `modules/模块_工具函数.bas` `WriteWideTable`(line 1048)/ `EnsureDiagnosticSheet`(line 138)
> - 现有 `modules/模块_抓汇率.bas` `EnsureFxRateCached`(Step 2 已实现并联网验证 5/5)
> - 现有 `modules/模块_工具函数.bas` `GetFxRate`(line 690)/ `ReadDisplayCurrency`(line 666)/ `LookupFxColForCurrency`(line 783)

## 整体语境(给 Generator 的背景说明)

本项目是一个**给个人用户(财务/审计专业人士)用的 Excel 桌面报表工具**:把已有的 4 市场(A 股/港股/美股/韩股)上市公司年报数据,从 4 张分市场宽表里展示,供用户做同业对标。

Step 1-2 已经把"读雪球 K 线 → 算汇率 → 写到本地『汇率』sheet 缓存"的部分做完了,包括所有 HTTP 调用。**本批次(Step 3-7)只做本地的 cell 读写、算术换算、UI 刷新和回归验证,不引入任何新的网络调用**。汇率值已经躺在 `汇率` sheet 里,通过 `GetFxRate(curCode, periodEnd, useEop)` 这个本地函数读出来用就行(命中缓存 < 1ms)。

## Step 2 → Step 3-7 task 映射

| Step 2 已交付 | Step 3-7 复用方式 |
|---|---|
| `Public Function GetFxRate(curCode, periodEnd, useEop) As Double` | Step 3 在 `WriteWideTable` 写 cell 前调用 — 缓存命中 < 1ms |
| `Public Function ReadDisplayCurrency() As String`(读样本池 B6) | Step 3 用它判断是否启用 RMB 换算; Step 5 用它选 A1 注释文案 |
| 『汇率』sheet 7 列布局 + Row 1 表头 | Step 3 不改;数据已由 Step 2 自动写入 |
| `Public Const FX_SHEET / FX_DATA_ROW` | 不复用(本期只走 GetFxRate 抽象, 不直接 hit sheet) |
| `tools/test_fx_live.py` win32com 回归驱动 | Step 6 直接复用,确认 FX 没退化 |

---

## Step 3 — `WriteWideTable` 加换算 hook + 4 市场 wrapper 各自传报告币种

### 目标

让 `WriteWideTable` 在写每个数据 cell 前查 `ReadDisplayCurrency()`:

- 返 `"原币"` → 直接写原值(等同于现状,**完全向后兼容**)
- 返 `"统一RMB"` → 写 `原值 × GetFxRate(reportingCurrency, periodEnd, useEopForKind)`

`reportingCurrency` 从调用方传进来(每家公司一个值,可以不同 — 港股尤其需要):

- A 股调用方: 全部传 `"RMB"`(`GetFxRate("RMB", *, *) = 1.0` 自动短路)
- 美股调用方: 全部传 `"USD"`
- 港股调用方: 每家公司从 Step 2 已捕获的 `currencyText`(`模块_抓港股财报.bas` line 236 `currencyText = HK_NzStr(dataRoot, "currency")`)读;**注意港股大陆公司报告币种通常是 `CNY` 不是 `HKD`**(如 02519 傲基、09618 京东、09988 阿里);本地港股公司才报 `HKD`(如 00700 腾讯)。如果 `currencyText` 是 `"REPORT_CURRENCY"` 占位符或空,fallback 到 `"HKD"`
- 韩股调用方: 全部传 `"KRW"`

`useEopForKind` 由调用方根据本张表的 statement kind 决定:

- `"BalanceSheet"` → `useEop = True`(期末汇率,资产负债表是时点数)
- `"Income"` / `"CashFlow"` → `useEop = False`(期间均值,损益类是发生数)
- `"Indicator"` → 不调用 `WriteWideTable`(指标表本来是 Excel 公式自动算,不在 hook 范围)

### 子任务 3A — `WriteWideTable` 签名扩展

**3A.1** 在 `modules/模块_工具函数.bas` line 1048 修改 `WriteWideTable` 签名,新增 2 个 Optional 参数(放最后,旧调用 zero-arg 兼容,旧行为 = 等同 `"原币"` 模式):

```vba
Public Sub WriteWideTable(ByVal ws As Worksheet, _
                           ByRef arrCodes As Variant, _
                           ByRef dictCompanyName As Object, _
                           ByRef dictData As Object, _
                           ByRef arrPeriodsSorted As Variant, _
                           ByRef arrIndicators As Variant, _
                           ByRef dictCategory As Object, _
                           Optional ByVal perCompanyPeriods As Boolean = False, _
                           Optional ByRef dictReportingCurrency As Object = Nothing, _
                           Optional ByVal statementKind As String = "")
```

**3A.2** 在 `WriteWideTable` 内部、写 `arrOut(k, intCol + j - 1) = dictPer(strInd)`(约 line 1227)前,插入换算逻辑:

```vba
                            If dictPer.Exists(strInd) Then
                                Dim rawVal As Variant: rawVal = dictPer(strInd)
                                Dim writeVal As Variant: writeVal = rawVal
                                ' Phase 4f Step 3: 统一RMB 显示模式 → 按报告币种乘汇率
                                If displayMode = "统一RMB" And IsNumeric(rawVal) Then
                                    Dim curCode As String: curCode = "RMB"
                                    If Not dictReportingCurrency Is Nothing Then
                                        If dictReportingCurrency.Exists(strCode) Then _
                                            curCode = CStr(dictReportingCurrency(strCode))
                                    End If
                                    Dim fx As Double
                                    fx = GetFxRate(curCode, strPeriod, useEopForBS)
                                    If fx > 0 Then
                                        writeVal = CDbl(rawVal) * fx
                                    Else
                                        writeVal = rawVal    ' fx 拿不到, 退化为原币 (留 1.0 fallback 比写 0 安全)
                                    End If
                                End If
                                arrOut(k, intCol + j - 1) = writeVal
                            End If
```

**3A.3** 在 `WriteWideTable` 顶部读一次 `displayMode` 和 `useEopForBS`(避免在内层循环里反复读样本池 B6):

```vba
    ' Phase 4f Step 3: RMB 换算预读 (避免内循环反复读 B6)
    Dim displayMode As String: displayMode = ReadDisplayCurrency()
    Dim useEopForBS As Boolean
    useEopForBS = (UCase$(Trim$(statementKind)) = "BALANCESHEET")
```

**3A.4 容错原则**(写代码时务必照顾):

| 情形 | 行为 |
|---|---|
| `displayMode = "原币"` | 完全跳过换算(保持现状字节级一致) |
| `dictReportingCurrency Is Nothing` | 视为全 RMB(A 股快速路径) |
| `rawVal` 非数字(eg 字符串、Empty) | 直接写原值,不调 GetFxRate |
| `GetFxRate` 返 0(网络失败/缓存缺) | 退化写原值,**不写 0**(0 会让用户误以为公司财务为零) |
| `dictReportingCurrency` 里某家公司缺 key | 视为 `"RMB"` |

### 子任务 3B — A 股 4 个 Main 调用点改造

A 股的 Main 已经全部走 `模块_工具函数.bas` line 2250 的 `WriteWideTable wsTarget, ...`(共用主流程 `RunStatement`),所以**只需要在那个调用点的实参里加 `dictReportingCurrency = Nothing` + `statementKind = <kind>`**。

**3B.1** 在 `模块_工具函数.bas` line 2250 修改:

```vba
    WriteWideTable wsTarget, arrCodes, dictCompanyName, dictData, _
                    arrPeriods, arrIndicators, dictCategoryMap, _
                    perCompanyPeriods:=False, _
                    dictReportingCurrency:=Nothing, _
                    statementKind:=strKind
```

(`strKind` 在那个 Sub 里已经是局部变量,值是 `"BalanceSheet"` / `"Income"` / `"CashFlow"` 之一)

如果 `strKind` 不在那个 Sub 的作用域,从 `targetSheet` 反推:

```vba
    Dim hookKind As String
    Select Case True
        Case InStr(targetSheet, "资产负债") > 0: hookKind = "BalanceSheet"
        Case InStr(targetSheet, "利润") > 0:     hookKind = "Income"
        Case InStr(targetSheet, "现金流") > 0:   hookKind = "CashFlow"
        Case Else:                               hookKind = ""
    End Select
```

### 子任务 3C — 美股调用点改造

**3C.1** `模块_抓美股财报.bas` line 170 修改 `WriteWideTable` 调用,新增构造美股 dictReportingCurrency:

```vba
    Dim dictReportingCurrency As Object: Set dictReportingCurrency = CreateObject("Scripting.Dictionary")
    Dim usCode As Variant
    For Each usCode In arrCodes
        dictReportingCurrency(CStr(usCode)) = "USD"
    Next usCode
    
    Dim hookKind As String
    Select Case True
        Case InStr(targetSheet, "资产负债") > 0: hookKind = "BalanceSheet"
        Case InStr(targetSheet, "利润") > 0:     hookKind = "Income"
        Case InStr(targetSheet, "现金流") > 0:   hookKind = "CashFlow"
        Case Else:                               hookKind = ""
    End Select

    WriteWideTable wsTarget, arrCodes, dictCompanyName, dictData, _
                    arrPeriods, arrIndicators, dictCategoryMap, _
                    perCompanyPeriods:=False, _
                    dictReportingCurrency:=dictReportingCurrency, _
                    statementKind:=hookKind
```

### 子任务 3D — 港股调用点改造(本期最复杂,**reporting currency 必须 per-company**)

**3D.1** 在 `模块_抓港股财报.bas` line 236 把 `currencyText` 收集到一个模块级或 Sub 级 dict,key=ticker,value=currency code(`"CNY"` / `"HKD"` / `"USD"`):

```vba
    ' 现有: Dim currencyText As String: currencyText = HK_NzStr(dataRoot, "currency")
    ' Phase 4f Step 3: 收集到主流程的 dictReportingCurrency
    If Len(currencyText) > 0 And currencyText <> "REPORT_CURRENCY" Then
        ' 标准化常见写法: "CNY"/"RMB"/"人民币" → "RMB"; "HKD"/"港币" → "HKD"; "USD"/"美元" → "USD"
        Dim normCur As String
        Select Case UCase$(currencyText)
            Case "CNY", "RMB":              normCur = "RMB"
            Case "HKD":                     normCur = "HKD"
            Case "USD":                     normCur = "USD"
            Case Else:                      normCur = currencyText  ' 未知保留原值
        End Select
        ' 把 normCur 塞进调用方传下来的 dictReportingCurrency(strTicker)
        ' 注意 strTicker 不带 .HK 后缀
    End If
```

**3D.2** 港股主流程(`模块_抓港股财报.bas` 主 Sub,`WriteWideTable` 调用点 line 146)把 dict 收集起来传下去:

```vba
    Dim dictReportingCurrency As Object: Set dictReportingCurrency = CreateObject("Scripting.Dictionary")
    ' 在 fetch 循环里, 当 FetchHKFromXueqiu 拿到 currencyText 时, 写入 dictReportingCurrency(strCode)
    ' (建议 FetchHKFromXueqiu 增加 ByRef dictReportingCurrency 参数, 由它直接写)

    ' fallback: 如果某家港股没拿到 currency, 默认 "HKD"
    Dim hkCode As Variant
    For Each hkCode In arrCodes
        If Not dictReportingCurrency.Exists(CStr(hkCode)) Then _
            dictReportingCurrency(CStr(hkCode)) = "HKD"
    Next hkCode

    WriteWideTable wsTarget, arrCodes, dictCompanyName, dictData, _
                    arrPeriods, arrIndicators, dictCategoryMap, _
                    perCompanyPeriods:=False, _
                    dictReportingCurrency:=dictReportingCurrency, _
                    statementKind:=hookKind
```

**3D.3 重要数据点(plan reviewer 已 cross-check)**:

| 港股 ticker | quote 显示 | reporting currency(雪球 finance API `data.currency`) |
|---|---|---|
| 00700 腾讯 | HKD 挂牌 | RMB |
| 09988 阿里 | HKD 挂牌 | RMB |
| 02519 傲基股份 | HKD 挂牌 | RMB(大陆公司,跨境电商) |
| 09618 京东 | HKD 挂牌 | RMB |
| 00005 汇丰 | HKD 挂牌 | USD |
| 02318 中国平安 | HKD 挂牌 | RMB |

→ 实测说明:**绝大多数港股大陆公司用 RMB 报告**,本地港股公司才用 HKD。`fallback "HKD"` 是 conservative,实际大多数会被 normCur 覆盖。

### 子任务 3E — 韩股调用点改造

**3E.1** `模块_抓韩股财报.bas` line 148 同 3C 模式,字典全填 `"KRW"`,statementKind 同样从 `targetSheet` 反推。

### 验证(Generator 自测,不需要联网)

- VBA 编译通过(在 Excel 里 Debug → Compile)
- A 股回归: 样本池 B6 = `"原币"`,跑 `一键A股`,数值跟 baseline 字节级一致(WriteWideTable 行为不变,因为 displayMode 短路)
- A 股 RMB toggle: B6 = `"统一RMB"`,跑 `一键A股`,数值仍跟 baseline 一致(因为 A 股 reporting = RMB,GetFxRate("RMB",*,*) = 1.0)

### Generator 不要做

- ❌ 不要改 `模块_抓汇率.bas`(Step 2 已 frozen)
- ❌ 不要改 `汇率` sheet 模板布局
- ❌ 不要给 `Indicator` 表加 hook(指标表是 Excel 公式自动算,自然继承换算后的 BS/IS/CF 数值)

---

## Step 4 — 诊断 sheet 加 11 列 `FX_Rate`

### 目标

诊断 sheet 现在 10 列(`公司 / 报表 / 输出指标 / 状态 / 数据源 / Taxonomy / 命中字段 / Unit / Score / 匹配方式+备注`),加第 11 列 `FX_Rate`,显示**写每个数据 cell 时使用的换算汇率**(原币模式 = 1.0,统一RMB 模式 = 期末或期均汇率)。

### 子任务 4A — `EnsureDiagnosticSheet` headers 改 11 列

**4A.1** 在 `模块_工具函数.bas` line 138 `EnsureDiagnosticSheet`:

```vba
    ' line 150 现有:
    headers = Array("公司", "报表", "输出指标", "状态", "数据源", _
                    "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注")
    ' Phase 4f Step 4: 加 FX_Rate 第 11 列
    headers = Array("公司", "报表", "输出指标", "状态", "数据源", _
                    "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注", "FX_Rate")
```

**4A.2** 同 Sub 内所有 `Cells(1, ..., 10)` / `Cells(2, ..., 10)` / `ColumnWidth("J")` 等硬编码 10 改为 11,新增 `Columns("K").ColumnWidth = 12`,`SetBorderLine` 范围扩展到 11 列。

### 子任务 4B — `AddDiagnosticRow` 签名加 fxRate 参数

**4B.1** `模块_工具函数.bas` line 226:

```vba
Public Sub AddDiagnosticRow(ByVal collRows As Collection, _
                            ByVal ticker As String, _
                            ByVal strKind As String, _
                            ByVal label As String, _
                            ByVal statusText As String, _
                            ByVal sourceText As String, _
                            ByVal taxonomyText As String, _
                            ByVal fieldText As String, _
                            ByVal unitText As String, _
                            ByVal scoreText As String, _
                            ByVal noteText As String, _
                            Optional ByVal fxRateText As String = "1.0")
    If collRows Is Nothing Then Exit Sub
    collRows.Add Array(ticker, strKind, label, statusText, sourceText, _
                       taxonomyText, fieldText, unitText, scoreText, noteText, fxRateText)
End Sub
```

(用 Optional 默认 `"1.0"` 的好处:旧调用 0-arg 兼容,显示原币时永远 1.0)

### 子任务 4C — `WriteDiagnosticForKind` arrOut sizing 10 → 11

**4C.1** `模块_工具函数.bas` line 243-265 把 `1 To 10` 改 `1 To 11`,`For j = 0 To 9` 改 `0 To 10`。

### 子任务 4D — 关键 hot-path callers 传 FX_Rate

**4D.1** 美股诊断写入(`模块_抓美股*.bas` 系列)、港股诊断写入、韩股诊断写入 — 在 `AddDiagnosticRow` 调用点加上第 11 个参数:

```vba
    Dim displayMode As String: displayMode = ReadDisplayCurrency()
    Dim fxText As String: fxText = "1.0"
    If displayMode = "统一RMB" Then
        Dim useEopFlag As Boolean: useEopFlag = (strKind = "BalanceSheet")
        Dim fxNum As Double: fxNum = GetFxRate(reportingCurrency, periodEnd, useEopFlag)
        If fxNum > 0 Then fxText = Format$(fxNum, "0.000000")
    End If
    AddDiagnosticRow collRows, ticker, strKind, label, statusText, sourceText, _
                     taxonomyText, fieldText, unitText, scoreText, noteText, fxText
```

(A 股诊断写入因为 reporting currency 一直 RMB,fxText 永远 1.0,可不改 — 但同步写一下更整齐)

### 子任务 4E — `tools/install_modules.py` 兼容老 xlsm

**4E.1** 在 `_install_buttons` 之后或 `main()` 末尾加一段 idempotent migration: 如果发现现有诊断 sheet 只有 10 列表头,自动加第 11 列。或者更简单:`EnsureDiagnosticSheet` 已经是 idempotent 的(每次跑都 rewrite 表头),所以下次跑数自动升级即可,不需要 install 时改。

### 验证

- 在 `统一RMB` 模式跑 `一键韩股`(只 1 家 Zinus,样本量小好看),`韩股_抓取诊断` sheet K 列出现 `0.005280` 这样的数值
- 在 `原币` 模式跑同一份,K 列全 `1.0`
- 旧 xlsm 重装后,首次跑 → K 列出现; 老的 10 列数据被覆盖

---

## Step 5 — A1 注释动态化 + 各表 R1 公司名加币种 tag

### 目标

让用户**一眼能看出**当前表是原币还是 RMB:

- **A1 注释**: 显示当前显示模式 + 单位说明 + 如何切换
- **R1 公司名**(可选): 在 `安克(300866)` 后面加 ` [RMB]` 之类的小 tag(只在 `统一RMB` 模式下加)— 给港股最有价值,因为港股每家币种不同

### 子任务 5A — 各 Run* 主流程末尾刷 A1 注释

**5A.1** 在 4 个市场 Main 的 `WriteWideTable` 之后追加:

```vba
    Dim displayMode As String: displayMode = ReadDisplayCurrency()
    On Error Resume Next
    If Not wsTarget.Range("A1").Comment Is Nothing Then wsTarget.Range("A1").Comment.Delete
    On Error GoTo 0

    Dim commentText As String
    If displayMode = "统一RMB" Then
        commentText = "单位: 百万 RMB (统一汇率换算; 汇率源见『汇率』sheet, 期末/期间均值混合)" & vbCrLf & _
                      "切回原币: 样本池 B6 改为 '原币' 后重跑"
    Else
        commentText = "单位: " & UnitDescriptionForMarket(targetSheet)    ' 见 5A.2
        commentText = commentText & vbCrLf & "统一显示 RMB: 样本池 B6 改为 '统一RMB' 后重跑"
    End If
    wsTarget.Range("A1").AddComment commentText
    wsTarget.Range("A1").Comment.Shape.TextFrame.AutoSize = True
```

**5A.2** 新增 helper `UnitDescriptionForMarket(targetSheet)` 在 `模块_工具函数.bas`:

```vba
Public Function UnitDescriptionForMarket(ByVal sheetName As String) As String
    Select Case True
        Case InStr(sheetName, "A股") > 0:    UnitDescriptionForMarket = "百万 RMB (新浪财报源)"
        Case InStr(sheetName, "美股") > 0:   UnitDescriptionForMarket = "百万 USD (EDGAR)"
        Case InStr(sheetName, "港股") > 0:   UnitDescriptionForMarket = "百万 (各家公司报告币种, 见 港股_抓取诊断 Unit/FX_Rate 列)"
        Case InStr(sheetName, "韩股") > 0:   UnitDescriptionForMarket = "十亿 KRW (stockanalysis)"
        Case Else:                           UnitDescriptionForMarket = "(单位见诊断 sheet)"
    End Select
End Function
```

### 子任务 5B(可选,默认做)— R1 公司名加币种 tag

**5B.1** 在 `WriteWideTable` line 1135 写 `strName = CStr(...) & "(" & strCode & ")"` 之后追加:

```vba
            ' Phase 4f Step 5: 统一RMB 模式下追加币种 tag, 提示用户原始币种
            If displayMode = "统一RMB" And Not dictReportingCurrency Is Nothing Then
                If dictReportingCurrency.Exists(strCode) Then
                    Dim origCur As String: origCur = CStr(dictReportingCurrency(strCode))
                    If origCur <> "RMB" And Len(origCur) > 0 Then
                        strName = strName & " [" & origCur & "→RMB]"
                    End If
                End If
            End If
```

### 验证

- B6 = `"统一RMB"`,跑 `一键全抓`:
  - A 股 sheet A1 注释 = "单位: 百万 RMB (统一汇率换算...)"
  - 美股 sheet R1 = `Apple(AAPL) [USD→RMB]`
  - 港股 sheet R1: `腾讯(00700) [RMB→RMB]`(其实就是无 tag, 因为 origCur=RMB)、`汇丰(00005) [USD→RMB]`、`傲基(02519)`(无 tag)
  - 韩股 sheet R1 = `Zinus(013890) [KRW→RMB]`
- B6 = `"原币"`: 全部 R1 没 tag,A1 注释回到原币说明

---

## Step 6 — 联网回归测试 + STATUS.md §U 收口

### 目标

复用 Step 2 已经写好的回归驱动 `tools/test_fx_live.py`(已联网验证 5/5),再加一份端到端 4 市场抓数 + RMB toggle diff 测试,把整个 Phase 4f 闭环。

### 子任务 6A — Step 2 回归(确认 Step 3-5 没把 Step 2 弄坏)

**6A.1** 直接跑现有 driver:

```bash
cd "VBA Captor"
py tools/test_fx_live.py
```

**期望**: 仍 5/5 通过(USD/HKD/KRW × 多个 periodEnd),`GetFxRate` 往返一致,`RMB`/`CNY` 短路返 1.0,缓存命中 0.00s。**任何一项 regress 都需要排查 Step 3 是否动了不该动的代码**。

### 子任务 6B — 4 市场端到端: B6 双状态 diff

**6B.1** 准备样本池(用户实际场景):

| 市场 | 代码列 | 简称 |
|---|---|---|
| A 股 (A:B) | 300866 / 603313 / 603008 / 301376 | 安克 / 梦百合 / 喜临门 / 致欧 |
| 港股 (I:J) | 02519 | 傲基股份 |
| 韩股 (M:N) | 013890 | Zinus |

A2 = 2024,A4 = Q4

**6B.2 第一轮**: B6 = `"原币"` → 跑 `一键全抓 4 市场` → 保存 4 张主表(BS/IS/CF/Indicator × 4 市场 = 16 张)的数据快照(用 openpyxl 读出来 dump JSON 即可)→ `samples/regression_phase4f_yuanbi.json`

**6B.3 第二轮**: B6 = `"统一RMB"` → 重跑 → 同样 dump → `samples/regression_phase4f_rmb.json`

**6B.4 自动 diff 脚本**(新建 `tools/diff_phase4f_rmb.py`):

```python
"""
Phase 4f Step 6: 比较原币 vs 统一RMB 两份 dump.
期望:
  - A 股 4 张表: 字节级一致 (因为 A 股 reporting = RMB, fx=1.0)
  - 港股 02519: 字节级一致 (因为 02519 reporting = RMB)
  - 美股 / 韩股 / 报告非 RMB 的港股: 数值差应该等于 (原币 × 缓存的 FX rate)
"""
```

输出: 每张表的 mismatch 数,以及"换算后差异 vs 期望"的最大相对误差(target < 1e-6)。

### 子任务 6C — 用户使用文档刷新

**6C.1** 在 `tools/install_modules.py` `update_intro_sheet` 的"汇率与币种"小节末尾追加(line 1016-1020 那段已有 RMB 介绍, 这次升级到带使用步骤):

```python
    "  1. 在样本池 A 列填代码、B 列填简称 (各市场分栏)",
    "  2. B5 可填雪球访客 cookie (一般留空也可, 系统会自动 warmup 拿)",
    "  3. B6 选 '原币' (默认) 或 '统一RMB' (4 市场全部按当期汇率换算成 RMB 显示)",
    "  4. 点 '一键全抓 4 市场', 等候 ~3 分钟",
    "  5. 切换 B6 后**需要重新点抓数按钮**, 数值才会重算 (本期不做实时 toggle)",
    "汇率值在『汇率』sheet 缓存; 用户可手填 cell override 系统拉取值, 备注列写理由",
```

**6C.2** README.md 同步加 Phase 4f Step 3-7 完成说明(可选,Codex 可把这段放进 commit message 里)

### 子任务 6D — `STATUS.md` 加 §U 收口

**6D.1** 模仿 §T 的格式,在 `VBA Captor/STATUS.md` 末尾加 §U:

```markdown
## U. Phase 4f Step 3-7 收口: 4 市场 RMB 换算 hook + UI 反馈 + 端到端回归

执行依据: `PHASE_4F_RMB_PLAN.md` v3 + `PHASE_4F_STEP3_7_TASKS.md`. 状态: ✅ Codex 实现 + 联网回归通过, Phase 4f 全期闭环.

### U.1 本阶段已完成
- WriteWideTable 加 dictReportingCurrency + statementKind 参数
- 4 市场 wrapper 各自构造 reportingCurrency dict (港股 per-company)
- 诊断 sheet 加 11 列 FX_Rate
- A1 注释动态化 + R1 公司名 [origCur→RMB] tag
- 端到端回归: 原币 vs RMB 双 toggle, A 股/港股大陆公司字节级一致

### U.2 验证结果
[6A 测试结果, 6B 4 市场 diff 结果]

### U.3 已知边界
- Indicator 表是 Excel 公式自动算, 不在 hook 范围 (会自动继承换算后的 BS/IS/CF)
- 切换 B6 后必须重新点抓数 (本期不做 toggle 即时刷新)
- 港股 fallback 币种 = HKD (实际几乎所有大陆公司 = RMB)
```

**6D.2** 同时把 `PHASE_4F_RMB_PLAN.md` 状态行从 `🚧 Step 2 PASS` 升级为 `✅ Phase 4f 全期闭环`。

### Generator 不要做

- ❌ 不要重写 `tools/test_fx_live.py`(已经验证可用)
- ❌ 不要在回归测试里加新的 HTTP 调用(数据已缓存)
- ❌ 不要把回归 dump JSON 提交到 git(放 `samples/regression_phase4f_*.json` 即可,加进 `.gitignore` 如果还没 ignore)

---

## Step 7 — 决策点: 是否做 #2(4 市场合表)

### 现状

Step 6 完成后,用户拿到的能力是:**4 张分市场 BS/IS/CF/Indicator 表 + B6 toggle 切原币/RMB**。这已经满足"跨市场对标 6 家公司"的核心需求 — 因为 RMB 模式下,A股_资产负债表 + 港股_资产负债表 + 韩股_资产负债表 三张表的数值已经可比(都是 RMB),用户手动写 INDIRECT/VLOOKUP 公式做 ad-hoc 合并即可。

### 决策(Planner 推荐 defer 到 Phase 4g)

**理由**:
1. 16 张分市场表已稳定,合表会破坏 4b-14a baseline + Phase 4c/4d/4e 全部回归
2. 用户 Step 6 跑完会知道 RMB 数值是否真的好用 — 实际可能更想要"原币 + 备注换算后值",而不是"全部换算成 RMB"
3. 合表设计本身需要新决策点(横向铺公司列 vs 纵向重复指标行 / 报告期对齐策略 / 跨市场指标命名差异),应该等用户用 1-2 周后反馈
4. Phase 4g 时机更合理: 等用户实际用 #5 出现"我每次还是要复制 3 张表到一个新 sheet 比"的反馈再做

### 子任务 7A — 关闭决策记录

**7A.1** 在 `PHASE_4F_RMB_PLAN.md` 末尾追加:

```markdown
## 收口: #2 合表决策

Phase 4f 全期闭环, #2 合表 defer 到 Phase 4g (等待用户实际使用反馈).

判断依据 (2026-05-XX, Step 6 完成后):
- [填写: 用户实际跑 6 家公司 RMB toggle 的体验]
- [填写: 是否出现 "需要在一张表里看 4 市场" 的明确需求]
```

**7A.2** Codex 不需要写代码,只需要把这段决策落到 PLAN.md。如果用户在 Step 6 演示后**明确说"现在就要做合表"**,再开 Phase 4g — 不在 Step 7 做。

---

## 整体顺序 + 估时

| Step | 内容 | Codex 估时 | 阻塞依赖 |
|---|---|---|---|
| 3 | WriteWideTable hook + 4 市场 wrapper | 2-3h | 无 |
| 4 | 诊断 sheet 11 列 | 1h | Step 3 已可独立做但建议合并 commit |
| 5 | A1 注释动态化 + R1 tag | 1h | 依赖 Step 3 (用 dictReportingCurrency) |
| 6 | 联网回归 + 4 市场 diff + STATUS §U | 2-3h | 依赖 Step 3-5 全部 |
| 7 | 决策点 (无代码) | 0.5h | 依赖 Step 6 验证结果 |

**Codex 工作流建议**:

```
Round 1: Step 3 (commit "Phase 4f step 3 RMB hook in WriteWideTable")
  → Planner review (跑 A 股原币 baseline diff, 确认 0 mismatch)
Round 2: Step 4 + Step 5 (commit "Phase 4f step 4-5 diagnostic FX_Rate col + A1 dynamic comment")
  → Planner review (跑韩股 RMB toggle 看 K 列 + R1 tag)
Round 3: Step 6 (commit "Phase 4f step 6 end-to-end RMB regression") + Step 7 (commit "Phase 4f closure")
  → Planner final review + 用户 sign-off
```

每轮 commit 完暂停等 Planner review,don't 一口气跑完 — Generator/Reviewer 联合套路保留。

---

## ⚠️ Generator 严禁动的东西

| 文件/区域 | 原因 |
|---|---|
| `modules/模块_抓汇率.bas` 全文 | Step 2 联网验证已锁,frozen |
| `modules/模块_工具函数.bas` line 535-820 (HTTP 抓数 + GetFxRate 区块) | Step 2 已 frozen, 任何改动可能破坏 5/5 测试 |
| `modules/模块_工具函数.bas` line 566/625/807 `Accept-Encoding: identity` | 美股 finance 端工作的关键, 不是 FX 端, 别动 |
| `tools/test_fx_live.py` | Step 2 联网验证驱动, frozen |
| 现有 4 市场抓数 fetch 函数 (`FetchHKFromXueqiu` 等) | 只在主流程那一层加 dict 收集 + 传参, **不要进 fetch 函数内部改 HTTP 行为** |
| 现有诊断 sheet 已有 10 列的所有调用方 | 只追加 Optional 第 11 个参数, 旧调用 0-arg 兼容 |

---

## ⚠️ 联系 Planner 触发条件

任何下述情形,Codex **暂停并写报告**等 Planner review,不要自行决定:

1. Step 3 改完 A 股回归 baseline diff > 0 mismatch(说明 displayMode="原币" 短路没生效)
2. 港股某家公司 `dictReportingCurrency` 拿到的 `currency` 不在 `RMB/HKD/USD` 三选一(说明雪球字段值多样性超出 plan 预期)
3. `GetFxRate` 在某个真实期末/期均值上返 0(意味着 Step 2 的 5.5 年缓存窗口不够,需要 plan 调整 count 上限)
4. Step 6 用户实际场景的 6 家公司里, RMB toggle 后某家公司的某指标 RMB 值看上去明显不合理(eg 安克总资产从 200 亿变成 200 万)
5. Step 7 用户明确说"立刻要做合表"

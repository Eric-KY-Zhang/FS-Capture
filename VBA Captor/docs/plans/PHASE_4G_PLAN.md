# Phase 4g: 跨市场合并指标表 + UX 收尾 + 备用数据源调研

> **版本**: v2(2026-05-04,Phase 4g 全期闭环)
> **状态**: ✅ Phase 4g 全期闭环
> **作者**: Claude(planner) + Codex(generator)
> **背景**: Phase 4f #5 RMB 已让 6 家公司在 3 个分市场表里数值可比,但仍要肉眼跨表对比。Phase 4g 把跨市场指标合并到 1 张表,完成"6 公司跨市场对标"工作流的最后一公里。

## 项目语境(给 Generator 的 anchor 段)

本项目是个人用户(财务/审计专业人士)用的 Excel 桌面报表工具。Phase 4g 全部工作是**本地 VBA 重构 + Excel 写表 + 1 处第三方公开网页 sample 抓样**。除了 Step 4 抓 5-6 个 sample HTML 页之外,不引入任何新的网络循环或自动化采集,Step 1/2/3/5 完全本地。

## 用户已锁定的决策(2026-05-04 AskUserQuestion 收集)

| 决策点 | 选项 |
|---|---|
| 范围 | **#2 合表 + #4 stockanalysis 调研** + UX hide-tab 按钮 + Phase 4f tail 小尾巴 |
| 老 16 张分市场表去留 | **保留** + 新增 hide-tab 按钮(切换显隐) |
| 报告期对齐 | **横向铺 公司×报告期, perCompanyPeriods=True**(跟港股/美股 sheet 一致) |
| 指标命名统一 | **复用现 18 项标准指标表, 只合 Indicator 表**(BS/IS/CF 不合, 留 Phase 4h+) |
| stockanalysis 范围 | **只调研 港股 + 中概美股 的覆盖, 不切换数据源**(切换留 Phase 4h+) |

## Step 总览

| Step | 内容 | 估时 | 阻塞依赖 |
|---|---|---|---|
| 1 | Phase 4f tail 小尾巴: install_modules 诊断 sheet 升级 self-heal | 0.3h | 无 |
| 2 | 新建『跨市场_指标表』sheet + 公式联动 4 张分市场指标表 | 2h | 无 |
| 3 | UX: hide-tab 按钮(4 市场 + 1 全局)+ POOL_DATA_START_ROW 10→11 迁移 | 2h | 无(独立) |
| 4 | stockanalysis 港股 + 中概美股 字段覆盖 sample 调研报告 | 1.5h | 无(可推后) |
| 5 | 端到端回归 + STATUS §V 收口 + plan v2 升级 | 1h | 依赖 1-4 |

**Codex 工作流建议**:Round 1 = Step 1+2(commit "Phase 4g step 1-2 cross-market indicator merge + diagnostic self-heal");Round 2 = Step 3 单独(commit "Phase 4g step 3 hide-tab buttons + pool row migration");Round 3 = Step 4+5(commit "Phase 4g step 4-5 stockanalysis coverage report + closure")。每轮 commit 后停下等 Planner review。

---

## Step 1 — Phase 4f tail 小尾巴(诊断 sheet 升级 self-heal)

### 背景

Phase 4f Step 4-5 把诊断 sheet 表头从 10 列扩到 11 列(加 `FX_Rate`),但 `tools/install_modules.py` `ensure_market_sheets` 函数只在 sheet 不存在时调 `_make_diagnostic_sheet`(11 列模板),sheet 已存在则直接 skip。结果:用户从旧 xlsm 升级到新版,装完打开看诊断 sheet 仍是 10 列(只在第一次跑 一键X股 后 VBA `EnsureDiagnosticSheet` 才把表头补成 11 列)。

### 子任务 1A — `ensure_market_sheets` 强制 rewrite 已存在诊断 sheet 的表头

**1A.1** 在 `tools/install_modules.py` line ~931-944(`for diag_name in (...)` 那段循环)修改:

```python
    for diag_name in ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"):
        if diag_name in {sh.Name for sh in wb.Sheets}:
            ws_diag = wb.Sheets(diag_name)
            # Phase 4g Step 1: 强制 rewrite 表头到 11 列, 让 Phase 4f FX_Rate 列即装即可见
            _refresh_diagnostic_headers(ws_diag)
            try:
                ws_diag.Visible = 0  # xlSheetHidden
            except Exception:
                pass
            print(f"  ~ sheet 已存在 (表头已升级到 11 列): {diag_name}")
        else:
            ws_diag = _make_diagnostic_sheet(wb, diag_name)
            try:
                ws_diag.Visible = 0
            except Exception:
                pass
            print(f"  + sheet 新建: {diag_name}")
```

**1A.2** 新增 helper `_refresh_diagnostic_headers`(放在 `_make_diagnostic_sheet` 之后):

```python
def _refresh_diagnostic_headers(ws):
    """Phase 4g Step 1: 把已存在的诊断 sheet Row 2 表头强制刷成 11 列 (含 FX_Rate)
       仅改 Row 1-2 的表头格式 + Column K 列宽, 不动 Row 3+ 数据"""
    headers = ["公司", "报表", "输出指标", "状态", "数据源",
               "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注", "FX_Rate"]
    widths = [14, 16, 30, 18, 18, 14, 42, 14, 10, 58, 12]
    for j, txt in enumerate(headers, start=1):
        c = ws.Cells(2, j)
        c.Value = txt
        c.Font.Name = "微软雅黑"
        c.Font.Size = 10
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
    for j, w in enumerate(widths, start=1):
        ws.Columns(j).ColumnWidth = w
    ws.Rows(2).RowHeight = 20
    # 标题 merge 也要跟着扩到 11 列
    try:
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).UnMerge
    except Exception:
        pass
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).Merge()
```

### 验证(Generator 自测)

- 在 worktree 跑 `py tools/install_modules.py`(对一个有旧 10 列诊断表头的 xlsm),装完手开 Excel 打开诊断 sheet,K2 = `FX_Rate`,K 列列宽 = 12

---

## Step 2 — 新建『跨市场_指标表』sheet + 公式联动 4 张分市场指标表

### 目标

新建 1 张 sheet `跨市场_指标表`(无市场前缀,区别于 4 张分市场指标表),layout 复用现 18 项标准指标 + 横向铺所有公司。每个数据 cell 是**只读公式**,直接 cell-reference 到对应分市场指标表的对应 cell。优点:用户跑数后两边自动同步;无需写新的换算逻辑(分市场指标表已经在 Phase 4f 完成 RMB 换算)。

### 子任务 2A — `tools/build_template.py` 新建跨市场指标表模板

**2A.1** 在 `build_template.py` 新增 `build_cross_market_indicator_sheet(ws)`(参考现有 `_make_wide_table_sheet` 风格):

```python
def build_cross_market_indicator_sheet(ws):
    """Phase 4g Step 2: 跨市场_指标表 模板
      Row 1-2: 静态表头容器 (A1='指标类型', B1='指标名称', C1='英文指标名')
              + R1 公司名 + R2 报告期 由 VBA BuildCrossMarketIndicatorSheet 动态写入
      Row 3+: 18 项标准指标公式由 VBA 动态填
      列宽: A=18 / B=28 / C=34 / D+=15.875"""
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 34
    fill = PatternFill("solid", fgColor=DARK_BLUE)
    for cell_addr, txt in (("A1", "指标类型"), ("B1", "指标名称"), ("C1", "英文指标名")):
        cell = ws[cell_addr]
        cell.value = txt
        cell.font = HEADER_FONT
        cell.fill = fill
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 20
    ws.freeze_panes = "D3"
```

**2A.2** 在 `main()` 末尾创建该 sheet:

```python
    ws_cross = wb.create_sheet("跨市场_指标表")
    build_cross_market_indicator_sheet(ws_cross)
```

### 子任务 2B — `tools/install_modules.py` `ensure_market_sheets` 也加跨市场指标表

**2B.1** 同 Phase 4f Step 2 的『汇率』sheet idempotent install 模式,新增:

```python
    if "跨市场_指标表" in {sh.Name for sh in wb.Sheets}:
        print("  ~ sheet 已存在: 跨市场_指标表")
    else:
        _make_cross_market_indicator_sheet(wb, "跨市场_指标表")
        print("  + sheet 新建: 跨市场_指标表")
```

加 `_make_cross_market_indicator_sheet` helper (类似 `_make_fx_sheet`)。

**2B.2** 把跨市场指标表加到 `reorder_report_sheets` 的 `desired_order` 末尾(诊断 sheet 之后,『汇率』之前)。

### 子任务 2C — 新建 VBA `Public Sub BuildCrossMarketIndicatorSheet`

**2C.1** 在 `模块_工具函数.bas` 新增(放在 `BuildStandardIndicatorSheet` 之后):

```vba
' --------- Phase 4g Step 2: 把 4 张分市场指标表合并展示到『跨市场_指标表』 ---------
'   - 复用 18 项标准指标 (StandardIndicatorDefs)
'   - 横向铺 公司×报告期, perCompanyPeriods=True (跟港股/美股 sheet 一致)
'   - 每个 cell 是 formula, 引到分市场指标表对应 cell (eg ='港股_指标表'!N5)
'   - 自动跳过空的分市场表 (eg 用户没跑韩股就只显示已有 3 市场)
Public Sub BuildCrossMarketIndicatorSheet()
    Const TARGET_SHEET As String = "跨市场_指标表"
    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(TARGET_SHEET)
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Err.Raise vbObjectError + 580, "BuildCrossMarketIndicatorSheet", _
            TARGET_SHEET & " sheet 不存在, 请重装模板"
    End If

    Application.ScreenUpdating = False
    On Error Resume Next
    wsTarget.UsedRange.UnMerge
    On Error GoTo 0
    wsTarget.Cells.Clear

    ' 1) 重画静态表头
    wsTarget.Range("A1").Value = "指标类型"
    wsTarget.Range("B1").Value = "指标名称"
    wsTarget.Range("C1").Value = "英文指标名"
    With wsTarget.Range("A1:C1")
        .Font.Name = "微软雅黑": .Font.Size = 11: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' 2) 收集 4 张分市场指标表的公司+报告期 layout (从 R1-R2 表头反推)
    '    返回结构: collCompanies(i) = Array(market, ticker, displayName, sourceSheet, sourceStartCol, periodCount)
    Dim collCompanies As Collection: Set collCompanies = New Collection
    Dim markets As Variant: markets = Array("A", "US", "HK", "KR")
    Dim m As Variant
    For Each m In markets
        Dim sourceSheet As String: sourceSheet = MarketIndicatorSheetName(CStr(m))
        If WorksheetExists(sourceSheet) Then
            CollectCompaniesFromIndicatorSheet ThisWorkbook.Sheets(sourceSheet), CStr(m), collCompanies
        End If
    Next m

    If collCompanies.Count = 0 Then
        wsTarget.Range("A3").Value = "(还没有任何分市场指标表数据, 请先点 一键X股 跑数后再合并)"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 3) 写 R1 公司名 (含市场 tag) + R2 报告期, 横向铺 perCompanyPeriods
    Dim targetCol As Long: targetCol = 4
    Dim i As Long, j As Long
    For i = 1 To collCompanies.Count
        Dim entry As Variant: entry = collCompanies(i)
        Dim mkt As String: mkt = CStr(entry(0))
        Dim ticker As String: ticker = CStr(entry(1))
        Dim displayName As String: displayName = CStr(entry(2)) & " [" & mkt & "]"
        Dim srcSheet As String: srcSheet = CStr(entry(3))
        Dim srcStartCol As Long: srcStartCol = CLng(entry(4))
        Dim periodCount As Long: periodCount = CLng(entry(5))

        Dim companyStartCol As Long: companyStartCol = targetCol
        wsTarget.Cells(1, companyStartCol).Value = displayName
        If periodCount > 1 Then
            wsTarget.Range(wsTarget.Cells(1, companyStartCol), _
                           wsTarget.Cells(1, companyStartCol + periodCount - 1)).Merge
        End If
        With wsTarget.Cells(1, companyStartCol)
            .Font.Name = "微软雅黑": .Font.Size = 11: .Font.Bold = True
            .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
        End With

        For j = 1 To periodCount
            ' R2 报告期 = 直接 cell-ref 到 source sheet 的 R2 同列
            wsTarget.Cells(2, targetCol).Formula = "='" & srcSheet & "'!" & _
                ThisWorkbook.Sheets(srcSheet).Cells(2, srcStartCol + j - 1).Address(False, False)
            wsTarget.Cells(2, targetCol).NumberFormat = "yyyy-mm-dd"
            With wsTarget.Cells(2, targetCol)
                .Font.Name = "微软雅黑": .Font.Size = 10: .Font.Bold = True
                .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(68, 114, 196)
                .HorizontalAlignment = xlCenter
            End With
            targetCol = targetCol + 1
        Next j
    Next i

    Dim lastCol As Long: lastCol = targetCol - 1

    ' 4) 写 18 标准指标行 (Row 3-20), 每个 cell = source sheet 的对应 cell ref
    Dim defs As Variant: defs = StandardIndicatorDefs()
    Dim stdCount As Long: stdCount = UBound(defs) - LBound(defs) + 1

    Dim k As Long
    For k = 0 To stdCount - 1
        Dim rowNum As Long: rowNum = 3 + k
        wsTarget.Cells(rowNum, 1).Value = CStr(defs(k)(0))
        wsTarget.Cells(rowNum, 2).Value = CStr(defs(k)(1))
        wsTarget.Cells(rowNum, 3).Value = CStr(defs(k)(2))

        ' 重新 walk collCompanies 写公式
        Dim writeCol As Long: writeCol = 4
        For i = 1 To collCompanies.Count
            entry = collCompanies(i)
            srcSheet = CStr(entry(3))
            srcStartCol = CLng(entry(4))
            periodCount = CLng(entry(5))
            ' Source 表的 18 标准指标固定在 Row 3-20 (BuildStandardIndicatorSheet 保证)
            Dim srcRow As Long: srcRow = 3 + k
            For j = 1 To periodCount
                wsTarget.Cells(rowNum, writeCol).Formula = "='" & srcSheet & "'!" & _
                    ThisWorkbook.Sheets(srcSheet).Cells(srcRow, srcStartCol + j - 1).Address(False, False)
                wsTarget.Cells(rowNum, writeCol).NumberFormat = CStr(defs(k)(4))
                writeCol = writeCol + 1
            Next j
        Next i
    Next k

    ' 5) 列宽 / 字体 / 边框 / 冻结
    wsTarget.Columns("A").ColumnWidth = 18
    wsTarget.Columns("B").ColumnWidth = 28
    wsTarget.Columns("C").ColumnWidth = 34
    Dim col As Long
    For col = 4 To lastCol
        wsTarget.Columns(col).ColumnWidth = 15.875
    Next col
    wsTarget.Rows(1).RowHeight = 22
    wsTarget.Rows(2).RowHeight = 20

    Call SetBorderLine(wsTarget.Range(wsTarget.Cells(1, 1), _
                                       wsTarget.Cells(2 + stdCount, lastCol)))

    With wsTarget.Range(wsTarget.Cells(3, 1), wsTarget.Cells(2 + stdCount, 3))
        .Font.Bold = True
    End With

    ' 冻结
    wsTarget.Activate
    ActiveWindow.FreezePanes = False
    wsTarget.Cells(3, 4).Select
    ActiveWindow.FreezePanes = True
    wsTarget.Cells(1, 1).Select

    ' A1 注释 (动态, 显示当前显示模式 + 公司数 / 市场数)
    On Error Resume Next
    If Not wsTarget.Range("A1").Comment Is Nothing Then wsTarget.Range("A1").Comment.Delete
    On Error GoTo 0
    Dim displayMode As String: displayMode = ReadDisplayCurrency()
    Dim commentText As String
    commentText = "跨市场指标合表 (公司数=" & collCompanies.Count & ")" & vbCrLf & _
                  "数据源: 4 张分市场指标表 (引用公式, 自动同步)" & vbCrLf & _
                  "当前显示模式: " & displayMode & vbCrLf & _
                  "切换 B6 后请先重跑各市场, 再点 '一键合并跨市场指标表'"
    wsTarget.Range("A1").AddComment commentText
    wsTarget.Range("A1").Comment.Shape.TextFrame.AutoSize = True

    Application.ScreenUpdating = True
End Sub


Private Function MarketIndicatorSheetName(ByVal market As String) As String
    Select Case UCase$(Trim$(market))
        Case "A":  MarketIndicatorSheetName = "A股_指标表"
        Case "US": MarketIndicatorSheetName = "美股_指标表"
        Case "HK": MarketIndicatorSheetName = "港股_指标表"
        Case "KR": MarketIndicatorSheetName = "韩股_指标表"
    End Select
End Function


Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(sheetName)
    WorksheetExists = (Err.Number = 0 And Not ws Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function


' --------- 从分市场指标表 Row 1 (合并的公司名 cell) 反推 公司起始列 + 报告期数 ---------
'   sourceSheet 已知 layout: 
'     Row 1: 公司名 (合并的 N 列, N = periodCount)
'     Row 2: 报告期 (每个公司 N 个报告期)
'     Row 3-20: 18 标准指标
'     Col 1-3: 静态 metaCols
'     Col 4+: 数据列
'   返回到 outColl, 每个 entry = Array(market, ticker, displayName, sourceSheet, sourceStartCol, periodCount)
Private Sub CollectCompaniesFromIndicatorSheet(ByVal ws As Worksheet, _
                                                ByVal market As String, _
                                                ByRef outColl As Collection)
    Dim startCol As Long: startCol = StandardDataStartCol(ws)    ' = 4
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < startCol Then Exit Sub

    Dim col As Long: col = startCol
    Do While col <= lastCol
        Dim companyHeader As String: companyHeader = StandardHeaderTextAt(ws, col)
        If Len(Trim$(companyHeader)) = 0 Then
            col = col + 1
        Else
            ' 找这家公司占了多少列 (R1 合并 / 同 R1 文本连续)
            Dim spanEnd As Long: spanEnd = col
            Do While spanEnd <= lastCol
                If StandardHeaderTextAt(ws, spanEnd) = companyHeader Then
                    spanEnd = spanEnd + 1
                Else
                    Exit Do
                End If
            Loop
            Dim periodCount As Long: periodCount = spanEnd - col

            ' 从 companyHeader 反解 ticker (假设格式 "简称(代码)" 或 "简称(代码)" 或 "代码")
            Dim ticker As String: ticker = ExtractTickerFromHeader(companyHeader)

            outColl.Add Array(market, ticker, companyHeader, ws.Name, col, periodCount)
            col = spanEnd
        End If
    Loop
End Sub


Private Function ExtractTickerFromHeader(ByVal header As String) As String
    ' "安克创新(300866)" -> "300866"; "Apple(AAPL)" -> "AAPL"; "300866" -> "300866"
    Dim p1 As Long, p2 As Long
    p1 = InStrRev(header, "(")
    p2 = InStrRev(header, ")")
    If p1 > 0 And p2 > p1 Then
        ExtractTickerFromHeader = Mid$(header, p1 + 1, p2 - p1 - 1)
    Else
        ExtractTickerFromHeader = Trim$(header)
    End If
End Function
```

**容错原则**:
- 任何分市场指标表不存在 / 为空 / 没数据列 → 静默跳过该市场,合表只展示已有市场
- 公司名解析失败(无括号)→ 整段当 ticker
- 全部 4 市场都空 → 写 A3 提示用户先跑数

### 子任务 2D — 在『一键全抓 4 市场』末尾追加 `BuildCrossMarketIndicatorSheet` 调用

**2D.1** 在 `模块_总入口.bas` `一键全抓` Sub 末尾(4 市场抓数都完成后)追加:

```vba
    ' Phase 4g Step 2: 一键全抓后自动刷新跨市场指标表
    On Error Resume Next
    BuildCrossMarketIndicatorSheet
    Err.Clear
    On Error GoTo 0
```

### 子任务 2E — 单独按钮『一键合并跨市场指标表』

**2E.1** 在 `tools/install_modules.py` BUTTONS list 增加 1 项(放在 Q4 一键全抓之后,row 30 单表按钮之前):

```python
    ("BtnBuildCrossInd",  "合并跨市场指标表",  "模块_工具函数.BuildCrossMarketIndicatorSheet", "Q5:Q7", PRIMARY_FILL,  PRIMARY_FG,  12,  True),
```

让用户在不一键全抓的情况下也能单独刷合表。

### 验证

- 跑 `py tools/install_modules.py` → 装完打开 xlsm,『跨市场_指标表』sheet 存在,Tab 顺序在诊断之前
- 手动给样本池填 1 家 A 股(eg 300866 安克),跑 `一键A股` → A股_指标表 写入 18 行
- 点『合并跨市场指标表』按钮 → 跨市场_指标表 R1 = "安克创新(300866) [A]",18 行公式 cell-ref 到 A股_指标表
- 再加 1 家美股(eg AAPL)跑 `一键美股`,再点合并 → 跨市场_指标表 增加 4 列(假设 AAPL 4 个报告期),R1 = "Apple(AAPL) [US]"

### Generator 不要做

- ❌ 不要改 `BuildStandardIndicatorSheet` 主体(分市场表是 source of truth,合表只是它的 view)
- ❌ 不要新增换算逻辑(分市场指标表已经在 Phase 4f 完成 RMB 换算)
- ❌ 不要把 BS/IS/CF 也合掉(本期只合 Indicator,BS/IS/CF defer Phase 4h)
- ❌ 不要动『跨市场_指标表』里的公司排序逻辑(默认按 markets = ["A", "US", "HK", "KR"] 顺序排)

---

## Step 3 — UX hide-tab 按钮 + POOL_DATA_START_ROW 10→11 迁移

### 目标

样本池 Row 8 的 `一键X股` 按钮下方增加切换按钮,**点一次隐藏本市场所有 sheet,再点一次显示回来**(toggle 语义)。`一键全抓 4 市场`(Q1:Q3)下方增加 1 个全局按钮,toggle 4 市场所有 sheet 显隐。

### 子任务 3A — POOL_DATA_START_ROW 10 → 11 迁移

**3A.1** `模块_工具函数.bas` line 27:

```vba
Public Const POOL_DATA_START_ROW As Long = 11
```

**3A.2** Python 装表脚本里所有写死的 row 10 改成 11:

- `tools/build_template.py` line 232 / 236 `for row in range(10, 31):` → `range(11, 32):`
- `tools/install_modules.py` line 486 `for row in range(10, 51):` → `range(11, 52):`
- `tools/install_modules.py` `migrate_old_sample_pool` 里 `for row in range(8, last_row + 1)` 不动(读旧数据用,旧版 row 8 起)
- `tools/install_modules.py` `migrate_old_sample_pool` 写新位置时把 `for idx, ... start=10` → `start=11`

**3A.3** `tools/install_modules.py` `layout_sample_pool`:

- 原来 row 8 = 一键按钮,row 9 = 列头`代码/简称`,row 10+ = 数据
- 改成 row 8 = 一键按钮,row 9 = hide-tab 按钮(新),row 10 = 列头`代码/简称`,row 11+ = 数据

### 子任务 3B — 4 个市场 hide-tab 按钮(Row 9 各市场列对)

**3B.1** `tools/install_modules.py` BUTTONS list 增加 4 项(放在 4 个一键X股按钮之后,row 30 单表按钮之前):

```python
    ("BtnHideA",  "切换 A 股 tabs 显隐",   "模块_总入口.切换A股tabs",  "A9:B9", SECONDARY_FILL, SECONDARY_FG, 9, False),
    ("BtnHideUS", "切换 美股 tabs 显隐",   "模块_总入口.切换美股tabs", "E9:F9", SECONDARY_FILL, SECONDARY_FG, 9, False),
    ("BtnHideHK", "切换 港股 tabs 显隐",   "模块_总入口.切换港股tabs", "I9:J9", SECONDARY_FILL, SECONDARY_FG, 9, False),
    ("BtnHideKR", "切换 韩股 tabs 显隐",   "模块_总入口.切换韩股tabs", "M9:N9", SECONDARY_FILL, SECONDARY_FG, 9, False),
```

(用 SECONDARY_FILL 浅蓝 + 字体小一号 9pt,跟一键 X 股按钮做视觉区分)

**3B.2** `layout_sample_pool` 中:

- 设 row 9 RowHeight = 22
- 列头`代码/简称`的写入位置从 `Range("A9").Value = ...` 改成 `Range("A10").Value = ...`(同步 4 个市场 E10/I10/M10/F10/J10/N10)
- 数据区 NumberFormat = "@" 的位置从 `Range("A10:A1000")` 改成 `Range("A11:A1000")`(同步 4 个市场)
- 冻结窗格的 `SplitRow = 9` 改成 `SplitRow = 10`

### 子任务 3C — 1 个全局 hide-tab 按钮(Q5)

**3C.1** BUTTONS list 增加:

```python
    ("BtnHideAll", "切换所有分市场 tabs 显隐", "模块_总入口.切换所有分市场tabs", "Q5:Q7", SECONDARY_FILL, SECONDARY_FG, 11, True),
```

### 子任务 3D — VBA toggle Sub(放在 `模块_总入口.bas`)

**3D.1** 新增 5 个 Sub:

```vba
Public Sub 切换A股tabs()
    ToggleMarketTabsVisibility "A"
End Sub
Public Sub 切换美股tabs()
    ToggleMarketTabsVisibility "US"
End Sub
Public Sub 切换港股tabs()
    ToggleMarketTabsVisibility "HK"
End Sub
Public Sub 切换韩股tabs()
    ToggleMarketTabsVisibility "KR"
End Sub
Public Sub 切换所有分市场tabs()
    Dim m As Variant
    For Each m In Array("A", "US", "HK", "KR")
        ToggleMarketTabsVisibility CStr(m)
    Next m
End Sub

Private Sub ToggleMarketTabsVisibility(ByVal market As String)
    Dim prefix As String
    Select Case UCase$(Trim$(market))
        Case "A":  prefix = "A股_"
        Case "US": prefix = "美股_"
        Case "HK": prefix = "港股_"
        Case "KR": prefix = "韩股_"
        Case Else: Exit Sub
    End Select

    ' 找到第一张匹配 sheet, 看当前可见状态; 然后全部反向 toggle
    Dim newVisible As Long: newVisible = -1    ' 默认显示
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Left$(ws.Name, Len(prefix)) = prefix Then
            If ws.Visible = -1 Then newVisible = 0    ' xlSheetHidden
            Exit For
        End If
    Next ws

    On Error Resume Next
    For Each ws In ThisWorkbook.Sheets
        If Left$(ws.Name, Len(prefix)) = prefix Then
            ws.Visible = newVisible
        End If
    Next ws
    Err.Clear
    On Error GoTo 0
End Sub
```

**注意**:港股_/美股_/韩股_ 各市场都有 `_抓取诊断` sheet,前缀匹配会把诊断也 toggle 进去 — 这是用户想要的(隐藏本市场全部相关 sheet)。

### 验证

- 装完打开 xlsm,样本池 row 9 看到 4 个浅蓝 hide-tab 按钮
- 点『切换 A 股 tabs 显隐』,A股_资产负债表/A股_利润表/A股_现金流量表/A股_指标表 全部 hidden
- 再点一次,全部回显
- Q5 全局按钮 toggle 4 市场所有 16 张表

### 风险点

- POOL_DATA_START_ROW migration 是 invasive change,Generator 必须 grep 所有 row 10 用法,逐个核对(包括 `Cells(10, ...)` 这种)
- 老用户 xlsm 升级:`migrate_old_sample_pool` 已经处理迁移,但要确认它读的是 Row 8+(旧布局),写到 Row 11+(新布局)
- 切换 sheet visibility 时如果切到当前 ActiveSheet 会报错,务必 `On Error Resume Next` 包

### Generator 不要做

- ❌ 不要把 visibility 改成 `xlSheetVeryHidden`(用户右键看不到了)
- ❌ 不要在 hide 时弹 MsgBox(toggle 是无声操作)
- ❌ 不要 hide 跨市场_指标表 / 汇率 / 样本池 / 使用说明 这 4 张共享 sheet

---

## Step 4 — stockanalysis 港股 + 中概美股 字段覆盖调研

### 背景

`stockanalysis.com` 是本项目 Phase 4d 已经在用的第三方数据源(韩股 KRX 路径)。Step 4 工作 = **手动检查同一站点是否覆盖港股 + 中概美股**,把 5-6 个 sample HTML 页保存到本地供离线分析,写一份覆盖度报告。**不切换数据源,不构建抓数循环**,只为 Phase 4h+ 提供决策依据(雪球 cookie 过期场景的 fallback 备选)。

### 子任务 4A — 写 1 个一次性 sample 收集脚本

**4A.1** 新建 `tools/probe_4a_stockanalysis_coverage.py`,完成以下手动操作:

- 检查 6 个 URL 的可达性 + 把页面 HTML 落盘到 `samples/stockanalysis_<ticker>.html`(每页 ~50KB,只跑 1 次)
- 6 个 URL 候选:
  - 港股: `https://stockanalysis.com/quote/hkg/00700/financials/`(腾讯)
  - 港股: `https://stockanalysis.com/quote/hkg/02519/financials/`(傲基)
  - 港股: `https://stockanalysis.com/quote/hkg/09988/financials/`(阿里港股)
  - 中概美股: `https://stockanalysis.com/stocks/baba/financials/`(阿里美股)
  - 中概美股: `https://stockanalysis.com/stocks/jd/financials/`(京东)
  - 中概美股: `https://stockanalysis.com/stocks/pdd/financials/`(拼多多)

- HTTP 客户端用 `requests`(项目已有),`User-Agent` 同 Phase 4d KR 路径(`Mozilla/5.0 ... Chrome/...`),不用 cookie
- 单线程顺序拉,**间隔 2 秒**(对站点友好;6 页总耗时 ~15s,远低于 Phase 4d KR 韩股已用的频率)
- **不要写循环 / 不要做批量化**;6 个 URL 写死在 list,跑完即止

**4A.2** 脚本输出格式:

- 每个 URL 跑完打印 `<ticker>: HTTP <code>, <size> bytes, saved to <path>`
- 全部跑完后总结:覆盖率(几个 200,几个 404 / 别的)

### 子任务 4B — 写覆盖度调研报告 `samples/STOCKANALYSIS_PROBE.md`

**4B.1** 报告内容(Generator 手动看 6 个 sample HTML 后填):

- 表 1: 6 个 ticker 的 HTTP 响应状态 + 页面是否含完整财报数据
- 表 2: stockanalysis 港股财报字段 vs 雪球 finance API 字段映射(挑 5 个核心指标 eg `Total Revenue` / `Total Assets`)
- 表 3: stockanalysis 中概美股 vs EDGAR + 雪球 fallback 的覆盖度对照
- 结论:Phase 4h 是否值得切换;切换的实施成本估算

### Generator 不要做

- ❌ 不要把 stockanalysis 接到 VBA(Phase 4g 只是调研,不动 VBA)
- ❌ 不要写循环抓任何超出 6 个 ticker 的页面
- ❌ 不要把 sample HTML 提交到 git(加 `.gitignore` 规则 `samples/stockanalysis_*.html` 如果还没)
- ❌ 不要重写 Phase 4d 韩股 stockanalysis 解析代码(那是 frozen 的)

---

## Step 5 — 端到端回归 + STATUS §V 收口 + plan v2 升级

### 子任务 5A — 跑现有回归驱动

```bash
cd "VBA Captor"
py tools/test_fx_live.py --skip-install     # Phase 4f Step 2 frozen 回归 (5/5)
py -u tools/diff_phase4f_step3_lite.py      # Phase 4f Step 3-5 frozen 回归 (smoke)
```

期望:**两份都 PASS**,任何一项 regress 都需排查 Step 1-3 是否动了不该动的代码。

### 子任务 5B — 跨市场指标表手动验收(无网络)

**5B.1** 用 `tools/inspect_phase4f_state.py`(Reviewer 已写)的相同模式,新增 `tools/inspect_phase4g_state.py`,dump:

- 跨市场_指标表 sheet:R1 公司名 + R2 报告期 + Row 3-5 头 3 个指标行的公式 + 计算值
- 4 个 hide-tab 按钮位置 + caption
- 1 个全局 hide-tab 按钮位置
- 所有分市场 sheet 的当前 visibility 状态(检查 toggle 动作)
- 诊断 sheet 表头是否 11 列含 `FX_Rate`(Step 1 验证)

### 子任务 5C — `STATUS.md` §V 收口

**5C.1** 模仿 §U 格式追加 §V:

```markdown
## V. Phase 4g 收口: 跨市场合并指标表 + UX hide-tab + 备用数据源调研

执行依据: `PHASE_4G_PLAN.md` v1。状态: ✅ Codex 已实现并通过本地回归 + 无网络验收;Phase 4g 全期闭环。

### V.1 本阶段已完成
- [Step 1] install_modules ensure_market_sheets 强制 rewrite 已存在诊断 sheet 表头 → 升级即可见 11 列
- [Step 2] 新建『跨市场_指标表』+ VBA BuildCrossMarketIndicatorSheet:18 标准指标 × 横向铺公司×报告期 perCompanyPeriods=True;每个数据 cell 是公式 cell-ref 到 4 张分市场指标表;一键全抓末尾自动刷
- [Step 3] hide-tab 按钮 5 个 (4 市场 + 1 全局), POOL_DATA_START_ROW 10 → 11 迁移
- [Step 4] stockanalysis 港股 + 中概美股 6 ticker 覆盖度调研报告 → samples/STOCKANALYSIS_PROBE.md;**不切换**

### V.2 验证结果
[5A 回归结果, 5B 跨市场指标表 inspect 结果]

### V.3 已知边界
- 跨市场指标表只合 18 项标准指标 (Indicator);BS/IS/CF 全合 defer 到 Phase 4h
- stockanalysis 切换 defer 到 Phase 4h(等雪球 cookie 失效或反爬触发再切)
- POOL_DATA_START_ROW 迁移 = invasive,老用户从 4f 升级时旧样本池数据自动迁移到 row 11+
```

**5C.2** `PHASE_4G_PLAN.md` 状态行 v1 → v2,标记 `✅ Phase 4g 全期闭环`。

### 子任务 5D — `README.md` 跨市场表使用说明

**5D.1** README.md「单元格说明」小节追加:

```markdown
- 跨市场_指标表: 4 张分市场指标表的合并视图,横向铺公司×报告期,18 项标准指标。每次跑数后点『合并跨市场指标表』按钮(或一键全抓后自动刷)
```

---

## ⚠️ 全 Phase 严禁动的东西

| 文件/区域 | 原因 |
|---|---|
| `modules/模块_抓汇率.bas` | Phase 4f Step 2 frozen |
| `modules/模块_工具函数.bas` line 535-820(HTTP + GetFxRate 区) | Phase 4f frozen |
| `modules/模块_工具函数.bas` `WriteWideTable` 主体逻辑 | Phase 4f Step 3-5 frozen |
| `modules/模块_工具函数.bas` `BuildStandardIndicatorSheet` 主体 | Phase 4b-8 / 4f frozen,跨市场表只 lookup 它,不改它 |
| `tools/test_fx_live.py` / `tools/diff_phase4f_step3_lite.py` | Phase 4f 验证驱动 frozen |
| 现有 4 市场 fetch 模块的 Run* + Fetch* 函数 | Phase 4c/4d frozen |
| Phase 4d 韩股 stockanalysis 解析代码 | 不动,Step 4 只调研不复用 |

## ⚠️ 联系 Planner 触发条件

- Step 2 跨市场指标表公式构造异常(eg `'港股_指标表'!N5` 这种 cell ref 因为 source sheet rebuild 时列偏移失效)
- Step 3 POOL_DATA_START_ROW 迁移后任意一市场抓数失败
- Step 4 stockanalysis 6 个 URL 有 4 个以上 404(说明覆盖太差,Phase 4h 切换不可行 — 需要重新决策)
- Step 5 回归测试任一退化

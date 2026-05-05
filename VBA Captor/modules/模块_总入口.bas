Attribute VB_Name = "模块_总入口"
Option Explicit

' =================================================================
'  上市公司财务数据查询 v1.0
'  作者: Eric Zhang  邮箱: 214978902@qq.com
'  许可: 个人 / 内部使用; 数据来源遵循各源站 Fair Use / Terms
'  开发期: 2026-05-02 ~ 2026-05-05 (vibe coding, 4 天 12 phase)
'
'  Sub 一键全抓: 顺序调用 A 股 + 美股 + 港股 + 韩股 16 张表, 静默模式, 最后弹一次汇总
'  基本资料 已废弃
' =================================================================

Public Sub 一键A股(Optional ByVal blnSilent As Boolean = False)
    Dim dtTime As Double: dtTime = Timer
    Dim runErrDesc As String
    Dim appState As TAppState
    Dim hasAppState As Boolean

    g_silentMode = True
    g_globalFails = 0
    g_globalLog = ""
    g_diagnosticAppendOnly = False
    On Error GoTo CleanUp
    appState = BeginAppState("正在抓取 A 股...")
    hasAppState = True

    Application.StatusBar = "[A股 1/4] 抓取资产负债表..."
    模块_抓资产负债表.Main
    Application.StatusBar = "[A股 2/4] 抓取利润表..."
    模块_抓利润表.Main
    Application.StatusBar = "[A股 3/4] 抓取现金流量表..."
    模块_抓现金流量表.Main
    Application.StatusBar = "[A股 4/4] 生成指标表..."
    模块_抓指标表.Main
    UnhideMarketTabs "A"

CleanUp:
    If Err.Number <> 0 Then
        runErrDesc = vbCrLf & vbCrLf & "运行中断: " & Err.Description
        Err.Clear
    End If
    g_silentMode = False
    g_diagnosticAppendOnly = False
    If hasAppState Then
        EndAppState appState
    Else
        Application.StatusBar = False
        Application.ScreenUpdating = True
    End If
    If Len(runErrDesc) > 0 Then Application.StatusBar = "一键 A 股出错: " & runErrDesc

    If Not blnSilent Then ShowMarketRunSummary "A股", dtTime, runErrDesc
End Sub


Public Sub 一键美股(Optional ByVal blnSilent As Boolean = False)
    Dim dtTime As Double: dtTime = Timer
    Dim runErrDesc As String
    Dim appState As TAppState
    Dim hasAppState As Boolean

    g_silentMode = True
    g_globalFails = 0
    g_globalLog = ""
    On Error GoTo CleanUp
    appState = BeginAppState("正在抓取 美股...")
    hasAppState = True
    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticAppendOnly = True

    Application.StatusBar = "[美股 1/4] 抓取资产负债表..."
    模块_抓美股资产负债表.Main
    Application.StatusBar = "[美股 2/4] 抓取利润表..."
    模块_抓美股利润表.Main
    Application.StatusBar = "[美股 3/4] 抓取现金流量表..."
    模块_抓美股现金流量表.Main
    Application.StatusBar = "[美股 4/4] 生成指标表..."
    模块_抓美股指标表.Main
    UnhideMarketTabs "US"

CleanUp:
    If Err.Number <> 0 Then
        runErrDesc = vbCrLf & vbCrLf & "运行中断: " & Err.Description
        Err.Clear
    End If
    g_silentMode = False
    g_diagnosticAppendOnly = False
    If hasAppState Then
        EndAppState appState
    Else
        Application.StatusBar = False
        Application.ScreenUpdating = True
    End If
    If Len(runErrDesc) > 0 Then Application.StatusBar = "一键 美股出错: " & runErrDesc

    If Not blnSilent Then ShowMarketRunSummary "美股", dtTime, runErrDesc
End Sub


Public Sub 一键港股(Optional ByVal blnSilent As Boolean = False)
    Dim dtTime As Double: dtTime = Timer
    Dim runErrDesc As String
    Dim appState As TAppState
    Dim hasAppState As Boolean

    g_silentMode = True
    g_globalFails = 0
    g_globalLog = ""
    On Error GoTo CleanUp
    appState = BeginAppState("正在抓取 港股...")
    hasAppState = True
    g_diagnosticSheetName = "港股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticAppendOnly = True

    Application.StatusBar = "[港股 1/4] 抓取资产负债表..."
    模块_抓港股资产负债表.Main
    Application.StatusBar = "[港股 2/4] 抓取利润表..."
    模块_抓港股利润表.Main
    Application.StatusBar = "[港股 3/4] 抓取现金流量表..."
    模块_抓港股现金流量表.Main
    Application.StatusBar = "[港股 4/4] 生成指标表..."
    模块_抓港股指标表.Main
    UnhideMarketTabs "HK"

CleanUp:
    If Err.Number <> 0 Then
        runErrDesc = vbCrLf & vbCrLf & "运行中断: " & Err.Description
        Err.Clear
    End If
    g_silentMode = False
    g_diagnosticAppendOnly = False
    If hasAppState Then
        EndAppState appState
    Else
        Application.StatusBar = False
        Application.ScreenUpdating = True
    End If
    If Len(runErrDesc) > 0 Then Application.StatusBar = "一键 港股出错: " & runErrDesc

    If Not blnSilent Then ShowMarketRunSummary "港股", dtTime, runErrDesc
End Sub


Public Sub 一键韩股(Optional ByVal blnSilent As Boolean = False)
    Dim dtTime As Double: dtTime = Timer
    Dim runErrDesc As String
    Dim appState As TAppState
    Dim hasAppState As Boolean

    g_silentMode = True
    g_globalFails = 0
    g_globalLog = ""
    On Error GoTo CleanUp
    appState = BeginAppState("正在抓取 韩股...")
    hasAppState = True
    g_diagnosticSheetName = "韩股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticAppendOnly = True

    Application.StatusBar = "[韩股 1/4] 抓取资产负债表..."
    模块_抓韩股资产负债表.Main
    Application.StatusBar = "[韩股 2/4] 抓取利润表..."
    模块_抓韩股利润表.Main
    Application.StatusBar = "[韩股 3/4] 抓取现金流量表..."
    模块_抓韩股现金流量表.Main
    Application.StatusBar = "[韩股 4/4] 生成指标表..."
    模块_抓韩股指标表.Main
    UnhideMarketTabs "KR"

CleanUp:
    If Err.Number <> 0 Then
        runErrDesc = vbCrLf & vbCrLf & "运行中断: " & Err.Description
        Err.Clear
    End If
    g_silentMode = False
    g_diagnosticAppendOnly = False
    If hasAppState Then
        EndAppState appState
    Else
        Application.StatusBar = False
        Application.ScreenUpdating = True
    End If
    If Len(runErrDesc) > 0 Then Application.StatusBar = "一键 韩股出错: " & runErrDesc

    If Not blnSilent Then ShowMarketRunSummary "韩股", dtTime, runErrDesc
End Sub


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
    Dim newVisible As Long: newVisible = -1    ' xlSheetVisible
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsOfficialMarketSheet(ws) Or ws.Name = "跨市场_指标表" Then
            If ws.Visible = -1 Then newVisible = 0    ' xlSheetHidden
            Exit For
        End If
    Next ws

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If IsOfficialMarketSheet(ws) Or ws.Name = "跨市场_指标表" Then
            ws.Visible = newVisible
        End If
    Next ws
    Err.Clear
    On Error GoTo 0
End Sub


Private Function IsOfficialMarketSheet(ByVal ws As Worksheet) As Boolean
    Dim prefixes As Variant: prefixes = Array("A股_", "美股_", "港股_", "韩股_")
    Dim p As Variant
    For Each p In prefixes
        If Left$(ws.Name, Len(CStr(p))) = CStr(p) _
           And InStr(ws.Name, "抓取诊断") = 0 Then
            IsOfficialMarketSheet = True
            Exit Function
        End If
    Next p
End Function


Private Sub ToggleMarketTabsVisibility(ByVal market As String)
    Dim prefix As String
    Select Case UCase$(Trim$(market))
        Case "A":  prefix = "A股_"
        Case "US": prefix = "美股_"
        Case "HK": prefix = "港股_"
        Case "KR": prefix = "韩股_"
        Case Else: Exit Sub
    End Select

    Dim newVisible As Long: newVisible = -1    ' xlSheetVisible
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(prefix)) = prefix _
           And InStr(ws.Name, "抓取诊断") = 0 Then
            If ws.Visible = -1 Then newVisible = 0    ' xlSheetHidden
            Exit For
        End If
    Next ws

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(prefix)) = prefix _
           And InStr(ws.Name, "抓取诊断") = 0 Then
            ws.Visible = newVisible
        End If
    Next ws
    Err.Clear
    On Error GoTo 0
End Sub


Public Sub UnhideMarketTabs(ByVal market As String)
    Dim prefix As String
    Select Case UCase$(Trim$(market))
        Case "A":  prefix = "A股_"
        Case "US": prefix = "美股_"
        Case "HK": prefix = "港股_"
        Case "KR": prefix = "韩股_"
        Case Else: Exit Sub
    End Select

    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(prefix)) = prefix _
           And InStr(ws.Name, "抓取诊断") = 0 Then
            ws.Visible = -1
        End If
    Next ws
    Err.Clear
    On Error GoTo 0
End Sub


Public Sub UnhideCrossMarketIndicator()
    On Error Resume Next
    ThisWorkbook.Sheets("跨市场_指标表").Visible = -1
    Err.Clear
    On Error GoTo 0
End Sub


Public Sub 一键清空所有数据(Optional ByVal blnSilent As Boolean = False)
    Dim oldAlerts As Boolean: oldAlerts = Application.DisplayAlerts
    Dim oldScreenUpdating As Boolean: oldScreenUpdating = Application.ScreenUpdating
    Dim clearedCount As Long
    Dim runErrDesc As String
    Dim sheetName As Variant

    On Error GoTo CleanUp
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.StatusBar = "清空生成的报表数据..."

    For Each sheetName In Array( _
        "A股_资产负债表", "A股_利润表", "A股_现金流量表", "A股_指标表", _
        "美股_资产负债表", "美股_利润表", "美股_现金流量表", "美股_指标表", _
        "港股_资产负债表", "港股_利润表", "港股_现金流量表", "港股_指标表", _
        "韩股_资产负债表", "韩股_利润表", "韩股_现金流量表", "韩股_指标表", _
        "跨市场_指标表")
        clearedCount = clearedCount + ClearSheetContentsIfExists(CStr(sheetName))
    Next sheetName

    clearedCount = clearedCount + ClearDiagnosticRowsIfExists("美股_抓取诊断")
    clearedCount = clearedCount + ClearDiagnosticRowsIfExists("港股_抓取诊断")
    clearedCount = clearedCount + ClearDiagnosticRowsIfExists("韩股_抓取诊断")

CleanUp:
    If Err.Number <> 0 Then
        runErrDesc = Err.Description
        Err.Clear
    End If
    Application.StatusBar = False
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreenUpdating

    If Not blnSilent Then
        If Len(runErrDesc) > 0 Then
            MsgBox "清空数据失败:" & vbCrLf & runErrDesc, _
                   vbExclamation, "上市公司财务数据查询"
        Else
            MsgBox "已清空生成的报表数据和格式" & vbCrLf & _
                   "样本池公司、参数、汇率和 HTTP 缓存已保留。", _
                   vbInformation, "上市公司财务数据查询"
        End If
    End If
End Sub


Private Function ClearSheetContentsIfExists(ByVal sheetName As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        ws.Cells.UnMerge
        ws.UsedRange.Clear
        ClearSheetContentsIfExists = 1
    End If
    Err.Clear
    On Error GoTo 0
End Function


Private Function ClearDiagnosticRowsIfExists(ByVal sheetName As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        Dim lastRow As Long
        lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If lastRow >= 3 Then
            ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 17)).Clear
        End If
        ClearDiagnosticRowsIfExists = 1
    End If
    Err.Clear
    On Error GoTo 0
End Function


Private Function MarketHasPoolRows(ByVal marketKey As String) As Boolean
    Dim codeCol As Long
    Select Case UCase$(Trim$(marketKey))
        Case "A": codeCol = POOL_A_CODE_COL
        Case "US": codeCol = POOL_US_CODE_COL
        Case "HK": codeCol = POOL_HK_CODE_COL
        Case "KR": codeCol = POOL_KR_CODE_COL
        Case Else: Exit Function
    End Select

    Dim wsPool As Worksheet: Set wsPool = ThisWorkbook.Sheets("样本池")
    Dim lastRow As Long
    lastRow = wsPool.Cells(wsPool.Rows.Count, codeCol).End(xlUp).Row
    If lastRow < POOL_DATA_START_ROW Then Exit Function

    Dim r As Long, codeText As String
    For r = POOL_DATA_START_ROW To lastRow
        codeText = Trim$(CStr(wsPool.Cells(r, codeCol).Value))
        If Len(codeText) > 0 And codeText <> "代码" Then
            MarketHasPoolRows = True
            Exit Function
        End If
    Next r
End Function


Public Sub 一键跨市场指标表(Optional ByVal blnSilent As Boolean = False)
    Dim dtTime As Double: dtTime = Timer
    Dim runErrDesc As String
    On Error GoTo CleanUp

    Application.StatusBar = "刷新跨市场指标表..."
    BuildCrossMarketIndicatorSheet
    UnhideCrossMarketIndicator

CleanUp:
    If Err.Number <> 0 Then
        runErrDesc = Err.Description
        Err.Clear
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True

    If Not blnSilent Then
        If Len(runErrDesc) > 0 Then
            MsgBox "跨市场指标表刷新失败:" & vbCrLf & runErrDesc, _
                   vbExclamation, "上市公司财务数据查询"
        Else
            MsgBox "跨市场指标表刷新完成" & vbCrLf & _
                   "用时: " & Format(Timer - dtTime, "0.0 秒"), _
                   vbInformation, "上市公司财务数据查询"
        End If
    End If
End Sub


Private Sub ShowMarketRunSummary(ByVal marketName As String, ByVal dtTime As Double, ByVal runErrDesc As String)
    Dim msg As String
    msg = "一键" & marketName & "完成" & vbCrLf & _
          "总用时: " & Format(Timer - dtTime, "0.0 秒")
    If g_globalFails > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "失败 " & g_globalFails & " 条:" & g_globalLog
    Else
        msg = msg & vbCrLf & "全部成功"
    End If
    msg = msg & runErrDesc

    Dim style As Long: style = vbInformation
    If g_globalFails > 0 Or Len(runErrDesc) > 0 Then style = vbExclamation
    MsgBox msg, style, "上市公司财务数据查询"
End Sub


Public Sub 一键全抓(Optional ByVal blnSilent As Boolean = False)
    Dim dtTime As Double: dtTime = Timer
    Dim hasAnyMarket As Boolean
    Dim appState As TAppState
    Dim hasAppState As Boolean

    ' 重置全局累计
    g_silentMode = True
    g_globalFails = 0
    g_globalLog = ""

    On Error GoTo CleanUp
    appState = BeginAppState("一键全抓准备中...")
    hasAppState = True

    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticSheetName = "港股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticSheetName = "韩股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticAppendOnly = True

    If MarketHasPoolRows("A") Then
        hasAnyMarket = True
        Application.StatusBar = "[A股 1/4] 抓取资产负债表..."
        DoEvents
        模块_抓资产负债表.Main
        Application.StatusBar = "[A股 2/4] 抓取利润表..."
        DoEvents
        模块_抓利润表.Main
        Application.StatusBar = "[A股 3/4] 抓取现金流量表..."
        DoEvents
        模块_抓现金流量表.Main
        Application.StatusBar = "[A股 4/4] 生成指标表..."
        DoEvents
        模块_抓指标表.Main
        UnhideMarketTabs "A"
    End If

    If MarketHasPoolRows("US") Then
        hasAnyMarket = True
        Application.StatusBar = "[美股 1/4] 抓取资产负债表..."
        DoEvents
        模块_抓美股资产负债表.Main
        Application.StatusBar = "[美股 2/4] 抓取利润表..."
        DoEvents
        模块_抓美股利润表.Main
        Application.StatusBar = "[美股 3/4] 抓取现金流量表..."
        DoEvents
        模块_抓美股现金流量表.Main
        Application.StatusBar = "[美股 4/4] 生成指标表..."
        DoEvents
        模块_抓美股指标表.Main
        UnhideMarketTabs "US"
    End If

    If MarketHasPoolRows("HK") Then
        hasAnyMarket = True
        Application.StatusBar = "[港股 1/4] 抓取资产负债表..."
        DoEvents
        模块_抓港股资产负债表.Main
        Application.StatusBar = "[港股 2/4] 抓取利润表..."
        DoEvents
        模块_抓港股利润表.Main
        Application.StatusBar = "[港股 3/4] 抓取现金流量表..."
        DoEvents
        模块_抓港股现金流量表.Main
        Application.StatusBar = "[港股 4/4] 生成指标表..."
        DoEvents
        模块_抓港股指标表.Main
        UnhideMarketTabs "HK"
    End If

    If MarketHasPoolRows("KR") Then
        hasAnyMarket = True
        Application.StatusBar = "[韩股 1/4] 抓取资产负债表..."
        DoEvents
        模块_抓韩股资产负债表.Main
        Application.StatusBar = "[韩股 2/4] 抓取利润表..."
        DoEvents
        模块_抓韩股利润表.Main
        Application.StatusBar = "[韩股 3/4] 抓取现金流量表..."
        DoEvents
        模块_抓韩股现金流量表.Main
        Application.StatusBar = "[韩股 4/4] 生成指标表..."
        DoEvents
        模块_抓韩股指标表.Main
        UnhideMarketTabs "KR"
    End If

    ' Phase 4j.1: 一键全抓后只刷新跨市场指标表
    On Error Resume Next
    BuildCrossMarketIndicatorSheet
    UnhideCrossMarketIndicator
    Err.Clear
    On Error GoTo CleanUp

CleanUp:
    Dim runErrDesc As String
    If Err.Number <> 0 Then
        runErrDesc = vbCrLf & vbCrLf & "运行中断: " & Err.Description
        Err.Clear
    End If

    g_silentMode = False
    g_diagnosticAppendOnly = False
    If hasAppState Then
        EndAppState appState
    Else
        Application.StatusBar = False
        Application.ScreenUpdating = True
    End If

    Dim msg As String
    msg = "一键全抓完成 (A股 + 美股 + 港股 + 韩股)" & vbCrLf & _
          "总用时: " & Format(Timer - dtTime, "0.0 秒")
    If Not hasAnyMarket Then
        msg = msg & vbCrLf & "未检测到样本池公司, 未执行抓数。"
    ElseIf g_globalFails > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "失败 " & g_globalFails & " 条:" & g_globalLog
    Else
        msg = msg & vbCrLf & "全部成功"
    End If
    msg = msg & runErrDesc

    If Not blnSilent Then
        Dim style As Long: style = vbInformation
        If g_globalFails > 0 Or Len(runErrDesc) > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询"
    End If
End Sub

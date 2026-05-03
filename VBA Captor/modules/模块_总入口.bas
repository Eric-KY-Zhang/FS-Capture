Attribute VB_Name = "模块_总入口"
Option Explicit

' =================================================================
'  Sub 一键全抓: 顺序调用 A 股 + 美股 + 港股 12 张表, 静默模式, 最后弹一次汇总
'  基本资料 已废弃
' =================================================================

Public Sub 一键全抓(Optional ByVal blnSilent As Boolean = False)
    Dim dtTime As Double: dtTime = Timer

    ' 重置全局累计
    g_silentMode = True
    g_globalFails = 0
    g_globalLog = ""

    On Error GoTo CleanUp
    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticSheetName = "港股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticAppendOnly = True

    ' A 股资产负债表
    Application.StatusBar = "[1/12] 抓取A股资产负债表..."
    DoEvents
    模块_抓资产负债表.Main

    ' A 股利润表
    Application.StatusBar = "[2/12] 抓取A股利润表..."
    DoEvents
    模块_抓利润表.Main

    ' A 股现金流量表
    Application.StatusBar = "[3/12] 抓取A股现金流量表..."
    DoEvents
    模块_抓现金流量表.Main

    ' A 股指标表
    Application.StatusBar = "[4/12] 生成A股指标表..."
    DoEvents
    模块_抓指标表.Main

    ' 美股资产负债表
    Application.StatusBar = "[5/12] 抓取美股资产负债表..."
    DoEvents
    模块_抓美股资产负债表.Main

    ' 美股利润表
    Application.StatusBar = "[6/12] 抓取美股利润表..."
    DoEvents
    模块_抓美股利润表.Main

    ' 美股现金流量表
    Application.StatusBar = "[7/12] 抓取美股现金流量表..."
    DoEvents
    模块_抓美股现金流量表.Main

    ' 美股指标表
    Application.StatusBar = "[8/12] 生成美股指标表..."
    DoEvents
    模块_抓美股指标表.Main

    ' 港股资产负债表
    Application.StatusBar = "[9/12] 抓取港股资产负债表..."
    DoEvents
    模块_抓港股资产负债表.Main

    ' 港股利润表
    Application.StatusBar = "[10/12] 抓取港股利润表..."
    DoEvents
    模块_抓港股利润表.Main

    ' 港股现金流量表
    Application.StatusBar = "[11/12] 抓取港股现金流量表..."
    DoEvents
    模块_抓港股现金流量表.Main

    ' 港股指标表
    Application.StatusBar = "[12/12] 生成港股指标表..."
    DoEvents
    模块_抓港股指标表.Main

CleanUp:
    Dim runErrDesc As String
    If Err.Number <> 0 Then
        runErrDesc = vbCrLf & vbCrLf & "运行中断: " & Err.Description
        Err.Clear
    End If

    g_silentMode = False
    g_diagnosticAppendOnly = False
    Application.StatusBar = False
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "一键全抓完成 (A股 + 美股 + 港股)" & vbCrLf & _
          "总用时: " & Format(Timer - dtTime, "0.0 秒")
    If g_globalFails > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "失败 " & g_globalFails & " 条:" & g_globalLog
    Else
        msg = msg & vbCrLf & "全部成功 ✓"
    End If
    msg = msg & runErrDesc

    If Not blnSilent Then
        Dim style As Long: style = vbInformation
        If g_globalFails > 0 Or Len(runErrDesc) > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询"
    End If
End Sub

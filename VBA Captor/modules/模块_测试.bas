Attribute VB_Name = "模块_测试"
Option Explicit

' =================================================================
'  Phase 4f reviewer smoke tests
'  本模块只做本地回归验证, 不触发任何网络抓数。
' =================================================================

Public Function TestOptionalBool(Optional ByVal blnSilent As Boolean = False) As Boolean
    TestOptionalBool = blnSilent
End Function


Public Sub TestStep3Smoke()
    Dim wsPool As Worksheet
    Set wsPool = ThisWorkbook.Worksheets("样本池")

    Dim savedB6 As Variant
    savedB6 = wsPool.Range("E6").Value

    Dim wsYuanbi As Worksheet
    Dim wsRmb As Worksheet
    Set wsYuanbi = GetOrClearSmokeSheet("_phase4f_step3_yuanbi")
    Set wsRmb = GetOrClearSmokeSheet("_phase4f_step3_rmb")

    Dim arrCodes(1 To 1) As String
    arrCodes(1) = "300866"

    Dim arrPeriods(1 To 2) As String
    arrPeriods(1) = "2024-12-31"
    arrPeriods(2) = "2023-12-31"

    Dim arrIndicators(1 To 2) As String
    arrIndicators(1) = "总资产"
    arrIndicators(2) = "审计意见"

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    dictCompanyName.Add "300866", "安克创新"

    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.Add "总资产", "资产"
    dictCategory.Add "审计意见", "文本"

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictPer2024 As Object: Set dictPer2024 = CreateObject("Scripting.Dictionary")
    Dim dictPer2023 As Object: Set dictPer2023 = CreateObject("Scripting.Dictionary")

    dictPer2024.Add "总资产", 123.45
    dictPer2024.Add "审计意见", "标准无保留"
    dictPer2023.Add "总资产", 67.89
    dictPer2023.Add "审计意见", "标准无保留"
    dictCompany.Add "2024-12-31", dictPer2024
    dictCompany.Add "2023-12-31", dictPer2023
    dictData.Add "300866", dictCompany

    wsPool.Range("E6").Value = "原币"
    WriteWideTable wsYuanbi, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=Nothing, statementKind:="BalanceSheet"

    wsPool.Range("E6").Value = "统一RMB"
    WriteWideTable wsRmb, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=Nothing, statementKind:="BalanceSheet"

    wsPool.Range("E6").Value = savedB6
End Sub


Public Sub TestStep45Smoke()
    Dim wsPool As Worksheet
    Set wsPool = ThisWorkbook.Worksheets("样本池")

    Dim savedB6 As Variant
    savedB6 = wsPool.Range("E6").Value

    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet

    Dim rows As Collection
    Set rows = New Collection
    AddDiagnosticRow rows, "AAPL", "BalanceSheet", "Total assets", "OK", "EDGAR", _
                     "us-gaap", "Assets", "USD", "100", "smoke", "7.123456"
    WriteDiagnosticForKind "BalanceSheet", rows

    Dim wsDiag As Worksheet
    Set wsDiag = ThisWorkbook.Worksheets("美股_抓取诊断")
    If CStr(wsDiag.Cells(2, 11).Value) <> "FX_Rate" Then _
        Err.Raise vbObjectError + 742, "TestStep45Smoke", "诊断 K 列表头不是 FX_Rate"
    If CStr(wsDiag.Cells(3, 11).Value) <> "7.123456" Then _
        Err.Raise vbObjectError + 743, "TestStep45Smoke", "诊断 K3 没写入 fx rate"

    Dim wsTag As Worksheet
    Set wsTag = GetOrClearSmokeSheet("_phase4f_step5_tag")

    Dim arrCodes(1 To 1) As String
    arrCodes(1) = "AAPL"

    Dim arrPeriods(1 To 1) As String
    arrPeriods(1) = "2024-12-31"

    Dim arrIndicators(1 To 1) As String
    arrIndicators(1) = "TextOnly"

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    dictCompanyName.Add "AAPL", "Apple"

    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.Add "TextOnly", "Smoke"

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictPer As Object: Set dictPer = CreateObject("Scripting.Dictionary")
    dictPer.Add "TextOnly", "non-numeric"
    dictCompany.Add "2024-12-31", dictPer
    dictData.Add "AAPL", dictCompany

    Dim dictCurrency As Object: Set dictCurrency = CreateObject("Scripting.Dictionary")
    dictCurrency.Add "AAPL", "USD"

    wsPool.Range("E6").Value = "统一RMB"
    WriteWideTable wsTag, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=dictCurrency, statementKind:="BalanceSheet"
    RefreshA1CurrencyComment wsTag, "美股_资产负债表"

    If InStr(CStr(wsTag.Cells(1, 3).Value), "[USD→RMB]") = 0 Then _
        Err.Raise vbObjectError + 744, "TestStep45Smoke", "R1 缺少 USD→RMB tag"
    If wsTag.Range("A1").Comment Is Nothing Then _
        Err.Raise vbObjectError + 745, "TestStep45Smoke", "A1 缺少动态注释"
    If InStr(wsTag.Range("A1").Comment.Text, "统一汇率换算") = 0 Then _
        Err.Raise vbObjectError + 746, "TestStep45Smoke", "A1 注释不是统一RMB文案"

    wsPool.Range("E6").Value = savedB6
End Sub


Public Sub TestPhase4hToggleSmoke()
    Dim wsPool As Worksheet
    Set wsPool = ThisWorkbook.Worksheets("样本池")

    Dim savedB6 As Variant
    savedB6 = wsPool.Range("E6").Value

    Dim wsSmoke As Worksheet
    Set wsSmoke = GetOrClearSmokeSheet("_phase4h_toggle_smoke")

    Dim arrCodes(1 To 1) As String
    arrCodes(1) = "AAPL"

    Dim arrPeriods(1 To 1) As String
    arrPeriods(1) = "2024-12-31"

    Dim arrIndicators(1 To 1) As String
    arrIndicators(1) = "Total assets"

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    dictCompanyName.Add "AAPL", "Apple"

    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.Add "Total assets", "Assets"

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictPer As Object: Set dictPer = CreateObject("Scripting.Dictionary")
    dictPer.Add "Total assets", 100#
    dictCompany.Add "2024-12-31", dictPer
    dictData.Add "AAPL", dictCompany

    Dim dictCurrency As Object: Set dictCurrency = CreateObject("Scripting.Dictionary")
    dictCurrency.Add "AAPL", "USD"

    wsPool.Range("E6").Value = "原币"
    WriteWideTable wsSmoke, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=dictCurrency, statementKind:="BalanceSheet"

    wsPool.Range("E6").Value = savedB6
End Sub


Private Function GetOrClearSmokeSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set GetOrClearSmokeSheet = ws
End Function

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
    savedB6 = wsPool.Range("B6").Value

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

    wsPool.Range("B6").Value = "原币"
    WriteWideTable wsYuanbi, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=Nothing, statementKind:="BalanceSheet"

    wsPool.Range("B6").Value = "统一RMB"
    WriteWideTable wsRmb, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=Nothing, statementKind:="BalanceSheet"

    wsPool.Range("B6").Value = savedB6
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

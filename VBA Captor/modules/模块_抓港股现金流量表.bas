Attribute VB_Name = "模块_抓港股现金流量表"
Option Explicit

' =================================================================
'  抓港股 Cash Flow Statement — Phase 4c
'  写入 Sheet: 港股_现金流量表
' =================================================================

Public Sub Main()
    RunHKStatement "CashFlow", "港股_现金流量表", GetHKCFConcepts(), 8
End Sub


Private Function GetHKCFConcepts() As Variant
    Dim a(0 To 10) As Variant
    a(0) = Array("Operating Activities", "Cash from operations")
    a(1) = Array("Operating Activities", "Depreciation & amortization")
    a(2) = Array("Investing Activities", "Cash from investing")
    a(3) = Array("Investing Activities", "Capex")
    a(4) = Array("Financing Activities", "Cash from financing")
    a(5) = Array("Financing Activities", "Dividends paid")
    a(6) = Array("Financing Activities", "Interest paid")
    a(7) = Array("Operating Activities", "Interest received")
    a(8) = Array("Cash", "FX effect on cash")
    a(9) = Array("Cash", "Cash at beginning of period")
    a(10) = Array("Cash", "Cash at end of period")
    GetHKCFConcepts = a
End Function

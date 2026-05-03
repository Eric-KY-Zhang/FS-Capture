Attribute VB_Name = "模块_抓韩股现金流量表"
Option Explicit

' =================================================================
'  抓韩股 Cash Flow Statement — Phase 4d
'  写入 Sheet: 韩股_现金流量表
' =================================================================

Public Sub Main()
    RunKRStatement "CashFlow", "韩股_现金流量表", GetKRCFConcepts(), 8
End Sub


Private Function GetKRCFConcepts() As Variant
    Dim a(0 To 12) As Variant
    a(0) = Array("Operating Activities", "Net income")
    a(1) = Array("Operating Activities", "Depreciation & amortization")
    a(2) = Array("Operating Activities", "Change in accounts receivable")
    a(3) = Array("Operating Activities", "Change in inventory")
    a(4) = Array("Operating Activities", "Change in accounts payable")
    a(5) = Array("Operating Activities", "Cash from operations")
    a(6) = Array("Investing Activities", "Capex")
    a(7) = Array("Investing Activities", "Cash from investing")
    a(8) = Array("Financing Activities", "Cash from financing")
    a(9) = Array("Financing Activities", "Dividends paid")
    a(10) = Array("Cash", "FX effect on cash")
    a(11) = Array("Cash", "Net cash flow")
    a(12) = Array("Cash", "Free cash flow")
    GetKRCFConcepts = a
End Function

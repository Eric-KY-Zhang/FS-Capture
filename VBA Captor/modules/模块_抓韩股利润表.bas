Attribute VB_Name = "模块_抓韩股利润表"
Option Explicit

' =================================================================
'  抓韩股 Income Statement — Phase 4d
'  写入 Sheet: 韩股_利润表
' =================================================================

Public Sub Main()
    RunKRStatement "Income", "韩股_利润表", GetKRISConcepts(), 8
End Sub


Private Function GetKRISConcepts() As Variant
    Dim a(0 To 12) As Variant
    a(0) = Array("Revenue", "Revenue")
    a(1) = Array("Revenue", "Cost of goods & services sold")
    a(2) = Array("Profitability", "Gross profit")
    a(3) = Array("Operating Expenses", "R&D expense")
    a(4) = Array("Operating Expenses", "SG&A expense")
    a(5) = Array("Operating Expenses", "Total operating expenses")
    a(6) = Array("Profitability", "Operating income")
    a(7) = Array("Non Operating", "Interest expense")
    a(8) = Array("Profitability", "Pre-tax income")
    a(9) = Array("Tax", "Income tax expense")
    a(10) = Array("Profitability", "Net income")
    a(11) = Array("EPS", "Basic EPS", "", "KRW/share", 1#)
    a(12) = Array("EPS", "Diluted EPS", "", "KRW/share", 1#)
    GetKRISConcepts = a
End Function

Attribute VB_Name = "模块_抓港股资产负债表"
Option Explicit

' =================================================================
'  抓港股 Balance Sheet — Phase 4c
'  写入 Sheet: 港股_资产负债表
' =================================================================

Public Sub Main()
    RunHKStatement "BalanceSheet", "港股_资产负债表", GetHKBSConcepts(), 8
End Sub


Private Function GetHKBSConcepts() As Variant
    Dim a(0 To 17) As Variant
    a(0) = Array("Current Assets", "Cash & equivalents")
    a(1) = Array("Current Assets", "Accounts receivable, net")
    a(2) = Array("Current Assets", "Inventory")
    a(3) = Array("Current Assets", "Total current assets")
    a(4) = Array("Non-Current Assets", "Property, plant & equipment, net")
    a(5) = Array("Non-Current Assets", "Investments")
    a(6) = Array("Non-Current Assets", "Total non-current assets")
    a(7) = Array("Assets", "Total assets")
    a(8) = Array("Current Liabilities", "Accounts payable")
    a(9) = Array("Current Liabilities", "Short-term debt")
    a(10) = Array("Current Liabilities", "Total current liabilities")
    a(11) = Array("Non-Current Liabilities", "Long-term debt")
    a(12) = Array("Non-Current Liabilities", "Total non-current liabilities")
    a(13) = Array("Liabilities", "Total liabilities")
    a(14) = Array("Equity", "Minority interests")
    a(15) = Array("Equity", "Total equity")
    a(16) = Array("Equity", "Total stockholders' equity")
    a(17) = Array("Liabilities & Equity", "Total liabilities & equity")
    GetHKBSConcepts = a
End Function

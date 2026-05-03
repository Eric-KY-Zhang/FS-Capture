Attribute VB_Name = "模块_抓资产负债表"
Option Explicit

' =================================================================
'  抓 A 股 资产负债表, 写入『资产负债表』Sheet
'  URL 由 RunOneStatement 内部按代码 + A2 年份 自拼 (新浪 vFD_BalanceSheet)
' =================================================================

Public Sub Main()
    RunOneStatement "BalanceSheetNewTable0", "balance", "A股_资产负债表", False, True
End Sub

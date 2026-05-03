Attribute VB_Name = "模块_抓利润表"
Option Explicit

' =================================================================
'  抓 A 股 利润表, 写入『利润表』Sheet
'  HTML table id = "ProfitStatementNewTable0"
' =================================================================

Public Sub Main()
    RunOneStatement "ProfitStatementNewTable0", "profit", "A股_利润表", False, True
End Sub

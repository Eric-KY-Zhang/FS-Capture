Attribute VB_Name = "模块_抓现金流量表"
Option Explicit

' =================================================================
'  抓 A 股 现金流量表, 写入『现金流量表』Sheet
'  注意: 新浪在现金流量表页面里复用了 "ProfitStatementNewTable0" 这个 id (历史遗留)
' =================================================================

Public Sub Main()
    RunOneStatement "ProfitStatementNewTable0", "cash", "A股_现金流量表"
End Sub

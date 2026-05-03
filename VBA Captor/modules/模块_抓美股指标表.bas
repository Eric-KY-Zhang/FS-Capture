Attribute VB_Name = "模块_抓美股指标表"
Option Explicit

' =================================================================
'  美股标准指标表
'  不再抓 EDGAR / 雪球 Indicator 原始 per-share 指标, 只根据美股 BS / IS 生成 18 个标准指标
' =================================================================

Public Sub Main()
    BuildStandardIndicatorSheet "US"
End Sub

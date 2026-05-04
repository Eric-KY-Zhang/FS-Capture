Attribute VB_Name = "模块_工具函数"
Option Explicit

' =================================================================
'  上市公司财务数据查询 — 公共工具函数
'  作者: 基于林铖 V2.2 重写, V3.0
'
'  本模块提供:
'    HttpGet            : WinHttp 抓取 + gb2312 解码 + 异常抛出
'    ExtractTable       : 用正则截取 <table id="..."> 片段
'    ParseFinancialHtml : 解析单家公司的财报 HTML, 累计到共享字典
'    ParseCorpInfoHtml  : 解析公司基本资料页, 返回 5 字段字典
'    WriteWideTable     : 把并集后的二维数据写到宽表 Sheet
'    BuildSinaUrls      : (已废弃) 按代码自动拼新浪 5 个 URL — 用户改用 HYPERLINK 公式
'    SetBorderLine      : 给 Range 加细边框 (来自 V2.2 模块2)
'    ExchangePrefix     : 按代码前缀推断 sh/sz
' =================================================================

' --------- 样本池约定 (兄弟模块共享)
'   Row 1-5: 全局配置区 (A1=年份, A2=year, A3=季度, A4=quarter, B5:F5=雪球 cookie)
'   Row 7  : 分市场标题
'   Row 8  : 分市场一键按钮
'   Row 9  : 各市场数据表头
'   Row 10+: 股票数据
'   列    : A:B=A股, E:F=美股, I:J=港股, M:N=韩股
'   URL 不再存 sheet, A 股抓数模块内部按代码+年份自拼 URL
Public Const POOL_DATA_START_ROW As Long = 10
Public Const POOL_A_CODE_COL As Long = 1
Public Const POOL_A_NAME_COL As Long = 2
Public Const POOL_US_CODE_COL As Long = 5
Public Const POOL_US_NAME_COL As Long = 6
Public Const POOL_HK_CODE_COL As Long = 9
Public Const POOL_HK_NAME_COL As Long = 10
Public Const POOL_KR_CODE_COL As Long = 13
Public Const POOL_KR_NAME_COL As Long = 14
Public Const POOL_LAST_COL As Long = 14       ' 新布局数据区到 N 列
Public Const POOL_MARKET_COL As Long = 3      ' legacy helper only: 旧 A:C 布局市场列

' --------- 全局状态 (用于一键全抓的静默调用 + 汇总错误) ---------
Public g_silentMode As Boolean      ' 一键全抓时设为 True, 各 Main 不弹 MsgBox
Public g_globalFails As Long
Public g_globalLog As String
Public g_diagnosticAppendOnly As Boolean    ' True=一键全抓累计追加; False=单表重写该报表诊断
Public g_diagnosticSheetName As String      ' 诊断目标 sheet; 默认美股_抓取诊断, 港股/韩股入口切到对应诊断 sheet

' --------- Phase 4b: 美股 ticker → CIK 会话级缓存 (Public 必须在所有 Sub 之前声明) ---------
Public g_dictTickerToCIK As Object

' (g_diagnosticAppendOnly 已在上方 line 33 声明 — Codex Layer 1 同步引入, 不重复)

' --------- Phase 4f Step 2: 汇率 sheet 常量 (跨模块共享) ---------
'   FX_SHEET     : 汇率缓存 sheet 名 (build_template/install_modules 同步)
'   FX_DATA_ROW  : 数据起始行 (Row 1 是表头)
Public Const FX_SHEET As String = "汇率"
Public Const FX_DATA_ROW As Long = 2


' --------- 自动化/一键全抓用: 控制各抓数入口是否弹窗 ---------
Public Sub SetSilentMode(ByVal blnSilent As Boolean)
    g_silentMode = blnSilent
End Sub


' --------- Phase 4b-14a: 美股 conceptMap entry 兼容读取 ---------
Public Function MapEntryCategory(ByVal entry As Variant) As String
    MapEntryCategory = CStr(entry(0))
End Function


Public Function MapEntryLabel(ByVal entry As Variant) As String
    MapEntryLabel = CStr(entry(1))
End Function


Public Function MapEntryUsGaapConcepts(ByVal entry As Variant) As String
    MapEntryUsGaapConcepts = CStr(entry(2))
End Function


Public Function MapEntryIfrsConcepts(ByVal entry As Variant) As String
    If UBound(entry) >= 5 Then
        MapEntryIfrsConcepts = CStr(entry(5))
    Else
        MapEntryIfrsConcepts = CStr(entry(2))
    End If
End Function


Public Function MapEntryUnit(ByVal entry As Variant) As String
    If UBound(entry) >= 4 Then
        MapEntryUnit = CStr(entry(3))
    Else
        MapEntryUnit = "USD"
    End If
End Function


Public Function MapEntryScale(ByVal entry As Variant) As Double
    If UBound(entry) >= 4 Then
        MapEntryScale = CDbl(entry(4))
    Else
        MapEntryScale = 1000000#
    End If
End Function


Public Function MapEntryFuzzyHint(ByVal entry As Variant) As String
    If UBound(entry) >= 6 Then
        MapEntryFuzzyHint = CStr(entry(6))
    Else
        MapEntryFuzzyHint = ""
    End If
End Function


Public Function CoreLabelsForKind(ByVal strKind As String) As Variant
    Select Case strKind
        Case "BalanceSheet"
            CoreLabelsForKind = Array("Total assets")
        Case "Income"
            CoreLabelsForKind = Array("Revenue", "Net income")
        Case "CashFlow"
            CoreLabelsForKind = Array("Cash from operations", "Cash at end of period")
        Case Else
            CoreLabelsForKind = Array()
    End Select
End Function


' --------- Phase 4b-14a: 美股抓取诊断表 ---------
Public Function CurrentDiagnosticSheetName() As String
    Dim s As String: s = Trim$(g_diagnosticSheetName)
    If Len(s) = 0 Then s = "美股_抓取诊断"
    CurrentDiagnosticSheetName = s
End Function


Public Sub EnsureDiagnosticSheet()
    Dim ws As Worksheet
    Dim diagName As String: diagName = CurrentDiagnosticSheetName()
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(diagName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = diagName
    End If

    Dim headers As Variant
    headers = Array("公司", "报表", "输出指标", "状态", "数据源", _
                    "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注", "FX_Rate")

    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).UnMerge
    Err.Clear
    On Error GoTo 0

    ws.Cells(1, 1).Value = Replace(diagName, "_", "") & " (每次跑数后自动刷新)"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).Merge
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 11))
        .Font.Name = "微软雅黑"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim i As Long
    For i = 0 To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Name = "微软雅黑"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 42
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 10
    ws.Columns("J").ColumnWidth = 58
    ws.Columns("K").ColumnWidth = 12
    ws.Range("A:A").NumberFormat = "@"
    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 20
    Call SetBorderLine(ws.Range(ws.Cells(1, 1), ws.Cells(2, 11)))

    On Error Resume Next
    ws.Visible = xlSheetHidden
    Err.Clear
    On Error GoTo 0
End Sub


Public Sub ClearDiagnosticSheet()
    EnsureDiagnosticSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CurrentDiagnosticSheetName())
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1000 Then lastRow = 1000
    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 11)).ClearContents
End Sub


Public Sub DeleteDiagnosticRowsForKind(ByVal strKind As String)
    EnsureDiagnosticSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CurrentDiagnosticSheetName())
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = lastRow To 3 Step -1
        If CStr(ws.Cells(r, 2).Value) = strKind Then ws.Rows(r).Delete
    Next r
End Sub


Public Sub AddDiagnosticRow(ByVal collRows As Collection, _
                            ByVal ticker As String, _
                            ByVal strKind As String, _
                            ByVal label As String, _
                            ByVal statusText As String, _
                            ByVal sourceText As String, _
                            ByVal taxonomyText As String, _
                            ByVal fieldText As String, _
                            ByVal unitText As String, _
                            ByVal scoreText As String, _
                            ByVal noteText As String, _
                            Optional ByVal fxRateText As String = "1.0")
    If collRows Is Nothing Then Exit Sub
    collRows.Add Array(ticker, strKind, label, statusText, sourceText, _
                       taxonomyText, fieldText, unitText, scoreText, noteText, fxRateText)
End Sub


Public Sub WriteDiagnosticForKind(ByVal strKind As String, ByVal collRows As Collection)
    EnsureDiagnosticSheet
    If Not g_diagnosticAppendOnly Then DeleteDiagnosticRowsForKind strKind
    If collRows Is Nothing Then Exit Sub
    If collRows.Count = 0 Then Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CurrentDiagnosticSheetName())
    Dim startRow As Long
    startRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If startRow < 3 Then startRow = 3

    Dim arrOut As Variant
    ReDim arrOut(1 To collRows.Count, 1 To 11)

    Dim i As Long, j As Long, rowData As Variant
    For i = 1 To collRows.Count
        rowData = collRows.Item(i)
        For j = 0 To 10
            If j <= UBound(rowData) Then
                arrOut(i, j + 1) = rowData(j)
            Else
                arrOut(i, j + 1) = "1.0"
            End If
        Next j
    Next i

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + collRows.Count - 1, 11)).Value = arrOut
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + collRows.Count - 1, 11))
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(startRow, 9), ws.Cells(startRow + collRows.Count - 1, 9)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(startRow, 11), ws.Cells(startRow + collRows.Count - 1, 11)).HorizontalAlignment = xlRight
    Call SetBorderLine(ws.Range(ws.Cells(1, 1), ws.Cells(startRow + collRows.Count - 1, 11)))
End Sub


Public Sub AddFuzzyDiagnosticCandidates(ByVal collRows As Collection, _
                                        ByVal ticker As String, _
                                        ByVal strKind As String, _
                                        ByVal label As String, _
                                        ByVal taxonomyName As String, _
                                        ByVal taxonomy As Object, _
                                        ByVal fuzzyHint As String, _
                                        ByVal expectedUnit As String)
    If collRows Is Nothing Then Exit Sub
    If taxonomy Is Nothing Then Exit Sub

    Dim posTokens As Variant: posTokens = FuzzyPositiveTokens(label, fuzzyHint)
    Dim negTokens As Variant: negTokens = FuzzyNegativeTokens(fuzzyHint)
    Dim topConcept(1 To 3) As String, topUnit(1 To 3) As String, topScore(1 To 3) As Double

    Dim k As Variant
    For Each k In taxonomy.Keys
        Dim conceptName As String: conceptName = CStr(k)
        If FuzzyHasNegative(conceptName, negTokens) Then GoTo NextConcept

        Dim score As Double
        score = FuzzyKeywordScore(conceptName, posTokens)
        If score <= 0 Then GoTo NextConcept

        Dim unitText As String
        If FuzzyConceptHasUnit(taxonomy.Item(conceptName), expectedUnit, unitText) Then
            score = score + 3
        End If

        If score >= 5 Then FuzzyInsertTop topConcept, topUnit, topScore, conceptName, unitText, score
NextConcept:
    Next k

    Dim i As Long
    For i = 1 To 3
        If Len(topConcept(i)) > 0 Then
            AddDiagnosticRow collRows, ticker, strKind, label, "RECOMMEND_FUZZY", "—", _
                             taxonomyName, topConcept(i), topUnit(i), Format$(topScore(i), "0.0"), _
                             "fuzzy_candidate (人肉确认后回填到 hardcode)"
        End If
    Next i
End Sub


Private Function FuzzyPositiveTokens(ByVal label As String, ByVal fuzzyHint As String) As Variant
    Dim s As String
    If InStr(1, fuzzyHint, "~", vbTextCompare) > 0 Then
        s = Left$(fuzzyHint, InStr(1, fuzzyHint, "~", vbTextCompare) - 1)
    Else
        s = fuzzyHint
    End If
    If Len(Trim$(s)) = 0 Then s = label
    s = Replace(s, "|", " ")
    FuzzyPositiveTokens = FuzzyTokenize(s)
End Function


Private Function FuzzyNegativeTokens(ByVal fuzzyHint As String) As Variant
    Dim s As String
    If InStr(1, fuzzyHint, "~", vbTextCompare) > 0 Then
        s = Mid$(fuzzyHint, InStr(1, fuzzyHint, "~", vbTextCompare) + 1)
    Else
        s = ""
    End If
    s = Replace(s, ",", " ")
    FuzzyNegativeTokens = FuzzyTokenize(s)
End Function


Private Function FuzzyTokenize(ByVal textValue As String) As Variant
    Dim s As String: s = LCase$(textValue)
    s = Replace(s, "&", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "_", " ")
    s = Replace(s, "(", " ")
    s = Replace(s, ")", " ")
    s = Replace(s, ",", " ")
    Dim raw As Variant: raw = Split(s, " ")
    Dim arr() As String, n As Long, i As Long, token As String
    ReDim arr(0 To 0)
    For i = LBound(raw) To UBound(raw)
        token = Trim$(CStr(raw(i)))
        If Len(token) >= 3 And Not FuzzyStopWord(token) Then
            If n = 0 Then
                ReDim arr(0 To 0)
            Else
                ReDim Preserve arr(0 To n)
            End If
            arr(n) = token
            n = n + 1
        End If
    Next i
    FuzzyTokenize = arr
End Function


Private Function FuzzyStopWord(ByVal token As String) As Boolean
    Select Case token
        Case "and", "the", "for", "from", "with", "used", "provided", "expense", "expenses"
            FuzzyStopWord = True
        Case Else
            FuzzyStopWord = False
    End Select
End Function


Private Function FuzzyHasNegative(ByVal conceptName As String, ByVal negTokens As Variant) As Boolean
    Dim i As Long, token As String
    For i = LBound(negTokens) To UBound(negTokens)
        token = Trim$(CStr(negTokens(i)))
        If Len(token) > 0 And InStr(1, conceptName, token, vbTextCompare) > 0 Then
            FuzzyHasNegative = True
            Exit Function
        End If
    Next i
End Function


Private Function FuzzyKeywordScore(ByVal conceptName As String, ByVal posTokens As Variant) As Double
    Dim i As Long, token As String, hits As Long, total As Long
    For i = LBound(posTokens) To UBound(posTokens)
        token = Trim$(CStr(posTokens(i)))
        If Len(token) > 0 Then
            total = total + 1
            If InStr(1, conceptName, token, vbTextCompare) > 0 Then hits = hits + 1
        End If
    Next i
    If total = 0 Or hits = 0 Then Exit Function
    If hits = total Then
        FuzzyKeywordScore = 5 + hits
    Else
        FuzzyKeywordScore = hits
    End If
End Function


Private Function FuzzyConceptHasUnit(ByVal conceptObj As Object, ByVal expectedUnit As String, ByRef unitText As String) As Boolean
    unitText = "—"
    If conceptObj Is Nothing Then Exit Function
    If Not conceptObj.Exists("units") Then Exit Function
    Dim units As Object: Set units = conceptObj.Item("units")
    If units.Exists(expectedUnit) Then
        unitText = expectedUnit
        FuzzyConceptHasUnit = True
        Exit Function
    End If
    Dim k As Variant
    For Each k In units.Keys
        unitText = CStr(k)
        Exit For
    Next k
End Function


Private Sub FuzzyInsertTop(ByRef topConcept() As String, ByRef topUnit() As String, ByRef topScore() As Double, _
                           ByVal conceptName As String, ByVal unitText As String, ByVal score As Double)
    Dim i As Long, j As Long
    For i = 1 To 3
        If score > topScore(i) Then
            For j = 3 To i + 1 Step -1
                topScore(j) = topScore(j - 1)
                topConcept(j) = topConcept(j - 1)
                topUnit(j) = topUnit(j - 1)
            Next j
            topScore(i) = score
            topConcept(i) = conceptName
            topUnit(i) = unitText
            Exit Sub
        End If
    Next i
End Sub


' --------- Phase 4b-14a: 公司整体 fetch 失败 (HTTP/JSON/etc) → 给该公司每个 conceptMap 指标
'           emit 一条 MISSING 诊断 (status=MISSING, 备注 = 失败原因)
'           供 RunUSStatement 在 mainErrNum<>0 路径上调用,使诊断 sheet 能完整反映"哪家公司哪张表全垮了"
Public Sub AddMissingDiagnosticsForCompany(ByVal ticker As String, ByVal strKind As String, _
                                            ByVal conceptMap As Variant, ByVal collRows As Collection, _
                                            ByVal noteReason As String)
    If collRows Is Nothing Then Exit Sub
    On Error Resume Next
    Dim ub As Long: ub = -1
    ub = UBound(conceptMap)
    On Error GoTo 0
    If ub < 0 Then Exit Sub

    Dim i As Long, mapEntry As Variant, lab As String
    For i = LBound(conceptMap) To ub
        mapEntry = conceptMap(i)
        On Error Resume Next
        lab = ""
        lab = MapEntryLabel(mapEntry)
        On Error GoTo 0
        If Len(lab) = 0 Then GoTo NextEntry
        AddDiagnosticRow collRows, ticker, strKind, lab, "MISSING", "—", "—", "—", "—", "—", noteReason
NextEntry:
    Next i
End Sub


' --------- 读取样本池 A4 单元格的季度选择 ---------
'   合法值: "全部" / "Q1" / "Q2" / "Q3" / "Q4"
'   读不到 / 空 / 异常 → 默认 "全部" (= 不过滤)
Public Function ReadQuarterSelection() As String
    On Error Resume Next
    Dim s As String
    s = Trim$(CStr(ThisWorkbook.Sheets("样本池").Range("A4").Value))
    If Err.Number <> 0 Or Len(s) = 0 Then
        ReadQuarterSelection = "全部"
        Err.Clear
    Else
        ReadQuarterSelection = s
    End If
    On Error GoTo 0
End Function


' --------- 读样本池 A2 年份选择 (0 = 留空 = 不过滤) ---------
Public Function ReadYearSelection() As Long
    On Error Resume Next
    Dim v As Variant
    v = ThisWorkbook.Sheets("样本池").Range("A2").Value
    If IsNumeric(v) Then
        ReadYearSelection = CLng(v)
    Else
        ReadYearSelection = 0
    End If
    On Error GoTo 0
End Function


' --------- 按季度过滤报告期字典 (in-place) ---------
'   strQuarter: 全部 / Q1 / Q2 / Q3 / Q4
'   "全部" → 不动 dictPeriodSet
'   Q1-Q4  → 只保留对应月末日期的 keys
Public Sub FilterPeriodsByQuarter(ByVal dictPeriodSet As Object, ByVal strQuarter As String)
    Dim suffix As String
    Select Case strQuarter
        Case "Q1": suffix = "-03-31"
        Case "Q2": suffix = "-06-30"
        Case "Q3": suffix = "-09-30"
        Case "Q4": suffix = "-12-31"
        Case Else: Exit Sub        ' 全部 / 未知 / 空 → 不过滤
    End Select

    Dim k As Variant, toRemove As Object
    Set toRemove = CreateObject("Scripting.Dictionary")
    For Each k In dictPeriodSet.Keys
        If Right$(CStr(k), Len(suffix)) <> suffix Then _
            toRemove.Add k, True
    Next k
    For Each k In toRemove.Keys
        dictPeriodSet.Remove k
    Next k
End Sub


' --------- HTTP 抓取 (Sina, gb2312 解码) ---------
Public Function HttpGet(ByVal strUrl As String) As String
    Dim objWinHttp As Object, arrByte() As Byte
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objWinHttp
        .SetTimeouts 30000, 30000, 30000, 30000
        .Open "GET", strUrl, False
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        .Send
        .WaitForResponse 30
        If .Status < 200 Or .Status >= 300 Then
            Err.Raise vbObjectError + 521, "HttpGet", _
                "HTTP " & .Status & " for " & strUrl
        End If
        arrByte = .ResponseBody
    End With
    HttpGet = ByteToStr(arrByte, "gb2312")
End Function


' --------- HTTP 抓取 (SEC EDGAR, UTF-8 解码) ---------
'   SEC fair-use 要求 User-Agent 含 name + email 标识请求方
'   Source: https://www.sec.gov/os/accessing-edgar-data
'   注意: 不能请求 gzip — WinHttpRequest 不会自动解压, 会拿到压缩字节
Public Function EdgarHttpGet(ByVal strUrl As String) As String
    Dim objWinHttp As Object, arrByte() As Byte
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objWinHttp
        .SetTimeouts 30000, 60000, 60000, 60000
        .Open "GET", strUrl, False
        .SetRequestHeader "User-Agent", "ListedCompanyFinancialData/1.0 (214978902@qq.com)"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Accept-Encoding", "identity"     ' 显式禁用压缩
        .Send
        .WaitForResponse 60
        If .Status = 404 Then
            Err.Raise vbObjectError + 526, "EdgarHttpGet", _
                "EDGAR 无数据 (404). 多见于 20-F filer (中概股/ADR 用 IFRS), SEC XBRL 接口不收录"
        ElseIf .Status < 200 Or .Status >= 300 Then
            Err.Raise vbObjectError + 526, "EdgarHttpGet", _
                "HTTP " & .Status & " for " & strUrl
        End If
        arrByte = .ResponseBody
    End With
    EdgarHttpGet = ByteToStr(arrByte, "utf-8")
End Function


' --------- Ticker → CIK 缓存加载 (一次会话内只下载一次 SEC tickers.json) ---------
'   tickers.json (~700 KB) 格式:
'     { "0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}, "1": {...}, ... }
'   返回 Dictionary: ticker(大写) → 10位补零的 CIK 字符串 (e.g. "0000320193")
'   全局变量 g_dictTickerToCIK 在文件顶部声明 (VBA 要求 Public var 在所有 procedure 前)
Public Function LoadTickerCIKMap() As Object
    If Not g_dictTickerToCIK Is Nothing Then
        Set LoadTickerCIKMap = g_dictTickerToCIK
        Exit Function
    End If

    Dim strJson As String
    strJson = EdgarHttpGetTickers()    ' tickers.json on www.sec.gov, not data.sec.gov
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(strJson)

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim k As Variant, entry As Object, cik As Long
    For Each k In parsed.Keys
        Set entry = parsed(k)
        cik = CLng(entry("cik_str"))
        dict.Item(UCase$(CStr(entry("ticker")))) = Format$(cik, "0000000000")
    Next k

    Set g_dictTickerToCIK = dict
    Set LoadTickerCIKMap = dict
End Function


' --------- HTTP 抓取 (雪球 Xueqiu, UTF-8 解码) ---------
'   雪球 API 需要登录后的 Cookie (用户从浏览器 F12 复制粘到 样本池!B5)
'   Source: https://xueqiu.com (登录后访问财报页面 → F12 → 拷 Cookie 头)
Public Function XueqiuHttpGet(ByVal strUrl As String, ByVal strCookie As String) As String
    Dim objWinHttp As Object, arrByte() As Byte
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objWinHttp
        .SetTimeouts 30000, 60000, 60000, 60000
        .Open "GET", strUrl, False
        .SetRequestHeader "User-Agent", _
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " & _
            "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        .SetRequestHeader "Accept", "application/json, text/plain, */*"
        .SetRequestHeader "Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8"
        .SetRequestHeader "Accept-Encoding", "identity"
        .SetRequestHeader "Referer", "https://xueqiu.com/"
        .SetRequestHeader "Origin", "https://xueqiu.com"
        If Len(strCookie) > 0 Then _
            .SetRequestHeader "Cookie", strCookie
        .Send
        .WaitForResponse 60
        If .Status < 200 Or .Status >= 300 Then
            Err.Raise vbObjectError + 535, "XueqiuHttpGet", _
                "HTTP " & .Status & " for " & strUrl
        End If
        arrByte = .ResponseBody
    End With
    XueqiuHttpGet = ByteToStr(arrByte, "utf-8")
End Function


' --------- 读样本池 B5 单元格的雪球 cookie ---------
'   用户在浏览器登录 xueqiu.com → F12 → Application → Cookies → 找 xq_a_token, 拷它的 value
'   粘到 B5; 也可以粘整段 Cookie 头 (含 xq_a_token=... 和别的 key)
'   - 单纯 token 值 (无 "=") → 自动包装成 "xq_a_token=<value>"
'   - 已含 "=" → 当成完整 Cookie 头用
Public Function ReadXueqiuCookie() As String
    On Error Resume Next
    Dim s As String
    s = Trim$(CStr(ThisWorkbook.Sheets("样本池").Range("B5").Value))
    If Len(s) = 0 Then
        ReadXueqiuCookie = ""
    ElseIf InStr(s, "=") = 0 Then
        ReadXueqiuCookie = "xq_a_token=" & s
    Else
        ReadXueqiuCookie = s
    End If
    Err.Clear
    On Error GoTo 0
End Function


' --------- Phase 4f Step 2: 读样本池 B6 显示币种切换 ---------
'   返回 "原币" (默认) 或 "统一RMB"
'   B6 不存在 / 空 / sheet 不存在 → "原币"; 用户改 B6 立即生效
Public Function ReadDisplayCurrency() As String
    On Error Resume Next
    Dim s As String
    s = Trim$(CStr(ThisWorkbook.Sheets("样本池").Range("B6").Value))
    If Err.Number <> 0 Or Len(s) = 0 Then s = "原币"
    Err.Clear
    On Error GoTo 0
    ReadDisplayCurrency = s
End Function


' --------- Phase 4f Step 2: 通用汇率查询 (供 Step 4 写表时调用) ---------
'   curCode   : "RMB" / "CNY" / "USD" / "HKD" / "KRW" (其他返 0)
'   periodEnd : 报告期 yyyy-mm-dd, 必须与 汇率 sheet A 列文本一致
'   useEop    : True=期末 (BS 用); False=期间均值 (IS/CF 用)
'
'   返回:
'     RMB / CNY → 1 (不查 sheet, 不抓数)
'     命中缓存  → 缓存值
'     未命中    → 调 模块_抓汇率.EnsureFxRateCached 拉数, 写 sheet, 再读
'     失败      → 0 (Step 4 调用方需 fallback 到 1#)
'
'   用户可手填 汇率 sheet 单元格 override 系统值; 只要 IsNumeric And > 0 就保留
'   注: 参数名避免用 "currency" (与 VBA 数据类型 Currency 冲突可能引起编译歧义)
Public Function GetFxRate(ByVal curCode As String, ByVal periodEnd As String, ByVal useEop As Boolean) As Double
    Dim c As String: c = UCase$(Trim$(curCode))
    If c = "RMB" Or c = "CNY" Then
        GetFxRate = 1#
        Exit Function
    End If

    Dim col As Long: col = LookupFxColForCurrency(c, useEop)
    If col = 0 Then
        GetFxRate = 0
        Exit Function
    End If

    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FX_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        GetFxRate = 0
        Exit Function
    End If

    Dim rowIdx As Long
    rowIdx = LookupFxRowByPeriod(ws, periodEnd)
    Dim v As Variant
    If rowIdx > 0 Then
        v = ws.Cells(rowIdx, col).Value
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then
                GetFxRate = CDbl(v)
                Exit Function
            End If
        End If
    End If

    ' Cache miss → 拉数. 跨模块调用走 Application.Run (dot syntax 对中文模块名编译不通)
    Dim ok As Boolean
    On Error Resume Next
    ok = CBool(Application.Run("模块_抓汇率.EnsureFxRateCached", periodEnd, c))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        GetFxRate = 0
        Exit Function
    End If
    On Error GoTo 0

    If Not ok Then
        GetFxRate = 0
        Exit Function
    End If

    ' 拉完再读
    rowIdx = LookupFxRowByPeriod(ws, periodEnd)
    If rowIdx = 0 Then
        GetFxRate = 0
        Exit Function
    End If
    v = ws.Cells(rowIdx, col).Value
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            GetFxRate = CDbl(v)
            Exit Function
        End If
    End If
    GetFxRate = 0
End Function


' --------- 汇率 sheet A 列查报告期所在行;0 表示未找到 ---------
Private Function LookupFxRowByPeriod(ByVal ws As Object, ByVal periodEnd As String) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row    ' xlUp
    If lastRow < FX_DATA_ROW Then
        LookupFxRowByPeriod = 0
        Exit Function
    End If
    Dim target As String: target = Trim$(periodEnd)
    For r = FX_DATA_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) = target Then
            LookupFxRowByPeriod = r
            Exit Function
        End If
    Next r
    LookupFxRowByPeriod = 0
End Function


' --------- 汇率 sheet 列号映射 (curCode × eop/avg) ---------
'   B=USDCNY期末(2)  C=USDCNY期均(3)
'   D=HKDCNY期末(4)  E=HKDCNY期均(5)
'   F=KRWCNY期末(6)  G=KRWCNY期均(7)
'   未知 curCode → 0
Public Function LookupFxColForCurrency(ByVal curCode As String, ByVal useEop As Boolean) As Long
    Dim c As String: c = UCase$(Trim$(curCode))
    Select Case c
        Case "USD"
            LookupFxColForCurrency = IIf(useEop, 2, 3)
        Case "HKD"
            LookupFxColForCurrency = IIf(useEop, 4, 5)
        Case "KRW"
            LookupFxColForCurrency = IIf(useEop, 6, 7)
        Case Else
            LookupFxColForCurrency = 0
    End Select
End Function


' --------- 同 EdgarHttpGet 但目标主机是 www.sec.gov (tickers.json) ---------
Private Function EdgarHttpGetTickers() As String
    Dim objWinHttp As Object, arrByte() As Byte
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objWinHttp
        .SetTimeouts 30000, 30000, 30000, 30000
        .Open "GET", "https://www.sec.gov/files/company_tickers.json", False
        .SetRequestHeader "User-Agent", "ListedCompanyFinancialData/1.0 (214978902@qq.com)"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Accept-Encoding", "identity"
        .Send
        .WaitForResponse 30
        If .Status < 200 Or .Status >= 300 Then
            Err.Raise vbObjectError + 527, "EdgarHttpGetTickers", _
                "HTTP " & .Status
        End If
        arrByte = .ResponseBody
    End With
    EdgarHttpGetTickers = ByteToStr(arrByte, "utf-8")
End Function


' --------- 按 ticker 查 CIK (含 "0000320193" 格式) ---------
Public Function LookupCIK(ByVal strTicker As String) As String
    Dim dict As Object: Set dict = LoadTickerCIKMap()
    Dim t As String: t = UCase$(Trim$(strTicker))
    If dict.Exists(t) Then
        LookupCIK = dict(t)
    Else
        Err.Raise vbObjectError + 528, "LookupCIK", "未找到 ticker: " & strTicker
    End If
End Function


' --------- 按代码字符特征推断市场 ---------
'   返回 "A" / "HK" / "US"; KR 不自动推断, 需用户在样本池 C 列手填
Public Function DetectMarket(ByVal strCode As String) As String
    Dim s As String: s = UCase$(Trim$(strCode))
    If Len(s) = 0 Then DetectMarket = "A": Exit Function

    ' 含字母 → US
    Dim i As Long, ch As String, hasLetter As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "A" And ch <= "Z" Then hasLetter = True
    Next i
    If hasLetter Then
        DetectMarket = "US"
    ElseIf Len(s) = 5 Then
        DetectMarket = "HK"      ' 5 位数字 → 港股
    Else
        DetectMarket = "A"       ' 其他默认 A 股 (6 位 / 短于 5 位)
    End If
End Function


' --------- 综合: 给定样本池一行 (代码, H 列市场), 返回最终市场 ---------
'   H 列优先 (用户显式指定); 否则按 ticker 推断
Public Function ResolveMarket(ByVal strCode As String, ByVal strHColMarket As String) As String
    Dim h As String: h = UCase$(Trim$(strHColMarket))
    Select Case h
        Case "A", "HK", "US", "KR"
            ResolveMarket = h
        Case Else
            ResolveMarket = DetectMarket(strCode)
    End Select
End Function


Private Function ByteToStr(arrByte, ByVal strCharSet As String) As String
    With CreateObject("Adodb.Stream")
        .Type = 1            ' adTypeBinary
        .Open
        .Write arrByte
        .Position = 0
        .Type = 2            ' adTypeText
        .Charset = strCharSet
        ByteToStr = .ReadText
        .Close
    End With
End Function


' --------- 截取 <table id="..."> 片段 ---------
'  用 InStr 而非 regex, 避免 .+ 在 80KB HTML 上的回溯爆炸
Public Function ExtractTable(ByVal strText As String, ByVal strID As String) As String
    ' 先把所有 \s+ 塌成单空格, 让后续 htmlfile DOM 解析更干净
    Dim objRegx As Object
    Set objRegx = CreateObject("VBScript.Regexp")
    objRegx.Global = True
    objRegx.Pattern = "\s+"
    strText = objRegx.Replace(strText, " ")

    Dim startTag As String, endTag As String
    startTag = "<table id=""" & strID & """"
    endTag = "</table>"

    Dim posStart As Long, posEnd As Long
    posStart = InStr(1, strText, startTag, vbTextCompare)
    If posStart = 0 Then
        Err.Raise vbObjectError + 522, "ExtractTable", "未找到 table id=" & strID
    End If
    posEnd = InStr(posStart + Len(startTag), strText, endTag, vbTextCompare)
    If posEnd = 0 Then
        Err.Raise vbObjectError + 522, "ExtractTable", "未找到 </table> for " & strID
    End If
    ExtractTable = Mid$(strText, posStart, posEnd - posStart + Len(endTag))
End Function


' --------- 解析单家公司的财报 HTML ---------
'   累计到调用方传入的 4 个共享字典:
'     dictData         : code -> period -> indicator -> value
'     dictPeriodSet    : period(YYYY-MM-DD) -> True   (只用 keys 当 set)
'     dictIndicatorSet : indicator -> ordinal         (保持插入顺序)
'     dictCategoryMap  : indicator -> category        (大类映射)
'
'   strID 是 HTML 里的 <table id="..."> 值
'   strHtml 是已经 ExtractTable 过的 <table>...</table> 片段
'   objHtml 是复用的 htmlfile DOM 对象 (调用方创建一次, 多公司共用)
Public Sub ParseFinancialHtml(ByVal strHtml As String, ByVal strID As String, _
                               ByVal strCode As String, _
                               ByVal objHtml As Object, _
                               ByRef dictData As Object, _
                               ByRef dictPeriodSet As Object, _
                               ByRef dictIndicatorSet As Object, _
                               ByRef dictCategoryMap As Object)
    Dim objTb As Object, objRow As Object
    Dim arrDates() As String, intPeriodCnt As Long, j As Long
    Dim blnDateLoaded As Boolean
    Dim strCategory As String, strIndicator As String
    Dim dictCompany As Object, dictPeriod As Object
    Dim varVal As Variant

    objHtml.body.innerHTML = strHtml
    Set objTb = objHtml.getElementById(strID)
    If objTb Is Nothing Then
        Err.Raise vbObjectError + 523, "ParseFinancialHtml", "未找到 table id=" & strID
    End If

    ' 给当前公司创建子字典
    If Not dictData.Exists(strCode) Then
        Set dictCompany = CreateObject("Scripting.Dictionary")
        dictData.Add strCode, dictCompany
    Else
        Set dictCompany = dictData(strCode)
    End If

    For Each objRow In objTb.Rows
        If objRow.Cells.Length = 1 Then
            ' 单 cell 行 = 大类标题
            strCategory = Trim$(objRow.Cells(0).innerText)
        ElseIf objRow.Cells.Length > 1 Then
            If Not blnDateLoaded Then
                ' 第一行多 cell = 报告期表头
                intPeriodCnt = objRow.Cells.Length - 1
                ReDim arrDates(1 To intPeriodCnt)
                For j = 1 To intPeriodCnt
                    arrDates(j) = Trim$(objRow.Cells(j).innerText)
                    If Not dictPeriodSet.Exists(arrDates(j)) Then _
                        dictPeriodSet.Add arrDates(j), True
                Next j
                blnDateLoaded = True
            Else
                ' 数据行: cell(0) = 指标名, cell(1..N) = 数值
                strIndicator = Trim$(objRow.Cells(0).innerText)
                If Len(strIndicator) > 0 Then
                    If Not dictIndicatorSet.Exists(strIndicator) Then
                        dictIndicatorSet.Add strIndicator, dictIndicatorSet.Count
                        dictCategoryMap.Add strIndicator, strCategory
                    End If

                    For j = 1 To intPeriodCnt
                        If j < objRow.Cells.Length Then
                            varVal = NormalizeValue(objRow.Cells(j).innerText)
                            If Not IsEmpty(varVal) Then
                                If Not dictCompany.Exists(arrDates(j)) Then _
                                    dictCompany.Add arrDates(j), CreateObject("Scripting.Dictionary")
                                Set dictPeriod = dictCompany(arrDates(j))
                                If Not dictPeriod.Exists(strIndicator) Then _
                                    dictPeriod.Add strIndicator, varVal
                            End If
                        End If
                    Next j
                End If
            End If
        End If
    Next objRow

    If Not blnDateLoaded Then
        Err.Raise vbObjectError + 524, "ParseFinancialHtml", _
            "未解析到报告期表头: " & strCode
    End If
End Sub


' --------- 单元格值标准化: -- 转空, 数字转 Double, 去千分位逗号 ---------
Public Function NormalizeValue(ByVal strRaw As String) As Variant
    Dim s As String
    s = Trim$(strRaw)
    If Len(s) = 0 Or s = "--" Or s = "-" Then
        NormalizeValue = Empty
        Exit Function
    End If
    ' 新浪 HTML 数字带千分位逗号 (e.g. "365,650.93"), CDbl 在某些 locale 不识别
    Dim sNum As String: sNum = Replace(s, ",", "")
    If IsNumeric(sNum) Then
        NormalizeValue = CDbl(sNum)
    Else
        NormalizeValue = s
    End If
End Function


' --------- 解析公司基本资料 (comInfo1) ---------
'   返回字典 keys: 公司名称 / 上市日期 / 所属行业 / 主营业务
Public Function ParseCorpInfoHtml(ByVal strHtml As String, ByVal objHtml As Object) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim objTb As Object, objRow As Object, objCell As Object
    Dim strLabel As String, strValue As String
    Dim i As Long

    objHtml.body.innerHTML = strHtml
    Set objTb = objHtml.getElementById("comInfo1")
    If objTb Is Nothing Then
        Err.Raise vbObjectError + 525, "ParseCorpInfoHtml", "未找到 comInfo1"
    End If

    For Each objRow In objTb.Rows
        For i = 0 To objRow.Cells.Length - 2 Step 2
            strLabel = Trim$(objRow.Cells(i).innerText)
            strValue = Trim$(objRow.Cells(i + 1).innerText)
            If Len(strLabel) > 0 And Not dict.Exists(strLabel) Then _
                dict.Add strLabel, strValue
        Next i
    Next objRow

    Set ParseCorpInfoHtml = dict
End Function


' --------- 写宽表 ---------
'   ws                : 目标 Worksheet
'   arrCodes          : 公司代码数组 (一维, 1-based, 顺序按样本池)
'   dictCompanyName   : code -> 简称
'   dictData          : code -> period -> indicator -> value (already populated)
'   arrPeriodsSorted  : 报告期数组 (一维, 1-based, 已按降序排好)
'   arrIndicators     : 指标名数组 (一维, 1-based, 已并集排序)
'   dictCategory      : indicator -> 大类
'   perCompanyPeriods : True 时每家公司只展开自己有数据的报告期 (美股用)
'   dictReportingCurrency : code -> 报告币种; Nothing 时按 RMB 处理
'   statementKind     : "BalanceSheet" / "Income" / "CashFlow", 用于选择期末/均值汇率
Public Sub WriteWideTable(ByVal ws As Worksheet, _
                           ByRef arrCodes As Variant, _
                           ByRef dictCompanyName As Object, _
                           ByRef dictData As Object, _
                           ByRef arrPeriodsSorted As Variant, _
                           ByRef arrIndicators As Variant, _
                           ByRef dictCategory As Object, _
                           Optional ByVal perCompanyPeriods As Boolean = False, _
                           Optional ByRef dictReportingCurrency As Object = Nothing, _
                           Optional ByVal statementKind As String = "")
    Dim numCompanies As Long, numPeriods As Long, numIndicators As Long
    Dim i As Long, j As Long, k As Long, intRow As Long, intCol As Long
    Dim strCode As String, strName As String, strInd As String, strPeriod As String
    Dim varValue As Variant

    ' Phase 4f Step 3: RMB 换算预读, 避免内循环反复读样本池 B6
    Dim displayMode As String: displayMode = ReadDisplayCurrency()
    Dim useEopForBS As Boolean
    useEopForBS = (UCase$(Trim$(statementKind)) = "BALANCESHEET")

    numCompanies = UBound(arrCodes) - LBound(arrCodes) + 1
    numPeriods = UBound(arrPeriodsSorted) - LBound(arrPeriodsSorted) + 1
    numIndicators = UBound(arrIndicators) - LBound(arrIndicators) + 1

    Dim metaCols As Long: metaCols = 2
    If ws.Name = "A股_指标表" Or ws.Name = "美股_指标表" Or _
       ws.Name = "港股_指标表" Or ws.Name = "韩股_指标表" Then metaCols = 3

    Dim companyPeriods As Object: Set companyPeriods = CreateObject("Scripting.Dictionary")
    Dim companyStartCols As Object: Set companyStartCols = CreateObject("Scripting.Dictionary")
    Dim totalDataCols As Long
    If perCompanyPeriods Then
        For i = 1 To numCompanies
            strCode = CStr(arrCodes(i))
            Dim collPeriods As Collection: Set collPeriods = New Collection
            If dictData.Exists(strCode) Then
                Dim dictCompanyForPeriods As Object: Set dictCompanyForPeriods = dictData(strCode)
                For j = 1 To numPeriods
                    strPeriod = CStr(arrPeriodsSorted(j))
                    If dictCompanyForPeriods.Exists(strPeriod) Then collPeriods.Add strPeriod
                Next j
            End If
            companyPeriods.Add strCode, collPeriods
            totalDataCols = totalDataCols + collPeriods.Count
        Next i
    Else
        totalDataCols = numCompanies * numPeriods
    End If

    Dim totalCols As Long: totalCols = metaCols + totalDataCols
    If totalCols < metaCols + 1 Then totalCols = metaCols + 1

    ' 1) 清旧数据 (保留 A1:B1 容器, 但 R1 公司名 + R2 + R3+ 数据全部重画)
    With ws
        ' 解除上次跑产生的合并 (含 R1 跨期合并)
        On Error Resume Next
        ws.UsedRange.UnMerge
        On Error GoTo 0

        Dim lastRow As Long, lastCol As Long
        lastRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
        lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        If lastRow < 2 Then lastRow = 2
        If lastCol < metaCols + 1 Then lastCol = metaCols + 1

        ' 数据列到末行末列全清
        .Range(.Cells(1, metaCols + 1), .Cells(lastRow, lastCol)).Clear
        ' A2:<metaCols><lastRow> 全清 (保留第 1 行静态表头)
        If lastRow >= 2 Then _
            .Range(.Cells(2, 1), .Cells(lastRow, metaCols)).Clear

        ' 2) 重新画静态表头
        If metaCols = 3 Then
            .Range("A1").Value = "指标类型"
            .Range("C1").Value = "英文指标名"
        Else
            .Range("A1").Value = "大类"
        End If
        .Range("B1").Value = "指标名称"
        With .Range(.Cells(1, 1), .Cells(1, metaCols))
            .Font.Name = "微软雅黑"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)    ' 4472C4
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' 3) 写 R1 (公司名, 跨期合并) + R2 (报告期)
        Dim currentCol As Long: currentCol = metaCols + 1
        For i = 1 To numCompanies
            strCode = CStr(arrCodes(i))
            If dictCompanyName.Exists(strCode) Then
                strName = CStr(dictCompanyName(strCode)) & "(" & strCode & ")"
            Else
                strName = strCode
            End If
            ' Phase 4f Step 5: 统一RMB 模式下追加原始报告币种 tag
            If displayMode = "统一RMB" And Not dictReportingCurrency Is Nothing Then
                If dictReportingCurrency.Exists(strCode) Then
                    Dim origCur As String: origCur = CStr(dictReportingCurrency(strCode))
                    If origCur <> "RMB" And Len(origCur) > 0 Then
                        strName = strName & " [" & origCur & "→RMB]"
                    End If
                End If
            End If

            Dim periodsForCompany As Collection
            Dim periodCount As Long
            If perCompanyPeriods Then
                Set periodsForCompany = companyPeriods(strCode)
                periodCount = periodsForCompany.Count
            Else
                periodCount = numPeriods
            End If
            If periodCount = 0 Then GoTo NextHeaderCompany

            intCol = currentCol
            companyStartCols(strCode) = intCol
            .Cells(1, intCol).Value = strName

            ' 合并 R1 这家公司占的 N 列
            If periodCount > 1 Then
                .Range(.Cells(1, intCol), .Cells(1, intCol + periodCount - 1)).Merge
            End If
            With .Cells(1, intCol)
                .Font.Name = "微软雅黑"
                .Font.Size = 11
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(68, 114, 196)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

            ' R2 报告期, 已按降序
            For j = 1 To periodCount
                If perCompanyPeriods Then
                    strPeriod = CStr(periodsForCompany.Item(j))
                Else
                    strPeriod = CStr(arrPeriodsSorted(j))
                End If
                With .Cells(2, intCol + j - 1)
                    On Error Resume Next
                    .Value = CDate(strPeriod)
                    If Err.Number <> 0 Then
                        .Value = strPeriod
                        Err.Clear
                    End If
                    On Error GoTo 0
                    .NumberFormat = "yyyy-mm-dd"
                    .Font.Name = "微软雅黑"
                    .Font.Size = 10
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(68, 114, 196)
                    .HorizontalAlignment = xlCenter
                End With
            Next j
            currentCol = currentCol + periodCount
NextHeaderCompany:
        Next i

        ' 4) 数据行 R3+
        '    用 Variant 二维数组一次性写入, 比逐 cell 快
        Dim arrOut As Variant
        ReDim arrOut(1 To numIndicators, 1 To totalCols)
        For k = 1 To numIndicators
            strInd = CStr(arrIndicators(k))
            If dictCategory.Exists(strInd) Then arrOut(k, 1) = dictCategory(strInd)
            arrOut(k, 2) = strInd

            For i = 1 To numCompanies
                strCode = CStr(arrCodes(i))
                If dictData.Exists(strCode) Then
                    Dim dictCompany As Object: Set dictCompany = dictData(strCode)
                    If perCompanyPeriods Then
                        If Not companyStartCols.Exists(strCode) Then GoTo NextDataCompany
                        Set periodsForCompany = companyPeriods(strCode)
                        intCol = CLng(companyStartCols(strCode))
                        periodCount = periodsForCompany.Count
                    Else
                        intCol = metaCols + 1 + (i - 1) * numPeriods
                        periodCount = numPeriods
                    End If
                    For j = 1 To periodCount
                        If perCompanyPeriods Then
                            strPeriod = CStr(periodsForCompany.Item(j))
                        Else
                            strPeriod = CStr(arrPeriodsSorted(j))
                        End If
                        If dictCompany.Exists(strPeriod) Then
                            Dim dictPer As Object: Set dictPer = dictCompany(strPeriod)
                            If dictPer.Exists(strInd) Then
                                Dim rawVal As Variant: rawVal = dictPer(strInd)
                                Dim writeVal As Variant: writeVal = rawVal
                                ' Phase 4f Step 3: 统一RMB 显示模式 -> 按报告币种乘汇率
                                If displayMode = "统一RMB" And IsNumeric(rawVal) Then
                                    Dim curCode As String: curCode = "RMB"
                                    If Not dictReportingCurrency Is Nothing Then
                                        If dictReportingCurrency.Exists(strCode) Then _
                                            curCode = CStr(dictReportingCurrency(strCode))
                                    End If
                                    Dim fx As Double
                                    fx = GetFxRate(curCode, strPeriod, useEopForBS)
                                    If fx > 0 Then
                                        writeVal = CDbl(rawVal) * fx
                                    Else
                                        writeVal = rawVal
                                    End If
                                End If
                                arrOut(k, intCol + j - 1) = writeVal
                            End If
                        End If
                    Next j
                End If
NextDataCompany:
            Next i
        Next k

        .Range(.Cells(3, 1), .Cells(2 + numIndicators, totalCols)).Value = arrOut

        ' 5) 列宽 / 字体 / 边框 / 冻结
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 40
        If metaCols = 3 Then .Columns("C").ColumnWidth = 32
        Dim rngData As Range
        Set rngData = .Range(.Cells(3, 1), .Cells(2 + numIndicators, totalCols))
        With rngData
            .Font.Name = "微软雅黑"
            .Font.Size = 10
        End With
        With .Range(.Cells(3, 1), .Cells(2 + numIndicators, metaCols))
            .Font.Bold = True
        End With
        With .Range(.Cells(3, metaCols + 1), .Cells(2 + numIndicators, totalCols))
            .NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
            .HorizontalAlignment = xlRight
        End With

        Dim rngAll As Range
        Set rngAll = .Range(.Cells(1, 1), .Cells(2 + numIndicators, totalCols))
        Call SetBorderLine(rngAll)

        ' 设置 C 列起的列宽
        Dim col As Long
        For col = metaCols + 1 To totalCols
            .Columns(col).ColumnWidth = 15.875
        Next col

        ' 行高
        .Rows(1).RowHeight = 22
        .Rows(2).RowHeight = 20

        ' 冻结数据区左上角 (前 2 行 + 静态列锚定)
        ws.Activate
        ActiveWindow.FreezePanes = False
        .Cells(3, metaCols + 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
    End With
End Sub


Public Function UnitDescriptionForMarket(ByVal sheetName As String) As String
    Select Case True
        Case InStr(sheetName, "A股") > 0
            UnitDescriptionForMarket = "百万 RMB (新浪财报源)"
        Case InStr(sheetName, "美股") > 0
            UnitDescriptionForMarket = "百万 USD (EDGAR)"
        Case InStr(sheetName, "港股") > 0
            UnitDescriptionForMarket = "百万 (各家公司报告币种, 见 港股_抓取诊断 Unit/FX_Rate 列)"
        Case InStr(sheetName, "韩股") > 0
            UnitDescriptionForMarket = "十亿 KRW (stockanalysis)"
        Case Else
            UnitDescriptionForMarket = "(单位见诊断 sheet)"
    End Select
End Function


Public Sub RefreshA1CurrencyComment(ByVal wsTarget As Worksheet, ByVal targetSheet As String)
    Dim displayMode As String: displayMode = ReadDisplayCurrency()
    On Error Resume Next
    If Not wsTarget.Range("A1").Comment Is Nothing Then wsTarget.Range("A1").Comment.Delete
    On Error GoTo 0

    Dim commentText As String
    If displayMode = "统一RMB" Then
        commentText = "单位: 百万 RMB (统一汇率换算; 汇率源见『汇率』sheet, 期末/期间均值混合)" & vbCrLf & _
                      "切回原币: 样本池 B6 改为 '原币' 后重跑"
    Else
        commentText = "单位: " & UnitDescriptionForMarket(targetSheet) & vbCrLf & _
                      "统一显示 RMB: 样本池 B6 改为 '统一RMB' 后重跑"
    End If

    wsTarget.Range("A1").AddComment commentText
    wsTarget.Range("A1").Comment.Shape.TextFrame.AutoSize = True
End Sub


Public Function FxRateTextForDiagnostic(ByVal reportingCurrency As String, _
                                        ByVal periodEnd As String, _
                                        ByVal strKind As String) As String
    FxRateTextForDiagnostic = "1.0"
    If ReadDisplayCurrency() <> "统一RMB" Then Exit Function
    If Len(Trim$(periodEnd)) = 0 Then Exit Function

    Dim curCode As String: curCode = Trim$(reportingCurrency)
    If Len(curCode) = 0 Then curCode = "RMB"

    Dim useEopFlag As Boolean
    useEopFlag = (UCase$(Trim$(strKind)) = "BALANCESHEET")

    Dim fxNum As Double
    fxNum = GetFxRate(curCode, periodEnd, useEopFlag)
    If fxNum > 0 Then FxRateTextForDiagnostic = Format$(fxNum, "0.000000")
End Function


' --------- 生成只含 18 个标准指标的指标表 ---------
'   market: "A" / "US" / "HK" / "KR"; 从对应资产负债表复制公司/报告期表头, 再填标准指标公式
Public Sub BuildStandardIndicatorSheet(ByVal market As String)
    Dim marketKey As String: marketKey = UCase$(Trim$(market))
    Dim targetSheet As String, sourceSheet As String
    If marketKey = "US" Then
        targetSheet = "美股_指标表"
        sourceSheet = "美股_资产负债表"
    ElseIf marketKey = "HK" Then
        targetSheet = "港股_指标表"
        sourceSheet = "港股_资产负债表"
    ElseIf marketKey = "KR" Then
        targetSheet = "韩股_指标表"
        sourceSheet = "韩股_资产负债表"
    Else
        marketKey = "A"
        targetSheet = "A股_指标表"
        sourceSheet = "A股_资产负债表"
    End If

    Dim wsTarget As Worksheet, wsSource As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)
    Set wsSource = ThisWorkbook.Sheets(sourceSheet)

    Dim sourceStartCol As Long: sourceStartCol = StandardDataStartCol(wsSource)
    Dim sourceLastCol As Long
    sourceLastCol = wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column
    If sourceLastCol < sourceStartCol Then
        Err.Raise vbObjectError + 560, "BuildStandardIndicatorSheet", _
            sourceSheet & " 还没有数据. 请先更新资产负债表和利润表"
    End If

    Application.ScreenUpdating = False
    On Error Resume Next
    wsTarget.UsedRange.UnMerge
    On Error GoTo 0
    wsTarget.Cells.Clear

    wsTarget.Range("A1").Value = "指标类型"
    wsTarget.Range("B1").Value = "指标名称"
    wsTarget.Range("C1").Value = "英文指标名"
    With wsTarget.Range("A1:C1")
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim targetCol As Long: targetCol = 4
    Dim c As Long, periodText As String, selectedYear As Long
    Dim quarterPick As String: quarterPick = ReadQuarterSelection()
    selectedYear = ReadYearSelection()

    For c = sourceStartCol To sourceLastCol
        periodText = StandardPeriodKey(wsSource.Cells(2, c).Value)
        If StandardTargetPeriodWanted(periodText, selectedYear, quarterPick, _
                                      marketKey, wsSource, StandardHeaderTextAt(wsSource, c), c) Then
            wsTarget.Cells(1, targetCol).Value = StandardHeaderTextAt(wsSource, c)
            wsTarget.Cells(2, targetCol).Value = wsSource.Cells(2, c).Value
            wsTarget.Cells(2, targetCol).NumberFormat = "yyyy-mm-dd"
            targetCol = targetCol + 1
        End If
    Next c

    If targetCol = 4 Then
        Err.Raise vbObjectError + 561, "BuildStandardIndicatorSheet", _
            sourceSheet & " 没有匹配当前 A2/A4 选择的报告期"
    End If

    Dim lastCol As Long: lastCol = targetCol - 1
    Call MergeStandardCompanyHeaders(wsTarget, 4, lastCol)

    With wsTarget.Range(wsTarget.Cells(1, 4), wsTarget.Cells(2, lastCol))
        .Font.Name = "微软雅黑"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    wsTarget.Columns("A").ColumnWidth = 18
    wsTarget.Columns("B").ColumnWidth = 28
    wsTarget.Columns("C").ColumnWidth = 34
    For c = 4 To lastCol
        wsTarget.Columns(c).ColumnWidth = 15.875
    Next c
    wsTarget.Rows(1).RowHeight = 22
    wsTarget.Rows(2).RowHeight = 20

    AppendStandardIndicators wsTarget, marketKey

    wsTarget.Activate
    ActiveWindow.FreezePanes = False
    wsTarget.Cells(3, 4).Select
    ActiveWindow.FreezePanes = True
    wsTarget.Cells(1, 1).Select
    RefreshA1CurrencyComment wsTarget, targetSheet
    Application.ScreenUpdating = True
End Sub


' --------- 指标表追加标准指标层 ---------
'   market: "A" / "US" / "HK" / "KR"
'   依赖对应市场的资产负债表、利润表已经生成; 缺字段时公式留空
Public Sub AppendStandardIndicators(ByVal ws As Worksheet, ByVal market As String)
    Dim marketKey As String: marketKey = UCase$(Trim$(market))
    If marketKey <> "A" And marketKey <> "US" And marketKey <> "HK" And marketKey <> "KR" Then Exit Sub

    Dim dataStartCol As Long: dataStartCol = StandardDataStartCol(ws)
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < dataStartCol Then Exit Sub

    Dim defs As Variant: defs = StandardIndicatorDefs()
    Dim stdCount As Long
    stdCount = UBound(defs) - LBound(defs) + 1

    ' 把抓取回来的原始指标下移, 标准指标固定展示在最上方
    Dim rawLastRow As Long
    rawLastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If rawLastRow >= 3 Then
        Dim rawData As Variant
        rawData = ws.Range(ws.Cells(3, 1), ws.Cells(rawLastRow, lastCol)).Value
        ws.Range(ws.Cells(3 + stdCount, 1), _
                 ws.Cells(rawLastRow + stdCount, lastCol)).Value = rawData
        ws.Range(ws.Cells(3, 1), ws.Cells(2 + stdCount, lastCol)).Clear
    End If

    Dim finalLastRow As Long: finalLastRow = rawLastRow + stdCount
    If finalLastRow < 2 + stdCount Then finalLastRow = 2 + stdCount
    With ws.Range(ws.Cells(3, 1), ws.Cells(finalLastRow, lastCol))
        .Font.Name = "微软雅黑"
        .Font.Size = 10
    End With
    With ws.Range(ws.Cells(3, 1), ws.Cells(finalLastRow, 3))
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(3, dataStartCol), ws.Cells(finalLastRow, lastCol))
        .NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
        .HorizontalAlignment = xlRight
    End With

    Dim rowMap As Object: Set rowMap = StandardRowMap(marketKey)
    Dim i As Long, rowNum As Long, c As Long
    For i = LBound(defs) To UBound(defs)
        rowNum = 3 + (i - LBound(defs))
        ws.Cells(rowNum, 1).Value = CStr(defs(i)(0))
        ws.Cells(rowNum, 2).Value = CStr(defs(i)(1))
        ws.Cells(rowNum, 3).Value = CStr(defs(i)(2))

        For c = dataStartCol To lastCol
            Dim companyHeader As String: companyHeader = StandardHeaderTextAt(ws, c)
            Dim periodText As String: periodText = StandardPeriodKey(ws.Cells(2, c).Value)
            Dim formulaText As String
            formulaText = StandardIndicatorFormula(CStr(defs(i)(3)), marketKey, rowMap, _
                                                   companyHeader, periodText, ws.Name, rowNum, c)
            If Len(formulaText) > 0 Then
                ws.Cells(rowNum, c).Formula = formulaText
                ws.Cells(rowNum, c).NumberFormat = CStr(defs(i)(4))
            Else
                ws.Cells(rowNum, c).Value = ""
            End If
        Next c
    Next i

    With ws.Range(ws.Cells(3, 1), ws.Cells(2 + stdCount, 3))
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(3, dataStartCol), ws.Cells(2 + stdCount, lastCol))
        .HorizontalAlignment = xlRight
    End With

    Call SetBorderLine(ws.Range(ws.Cells(1, 1), ws.Cells(finalLastRow, lastCol)))
End Sub


Private Function StandardIndicatorDefs() As Variant
    Dim a(0 To 17) As Variant
    a(0) = Array("盈利性指标", "销售净利率", "Net Profit Margin", "NPM", "0.00%")
    a(1) = Array("盈利性指标", "毛利率", "Gross Profit Margin", "GPM", "0.00%")
    a(2) = Array("盈利性指标", "期间费用率", "Operating Expense Ratio", "OER", "0.00%")
    a(3) = Array("盈利性指标", "总资产回报率 (ROA)", "Return on Assets (ROA)", "ROA", "0.00%")
    a(4) = Array("盈利性指标", "股东权益回报率 (ROE)", "Return on Equity (ROE)", "ROE", "0.00%")
    a(5) = Array("成长性指标", "总资产增长率", "Total Assets Growth Rate", "TAGR", "0.00%")
    a(6) = Array("成长性指标", "主营业务收入增长率", "Revenue Growth Rate", "RGR", "0.00%")
    a(7) = Array("成长性指标", "净利润增长率", "Net Profit Growth Rate", "NPGR", "0.00%")
    a(8) = Array("偿债能力指标", "流动比率", "Current Ratio", "CR", "0.00")
    a(9) = Array("偿债能力指标", "速动比率", "Quick Ratio", "QR", "0.00")
    a(10) = Array("偿债能力指标", "现金比率", "Cash Ratio", "CASHR", "0.00")
    a(11) = Array("偿债能力指标", "资产负债率", "Debt-to-Asset Ratio", "DAR", "0.00%")
    a(12) = Array("运营能力指标", "存货周转天数", "Days Inventory Outstanding (DIO)", "DIO", "0.00")
    a(13) = Array("运营能力指标", "应收款周转天数", "Days Sales Outstanding (DSO)", "DSO", "0.00")
    a(14) = Array("运营能力指标", "应付账款周转天数", "Days Payable Outstanding (DPO)", "DPO", "0.00")
    a(15) = Array("运营能力指标", "营运资金周转天数", "Cash Conversion Cycle (CCC)", "CCC", "0.00")
    a(16) = Array("运营能力指标", "流动资产周转率", "Current Asset Turnover", "CAT", "0.00")
    a(17) = Array("运营能力指标", "总资产周转率", "Total Asset Turnover", "TAT", "0.00")
    StandardIndicatorDefs = a
End Function


Private Function StandardRowMap(ByVal marketKey As String) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    If marketKey = "US" Then
        AddStandardRow m, "REV", "美股_利润表", Array("Revenue")
        AddStandardRow m, "COGS", "美股_利润表", Array("Cost of goods & services sold")
        AddStandardRow m, "GP", "美股_利润表", Array("Gross profit")
        AddStandardRow m, "RD", "美股_利润表", Array("R&D expense")
        AddStandardRow m, "SGA", "美股_利润表", Array("SG&A expense")
        AddStandardRow m, "OPEX", "美股_利润表", Array("Total operating expenses")
        AddStandardRow m, "NI", "美股_利润表", Array("Net income")
        AddStandardRow m, "PNI", "美股_利润表", Array("Net income")
        AddStandardRow m, "TA", "美股_资产负债表", Array("Total assets")
        AddStandardRow m, "TL", "美股_资产负债表", Array("Total liabilities")
        AddStandardRow m, "EQ", "美股_资产负债表", Array("Total stockholders' equity")
        AddStandardRow m, "CA", "美股_资产负债表", Array("Total current assets")
        AddStandardRow m, "CL", "美股_资产负债表", Array("Total current liabilities")
        AddStandardRow m, "INV", "美股_资产负债表", Array("Inventory")
        AddStandardRow m, "AR", "美股_资产负债表", Array("Accounts receivable, net")
        AddStandardRow m, "AP", "美股_资产负债表", Array("Accounts payable")
        AddStandardRow m, "CASH", "美股_资产负债表", Array("Cash & equivalents")
    ElseIf marketKey = "HK" Then
        AddStandardRow m, "REV", "港股_利润表", Array("Revenue")
        AddStandardRow m, "COGS", "港股_利润表", Array("Cost of goods & services sold")
        AddStandardRow m, "GP", "港股_利润表", Array("Gross profit")
        AddStandardRow m, "RD", "港股_利润表", Array("R&D expense")
        AddStandardRow m, "SELL", "港股_利润表", Array("Selling expense")
        AddStandardRow m, "ADMIN", "港股_利润表", Array("Administrative expense")
        AddStandardRow m, "OPEX", "港股_利润表", Array("Total operating expenses")
        AddStandardRow m, "NI", "港股_利润表", Array("Net income")
        AddStandardRow m, "PNI", "港股_利润表", Array("Net income")
        AddStandardRow m, "TA", "港股_资产负债表", Array("Total assets")
        AddStandardRow m, "TL", "港股_资产负债表", Array("Total liabilities")
        AddStandardRow m, "EQ", "港股_资产负债表", Array("Total stockholders' equity", "Total equity")
        AddStandardRow m, "CA", "港股_资产负债表", Array("Total current assets")
        AddStandardRow m, "CL", "港股_资产负债表", Array("Total current liabilities")
        AddStandardRow m, "INV", "港股_资产负债表", Array("Inventory")
        AddStandardRow m, "AR", "港股_资产负债表", Array("Accounts receivable, net")
        AddStandardRow m, "AP", "港股_资产负债表", Array("Accounts payable")
        AddStandardRow m, "CASH", "港股_资产负债表", Array("Cash & equivalents")
    ElseIf marketKey = "KR" Then
        AddStandardRow m, "REV", "韩股_利润表", Array("Revenue")
        AddStandardRow m, "COGS", "韩股_利润表", Array("Cost of goods & services sold")
        AddStandardRow m, "GP", "韩股_利润表", Array("Gross profit")
        AddStandardRow m, "RD", "韩股_利润表", Array("R&D expense")
        AddStandardRow m, "SGA", "韩股_利润表", Array("SG&A expense")
        AddStandardRow m, "OPEX", "韩股_利润表", Array("Total operating expenses")
        AddStandardRow m, "NI", "韩股_利润表", Array("Net income")
        AddStandardRow m, "PNI", "韩股_利润表", Array("Net income")
        AddStandardRow m, "TA", "韩股_资产负债表", Array("Total assets")
        AddStandardRow m, "TL", "韩股_资产负债表", Array("Total liabilities")
        AddStandardRow m, "EQ", "韩股_资产负债表", Array("Total stockholders' equity", "Total equity")
        AddStandardRow m, "CA", "韩股_资产负债表", Array("Total current assets")
        AddStandardRow m, "CL", "韩股_资产负债表", Array("Total current liabilities")
        AddStandardRow m, "INV", "韩股_资产负债表", Array("Inventory")
        AddStandardRow m, "AR", "韩股_资产负债表", Array("Accounts receivable, net")
        AddStandardRow m, "AP", "韩股_资产负债表", Array("Accounts payable")
        AddStandardRow m, "CASH", "韩股_资产负债表", Array("Cash & equivalents")
    Else
        AddStandardRow m, "REV", "A股_利润表", Array("营业收入", "一、营业总收入")
        AddStandardRow m, "COGS", "A股_利润表", Array("营业成本")
        AddStandardRow m, "TAXSUR", "A股_利润表", Array("营业税金及附加")
        AddStandardRow m, "SELL", "A股_利润表", Array("销售费用")
        AddStandardRow m, "ADMIN", "A股_利润表", Array("管理费用")
        AddStandardRow m, "FIN", "A股_利润表", Array("财务费用")
        AddStandardRow m, "RD", "A股_利润表", Array("研发费用")
        AddStandardRow m, "NI", "A股_利润表", Array("五、净利润")
        AddStandardRow m, "PNI", "A股_利润表", Array("归属于母公司所有者的净利润")
        AddStandardRow m, "TA", "A股_资产负债表", Array("资产总计")
        AddStandardRow m, "TL", "A股_资产负债表", Array("负债合计")
        AddStandardRow m, "EQ", "A股_资产负债表", Array("归属于母公司股东权益合计", "所有者权益(或股东权益)合计")
        AddStandardRow m, "CA", "A股_资产负债表", Array("流动资产合计")
        AddStandardRow m, "CL", "A股_资产负债表", Array("流动负债合计")
        AddStandardRow m, "INV", "A股_资产负债表", Array("存货")
        AddStandardRow m, "AR", "A股_资产负债表", Array("应收账款", "应收票据及应收账款")
        AddStandardRow m, "AP", "A股_资产负债表", Array("应付账款", "应付票据及应付账款")
        AddStandardRow m, "CASH", "A股_资产负债表", Array("货币资金")
    End If
    Set StandardRowMap = m
End Function


Private Sub AddStandardRow(ByVal rowMap As Object, ByVal key As String, _
                           ByVal sheetName As String, ByVal candidates As Variant)
    Dim rowNum As Long
    rowNum = FindStandardIndicatorRow(sheetName, candidates)
    If rowNum > 0 Then rowMap.Item(key) = Array(sheetName, rowNum)
End Sub


Private Function StandardIndicatorFormula(ByVal metricKey As String, ByVal marketKey As String, _
                                          ByVal rowMap As Object, ByVal companyHeader As String, _
                                          ByVal periodText As String, ByVal currentSheet As String, _
                                          ByVal rowNum As Long, ByVal colNum As Long) As String
    Dim bsSheet As String, isSheet As String
    If marketKey = "US" Then
        bsSheet = "美股_资产负债表"
        isSheet = "美股_利润表"
    ElseIf marketKey = "HK" Then
        bsSheet = "港股_资产负债表"
        isSheet = "港股_利润表"
    ElseIf marketKey = "KR" Then
        bsSheet = "韩股_资产负债表"
        isSheet = "韩股_利润表"
    Else
        bsSheet = "A股_资产负债表"
        isSheet = "A股_利润表"
    End If

    Dim bsCol As Long: bsCol = FindStandardStatementColumn(bsSheet, companyHeader, periodText)
    Dim isCol As Long: isCol = FindStandardStatementColumn(isSheet, companyHeader, periodText)
    Dim bsBaseCol As Long: bsBaseCol = FindPriorFiscalYearEndStatementColumn(bsSheet, companyHeader, periodText, marketKey)
    Dim isSamePeriodCol As Long: isSamePeriodCol = FindPriorSamePeriodStatementColumn(isSheet, companyHeader, periodText)

    Dim rev As String: rev = StandardRef(rowMap, "REV", isCol)
    Dim revP As String: revP = StandardRef(rowMap, "REV", isSamePeriodCol)
    Dim cogs As String: cogs = StandardRef(rowMap, "COGS", isCol)
    Dim taxSur As String: taxSur = StandardRef(rowMap, "TAXSUR", isCol)
    Dim ni As String: ni = StandardRef(rowMap, "NI", isCol)
    Dim pni As String: pni = StandardRef(rowMap, "PNI", isCol)
    Dim niP As String: niP = StandardRef(rowMap, "NI", isSamePeriodCol)
    Dim ta As String: ta = StandardRef(rowMap, "TA", bsCol)
    Dim taP As String: taP = StandardRef(rowMap, "TA", bsBaseCol)
    Dim tl As String: tl = StandardRef(rowMap, "TL", bsCol)
    Dim eq As String: eq = StandardRef(rowMap, "EQ", bsCol)
    Dim eqP As String: eqP = StandardRef(rowMap, "EQ", bsBaseCol)
    Dim ca As String: ca = StandardRef(rowMap, "CA", bsCol)
    Dim caP As String: caP = StandardRef(rowMap, "CA", bsBaseCol)
    Dim cl As String: cl = StandardRef(rowMap, "CL", bsCol)
    Dim inv As String: inv = StandardRef(rowMap, "INV", bsCol)
    Dim invP As String: invP = StandardRef(rowMap, "INV", bsBaseCol)
    Dim ar As String: ar = StandardRef(rowMap, "AR", bsCol)
    Dim arP As String: arP = StandardRef(rowMap, "AR", bsBaseCol)
    Dim ap As String: ap = StandardRef(rowMap, "AP", bsCol)
    Dim apP As String: apP = StandardRef(rowMap, "AP", bsBaseCol)
    Dim cashRef As String: cashRef = StandardRef(rowMap, "CASH", bsCol)
    Dim daysText As String: daysText = CStr(StandardDaysForPeriod(periodText, marketKey))

    Select Case metricKey
        Case "NPM"
            StandardIndicatorFormula = RatioFormula(ni, rev)
        Case "GPM"
            Dim gp As String: gp = StandardRef(rowMap, "GP", isCol)
            If Len(gp) > 0 Then
                StandardIndicatorFormula = RatioFormula(gp, rev)
            ElseIf Len(rev) > 0 And Len(cogs) > 0 Then
                If Len(taxSur) = 0 Then taxSur = "0"
                StandardIndicatorFormula = RatioFormula(rev & "-" & cogs & "-" & taxSur, rev)
            End If
        Case "OER"
            Dim opex As String
            If marketKey = "US" Or marketKey = "KR" Then
                opex = StandardRef(rowMap, "OPEX", isCol)
                If Len(opex) = 0 Then _
                    opex = SumExpression(Array(StandardRef(rowMap, "RD", isCol), StandardRef(rowMap, "SGA", isCol)))
            ElseIf marketKey = "HK" Then
                opex = StandardRef(rowMap, "OPEX", isCol)
                If Len(opex) = 0 Then _
                    opex = SumExpression(Array(StandardRef(rowMap, "SELL", isCol), _
                                               StandardRef(rowMap, "ADMIN", isCol), _
                                               StandardRef(rowMap, "RD", isCol)))
            Else
                opex = SumExpression(Array(StandardRef(rowMap, "SELL", isCol), StandardRef(rowMap, "ADMIN", isCol), _
                                           StandardRef(rowMap, "FIN", isCol)))
            End If
            StandardIndicatorFormula = RatioFormula(opex, rev)
        Case "ROA"
            StandardIndicatorFormula = RatioFormula(ni, AverageExpression(ta, taP))
        Case "ROE"
            If Len(pni) = 0 Then pni = ni
            If marketKey = "A" Then
                StandardIndicatorFormula = RatioFormula(pni, eq)
            Else
                StandardIndicatorFormula = RatioFormula(ni, AverageExpression(eq, eqP))
            End If
        Case "TAGR"
            StandardIndicatorFormula = GrowthFormula(ta, taP)
        Case "RGR"
            StandardIndicatorFormula = GrowthFormula(rev, revP)
        Case "NPGR"
            StandardIndicatorFormula = GrowthFormula(ni, niP)
        Case "CR"
            StandardIndicatorFormula = RatioFormula(ca, cl)
        Case "QR"
            If Len(ca) > 0 And Len(cl) > 0 Then
                If Len(inv) = 0 Then inv = "0"
                StandardIndicatorFormula = RatioFormula(ca & "-" & inv, cl)
            End If
        Case "CASHR"
            StandardIndicatorFormula = RatioFormula(cashRef, cl)
        Case "DAR"
            StandardIndicatorFormula = RatioFormula(tl, ta)
        Case "DIO"
            Dim avgInv As String: avgInv = AverageExpression(inv, invP)
            If Len(avgInv) > 0 Then StandardIndicatorFormula = RatioFormula(avgInv & "*" & daysText, cogs)
        Case "DSO"
            Dim avgAr As String: avgAr = AverageExpression(ar, arP)
            If Len(avgAr) > 0 Then StandardIndicatorFormula = RatioFormula(avgAr & "*" & daysText, rev)
        Case "DPO"
            Dim avgAp As String: avgAp = AverageExpression(ap, apP)
            If Len(avgAp) > 0 Then StandardIndicatorFormula = RatioFormula(avgAp & "*" & daysText, cogs)
        Case "CCC"
            StandardIndicatorFormula = CccFormula(rowNum, colNum)
        Case "CAT"
            StandardIndicatorFormula = RatioFormula(rev, AverageExpression(ca, caP))
        Case "TAT"
            StandardIndicatorFormula = RatioFormula(rev, AverageExpression(ta, taP))
    End Select
End Function


Private Function RatioFormula(ByVal numerator As String, ByVal denominator As String) As String
    If Len(numerator) = 0 Or Len(denominator) = 0 Then Exit Function
    RatioFormula = "=IFERROR((" & numerator & ")/(" & denominator & "),"""")"
End Function


Private Function GrowthFormula(ByVal currentRef As String, ByVal priorRef As String) As String
    If Len(currentRef) = 0 Or Len(priorRef) = 0 Then Exit Function
    GrowthFormula = "=IFERROR((" & currentRef & ")/(" & priorRef & ")-1,"""")"
End Function


Private Function AverageExpression(ByVal currentRef As String, ByVal priorRef As String) As String
    If Len(currentRef) = 0 Then Exit Function
    If Len(priorRef) > 0 Then
        AverageExpression = "AVERAGE(" & currentRef & "," & priorRef & ")"
    Else
        AverageExpression = currentRef
    End If
End Function


Private Function SumExpression(ByVal refs As Variant) As String
    Dim i As Long, part As String
    For i = LBound(refs) To UBound(refs)
        part = CStr(refs(i))
        If Len(part) > 0 Then
            If Len(SumExpression) > 0 Then
                SumExpression = SumExpression & "+" & part
            Else
                SumExpression = part
            End If
        End If
    Next i
End Function


Private Function CccFormula(ByVal rowNum As Long, ByVal colNum As Long) As String
    Dim colLetter As String: colLetter = StandardColumnLetter(colNum)
    Dim dioRef As String: dioRef = colLetter & (rowNum - 3)
    Dim dsoRef As String: dsoRef = colLetter & (rowNum - 2)
    Dim dpoRef As String: dpoRef = colLetter & (rowNum - 1)
    CccFormula = "=IFERROR(IF(OR(" & dioRef & "=""""," & dsoRef & "=""""," & _
                 dpoRef & "=""""),""""," & dioRef & "+" & dsoRef & "-" & dpoRef & "),"""")"
End Function


Private Function StandardRef(ByVal rowMap As Object, ByVal key As String, ByVal colNum As Long) As String
    If colNum <= 0 Then Exit Function
    If Not rowMap.Exists(key) Then Exit Function
    Dim meta As Variant: meta = rowMap.Item(key)
    Dim sheetName As String: sheetName = CStr(meta(0))
    Dim rowNum As Long: rowNum = CLng(meta(1))
    If rowNum <= 0 Then Exit Function
    StandardRef = "'" & sheetName & "'!" & StandardColumnLetter(colNum) & rowNum
End Function


Private Function FindStandardIndicatorRow(ByVal sheetName As String, ByVal candidates As Variant) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then GoTo CleanExit

    Dim i As Long, found As Range
    For i = LBound(candidates) To UBound(candidates)
        Set found = ws.Range("B:B").Find(What:=CStr(candidates(i)), LookIn:=xlValues, _
                                         LookAt:=xlWhole, MatchCase:=True)
        If Not found Is Nothing Then
            FindStandardIndicatorRow = found.Row
            GoTo CleanExit
        End If
    Next i

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Function FindStandardStatementColumn(ByVal sheetName As String, _
                                             ByVal companyHeader As String, _
                                             ByVal periodText As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then GoTo CleanExit
    If Len(companyHeader) = 0 Or Len(periodText) = 0 Then GoTo CleanExit

    Dim dataStartCol As Long: dataStartCol = StandardDataStartCol(ws)
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = dataStartCol To lastCol
        If StandardHeaderTextAt(ws, c) = companyHeader _
           And StandardPeriodKey(ws.Cells(2, c).Value) = periodText Then
            FindStandardStatementColumn = c
            GoTo CleanExit
        End If
    Next c

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Function FindPriorStandardStatementColumn(ByVal sheetName As String, _
                                                  ByVal companyHeader As String, _
                                                  ByVal periodText As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then GoTo CleanExit
    If Len(companyHeader) = 0 Or Len(periodText) = 0 Then GoTo CleanExit

    Dim currentKey As String: currentKey = StandardPeriodKey(periodText)
    Dim targetKey As String, targetMd As String
    If IsDate(currentKey) Then
        targetKey = Format$(DateAdd("yyyy", -1, CDate(currentKey)), "yyyy-mm-dd")
        targetMd = Mid$(currentKey, 6, 5)
    End If

    Dim dataStartCol As Long: dataStartCol = StandardDataStartCol(ws)
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, pKey As String
    For c = dataStartCol To lastCol
        pKey = StandardPeriodKey(ws.Cells(2, c).Value)
        If StandardHeaderTextAt(ws, c) = companyHeader And pKey = targetKey Then
            FindPriorStandardStatementColumn = c
            GoTo CleanExit
        End If
    Next c

    Dim bestKey As String, bestCol As Long
    If Len(targetMd) > 0 Then
        For c = dataStartCol To lastCol
            pKey = StandardPeriodKey(ws.Cells(2, c).Value)
            If StandardHeaderTextAt(ws, c) = companyHeader _
               And pKey < currentKey And Mid$(pKey, 6, 5) = targetMd Then
                If pKey > bestKey Then
                    bestKey = pKey
                    bestCol = c
                End If
            End If
        Next c
    End If
    If bestCol > 0 Then
        FindPriorStandardStatementColumn = bestCol
        GoTo CleanExit
    End If

    For c = dataStartCol To lastCol
        pKey = StandardPeriodKey(ws.Cells(2, c).Value)
        If StandardHeaderTextAt(ws, c) = companyHeader And pKey < currentKey Then
            If pKey > bestKey Then
                bestKey = pKey
                bestCol = c
            End If
        End If
    Next c
    FindPriorStandardStatementColumn = bestCol

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Function FindPriorSamePeriodStatementColumn(ByVal sheetName As String, _
                                                    ByVal companyHeader As String, _
                                                    ByVal periodText As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then GoTo CleanExit
    If Len(companyHeader) = 0 Or Len(periodText) = 0 Then GoTo CleanExit

    Dim currentKey As String: currentKey = StandardPeriodKey(periodText)
    If Not IsDate(currentKey) Then GoTo CleanExit
    Dim targetKey As String: targetKey = Format$(DateAdd("yyyy", -1, CDate(currentKey)), "yyyy-mm-dd")

    Dim dataStartCol As Long: dataStartCol = StandardDataStartCol(ws)
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = dataStartCol To lastCol
        If StandardHeaderTextAt(ws, c) = companyHeader _
           And StandardPeriodKey(ws.Cells(2, c).Value) = targetKey Then
            FindPriorSamePeriodStatementColumn = c
            GoTo CleanExit
        End If
    Next c

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Function FindPriorFiscalYearEndStatementColumn(ByVal sheetName As String, _
                                                       ByVal companyHeader As String, _
                                                       ByVal periodText As String, _
                                                       ByVal marketKey As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then GoTo CleanExit
    If Len(companyHeader) = 0 Or Len(periodText) = 0 Then GoTo CleanExit

    Dim currentKey As String: currentKey = StandardPeriodKey(periodText)
    If Not IsDate(currentKey) Then GoTo CleanExit

    Dim dataStartCol As Long: dataStartCol = StandardDataStartCol(ws)
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, pKey As String, targetKey As String

    If marketKey = "A" Then
        targetKey = CStr(Year(CDate(currentKey)) - 1) & "-12-31"
        For c = dataStartCol To lastCol
            If StandardHeaderTextAt(ws, c) = companyHeader _
               And StandardPeriodKey(ws.Cells(2, c).Value) = targetKey Then
                FindPriorFiscalYearEndStatementColumn = c
                GoTo CleanExit
            End If
        Next c
        GoTo CleanExit
    End If

    ' 美股 fiscal year-end 未必是 12/31: 取上一自然年里该公司最后一个报告期,
    ' 通常就是上一财年的年报期末 (例如 2024-12-31 或 AAPL 的 2024-09-28)。
    Dim bestKey As String, bestCol As Long
    For c = dataStartCol To lastCol
        pKey = StandardPeriodKey(ws.Cells(2, c).Value)
        If StandardHeaderTextAt(ws, c) = companyHeader _
           And pKey < currentKey _
           And IsDate(pKey) _
           And Year(CDate(pKey)) <= Year(CDate(currentKey)) - 1 Then
            If pKey > bestKey Then
                bestKey = pKey
                bestCol = c
            End If
        End If
    Next c
    If bestCol > 0 Then
        FindPriorFiscalYearEndStatementColumn = bestCol
        GoTo CleanExit
    End If

    For c = dataStartCol To lastCol
        pKey = StandardPeriodKey(ws.Cells(2, c).Value)
        If StandardHeaderTextAt(ws, c) = companyHeader And pKey < currentKey Then
            If pKey > bestKey Then
                bestKey = pKey
                bestCol = c
            End If
        End If
    Next c
    FindPriorFiscalYearEndStatementColumn = bestCol

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Function StandardTargetPeriodWanted(ByVal periodText As String, _
                                            ByVal selectedYear As Long, _
                                            ByVal quarterPick As String, _
                                            Optional ByVal marketKey As String = "A", _
                                            Optional ByVal wsSource As Worksheet = Nothing, _
                                            Optional ByVal companyHeader As String = "", _
                                            Optional ByVal sourceCol As Long = 0) As Boolean
    StandardTargetPeriodWanted = False
    If Len(periodText) < 10 Then Exit Function

    If UCase$(marketKey) = "US" And quarterPick <> "全部" Then
        ' 美股 fiscal quarter 不一定落在自然季末, 例如 AAPL FY2024 Q4 = 2024-09-28。
        ' 前序美股抓数已经按 EDGAR fp(Q1/Q2/Q3/FY) 过滤; 指标表这里不再用日期后缀判断。
        If selectedYear = 0 Then
            StandardTargetPeriodWanted = True
        ElseIf Not wsSource Is Nothing Then
            StandardTargetPeriodWanted = IsLatestSourceColumnForCompany(wsSource, companyHeader, sourceCol)
        End If
        Exit Function
    End If

    If UCase$(marketKey) = "HK" And quarterPick <> "全部" Then
        ' 港股抓数已按 month_num + ed 过滤; Q4 年报可能是 03-31/06-30/12-31,
        ' 指标表这里只按年份保留,避免误删阿里 H 这类 3 月财年年报。
        If selectedYear = 0 Then
            StandardTargetPeriodWanted = True
        Else
            StandardTargetPeriodWanted = (CLng(Left$(periodText, 4)) = selectedYear)
        End If
        Exit Function
    End If

    If selectedYear > 0 Then
        If CLng(Left$(periodText, 4)) <> selectedYear Then Exit Function
    End If

    Dim suffix As String
    Select Case quarterPick
        Case "Q1": suffix = "-03-31"
        Case "Q2": suffix = "-06-30"
        Case "Q3": suffix = "-09-30"
        Case "Q4": suffix = "-12-31"
        Case Else
            StandardTargetPeriodWanted = True
            Exit Function
    End Select
    StandardTargetPeriodWanted = (Right$(periodText, Len(suffix)) = suffix)
End Function


Private Function IsLatestSourceColumnForCompany(ByVal ws As Worksheet, _
                                                ByVal companyHeader As String, _
                                                ByVal sourceCol As Long) As Boolean
    On Error Resume Next
    IsLatestSourceColumnForCompany = False
    If ws Is Nothing Or Len(companyHeader) = 0 Or sourceCol <= 0 Then GoTo CleanExit

    Dim currentKey As String: currentKey = StandardPeriodKey(ws.Cells(2, sourceCol).Value)
    If Len(currentKey) = 0 Then GoTo CleanExit

    Dim dataStartCol As Long: dataStartCol = StandardDataStartCol(ws)
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, pKey As String
    For c = dataStartCol To lastCol
        If c <> sourceCol And StandardHeaderTextAt(ws, c) = companyHeader Then
            pKey = StandardPeriodKey(ws.Cells(2, c).Value)
            If pKey > currentKey Then GoTo CleanExit
        End If
    Next c

    IsLatestSourceColumnForCompany = True

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Sub MergeStandardCompanyHeaders(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long)
    If lastCol < firstCol Then Exit Sub

    Dim startCol As Long: startCol = firstCol
    Dim c As Long, currentHeader As String, nextHeader As String
    currentHeader = Trim$(CStr(ws.Cells(1, firstCol).Value))

    For c = firstCol + 1 To lastCol + 1
        If c <= lastCol Then
            nextHeader = Trim$(CStr(ws.Cells(1, c).Value))
        Else
            nextHeader = vbNullString
        End If

        If nextHeader <> currentHeader Then
            If c - startCol > 1 Then
                ws.Range(ws.Cells(1, startCol), ws.Cells(1, c - 1)).Merge
                ws.Cells(1, startCol).Value = currentHeader
            End If
            startCol = c
            currentHeader = nextHeader
        End If
    Next c
End Sub


Private Function StandardDataStartCol(ByVal ws As Worksheet) As Long
    If ws.Name = "A股_指标表" Or ws.Name = "美股_指标表" Or _
       ws.Name = "港股_指标表" Or ws.Name = "韩股_指标表" Then
        StandardDataStartCol = 4
    Else
        StandardDataStartCol = 3
    End If
End Function


Private Function StandardHeaderTextAt(ByVal ws As Worksheet, ByVal col As Long) As String
    On Error Resume Next
    If ws.Cells(1, col).MergeCells Then
        StandardHeaderTextAt = Trim$(CStr(ws.Cells(1, col).MergeArea.Cells(1, 1).Value))
    Else
        StandardHeaderTextAt = Trim$(CStr(ws.Cells(1, col).Value))
    End If
    Err.Clear
    On Error GoTo 0
End Function


Private Function StandardPeriodKey(ByVal v As Variant) As String
    On Error Resume Next
    If IsDate(v) Then
        StandardPeriodKey = Format$(CDate(v), "yyyy-mm-dd")
    Else
        StandardPeriodKey = Trim$(CStr(v))
        If Len(StandardPeriodKey) >= 10 And Mid$(StandardPeriodKey, 5, 1) = "-" Then _
            StandardPeriodKey = Left$(StandardPeriodKey, 10)
    End If
    Err.Clear
    On Error GoTo 0
End Function


Private Function StandardDaysForPeriod(ByVal periodText As String, ByVal marketKey As String) As Long
    If marketKey = "A" Then
        Select Case Mid$(StandardPeriodKey(periodText), 6, 5)
            Case "03-31": StandardDaysForPeriod = 90
            Case "06-30": StandardDaysForPeriod = 180
            Case "09-30": StandardDaysForPeriod = 270
            Case Else:    StandardDaysForPeriod = 360
        End Select
    Else
        Select Case Mid$(StandardPeriodKey(periodText), 6, 5)
            Case "03-31": StandardDaysForPeriod = 90
            Case "06-30": StandardDaysForPeriod = 180
            Case "09-30": StandardDaysForPeriod = 270
            Case Else:    StandardDaysForPeriod = 365
        End Select
    End If
End Function


Private Function StandardColumnLetter(ByVal col As Long) As String
    StandardColumnLetter = Split(ThisWorkbook.Sheets(1).Cells(1, col).Address, "$")(1)
End Function


' --------- 通用单张财务报表抓数主流程 (4 个抓数模块共用) ---------
'   strID         : HTML <table id="..."> 值
'   strType       : "balance" / "profit" / "cash" / "indicator" — 用于内部拼新浪 URL
'   targetSheet   : 目标宽表 Sheet 名
'   blnSilent     : True 时不弹 MsgBox, 用于一键全抓静默调用
'   includePriorYear: 年份选择时额外抓上一年, 供标准指标同比/上一年末公式引用
Public Sub RunOneStatement(ByVal strID As String, ByVal strType As String, _
                            ByVal targetSheet As String, _
                            Optional ByVal blnSilent As Boolean = False, _
                            Optional ByVal includePriorYear As Boolean = False)
    Dim wsPool As Worksheet, wsTarget As Worksheet
    Dim dictData As Object, dictPeriodSet As Object
    Dim dictIndicatorSet As Object, dictCategoryMap As Object
    Dim dictCompanyName As Object
    Dim objHtml As Object
    Dim collCodes As Collection
    Dim arrPool As Variant
    Dim i As Long, lngRow As Long, numCompanies As Long
    Dim intFailCnt As Long, strErrLog As String
    Dim strCode As String, strName As String, strUrl As String
    Dim strHtml As String, strTbl As String
    Dim lngYear As Long
    Dim dtTime As Double
    dtTime = Timer

    Set wsPool = ThisWorkbook.Sheets("样本池")
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)
    Set collCodes = New Collection
    Set dictData = CreateObject("Scripting.Dictionary")
    Set dictPeriodSet = CreateObject("Scripting.Dictionary")
    Set dictIndicatorSet = CreateObject("Scripting.Dictionary")
    Set dictCategoryMap = CreateObject("Scripting.Dictionary")
    Set dictCompanyName = CreateObject("Scripting.Dictionary")
    Set objHtml = CreateObject("htmlfile")

    On Error GoTo CleanUp
    Application.ScreenUpdating = False

    ' A2 年份: 0=取最新季度, >0=该年报告 (跟以前 HYPERLINK 公式逻辑等价)
    lngYear = ReadYearSelection()
    Dim collFetchYears As Collection: Set collFetchYears = New Collection
    collFetchYears.Add lngYear
    If includePriorYear And lngYear > 0 Then collFetchYears.Add lngYear - 1

    ' ---- 读样本池 A股区 (A=代码 / B=简称) ----
    lngRow = wsPool.Cells(wsPool.Rows.Count, POOL_A_CODE_COL).End(xlUp).Row
    If lngRow < POOL_DATA_START_ROW Then
        intFailCnt = 1
        strErrLog = "样本池 A股区为空, 请在第 " & POOL_DATA_START_ROW & " 行起录入股票代码和简称"
        GoTo CleanUp
    End If
    arrPool = wsPool.Range(wsPool.Cells(POOL_DATA_START_ROW, POOL_A_CODE_COL), _
                           wsPool.Cells(lngRow, POOL_A_NAME_COL)).Value
    If Not IsArray(arrPool) Then
        Dim singleVal As Variant: singleVal = arrPool
        ReDim arrPool(1 To 1, 1 To 2)
        arrPool(1, 1) = singleVal
    End If
    numCompanies = UBound(arrPool, 1)

    ' ---- 第一遍: 逐家公司抓 + 解析 ----
    For i = 1 To numCompanies
        strCode = Trim$(CStr(arrPool(i, 1)))
        If Len(strCode) = 0 Then GoTo NextRow
        strName = Trim$(CStr(arrPool(i, 2)))

        Application.StatusBar = "抓取中: " & targetSheet & " (" & i & "/" & numCompanies & ") " & strName
        DoEvents

        Dim fetchYear As Variant
        Dim mainYearOk As Boolean: mainYearOk = False
        Dim anyYearOk As Boolean: anyYearOk = False
        Dim mainErrDesc As String: mainErrDesc = ""

        For Each fetchYear In collFetchYears
            ' 内部按代码 + 类型 + 年份拼 URL (取代之前的 HYPERLINK 公式列)
            strUrl = BuildSinaFinancialUrl(strCode, strType, CLng(fetchYear))

            On Error Resume Next
            Err.Clear
            strHtml = HttpGet(strUrl)
            If Err.Number = 0 Then strTbl = ExtractTable(strHtml, strID)
            If Err.Number = 0 Then _
                ParseFinancialHtml strTbl, strID, strCode, objHtml, _
                                    dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap
            If Err.Number <> 0 Then
                If CLng(fetchYear) = lngYear Then mainErrDesc = Err.Description
                Err.Clear
            Else
                anyYearOk = True
                If CLng(fetchYear) = lngYear Then mainYearOk = True
            End If
            On Error GoTo CleanUp

            ' 限速 1s, 防触发反爬
            Application.Wait Now + TimeSerial(0, 0, 1)
        Next fetchYear

        If (lngYear = 0 And Not anyYearOk) Or (lngYear > 0 And Not mainYearOk) Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": " & mainErrDesc
        Else
            If Not dictCompanyName.Exists(strCode) Then dictCompanyName.Add strCode, strName
            collCodes.Add strCode
        End If

NextRow:
    Next i

    ' ---- 第二遍: 排序并写宽表 ----
    If collCodes.Count = 0 Then GoTo CleanUp

    Dim arrCodes() As String
    ReDim arrCodes(1 To collCodes.Count)
    For i = 1 To collCodes.Count
        arrCodes(i) = collCodes(i)
    Next i

    ' Phase 3: 季度过滤 (读样本池 A4, 留下匹配后缀的报告期)
    FilterPeriodsByQuarter dictPeriodSet, ReadQuarterSelection()

    Dim arrPeriods As Variant, arrIndicators As Variant
    arrPeriods = SortPeriodsDesc(dictPeriodSet)
    arrIndicators = IndicatorsByInsertion(dictIndicatorSet)

    Application.StatusBar = "写入: " & targetSheet
    DoEvents
    WriteWideTable wsTarget, arrCodes, dictCompanyName, dictData, _
                    arrPeriods, arrIndicators, dictCategoryMap, _
                    perCompanyPeriods:=False, _
                    dictReportingCurrency:=Nothing, _
                    statementKind:=strKind
    RefreshA1CurrencyComment wsTarget, targetSheet

CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' 累计到全局 (一键全抓 用)
    g_globalFails = g_globalFails + intFailCnt
    If Len(strErrLog) > 0 Then _
        g_globalLog = g_globalLog & vbCrLf & "[" & targetSheet & "]" & strErrLog

    If Not blnSilent And Not g_silentMode Then
        Dim msg As String
        msg = targetSheet & " 抓取完成" & vbCrLf & _
              "用时: " & Format(Timer - dtTime, "0.0 秒") & vbCrLf & _
              "公司数: " & collCodes.Count & " / 期数: " & dictPeriodSet.Count & _
              " / 指标数: " & dictIndicatorSet.Count
        If intFailCnt > 0 Then msg = msg & vbCrLf & vbCrLf & "失败 " & intFailCnt & " 条:" & strErrLog

        Dim style As Long: style = vbInformation
        If intFailCnt > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询"
    End If
End Sub


' --------- 公司基本资料抓数主流程 (写到 上市公司基本资料 Sheet, 平表) ---------
'   targetSheet = "上市公司基本资料"
'   urlCol      = 7 (G 列)
'   blnSilent   = True 时不弹 MsgBox
Public Sub RunCorpInfoFetch(Optional ByVal blnSilent As Boolean = False)
    Const TARGET_SHEET As String = "上市公司基本资料"
    Const URL_COL As Long = 7
    Dim wsPool As Worksheet, wsTarget As Worksheet
    Dim arrPool As Variant
    Dim objHtml As Object
    Dim i As Long, lngRow As Long, numCompanies As Long
    Dim intFailCnt As Long, strErrLog As String, intSuccessCnt As Long
    Dim strCode As String, strName As String, strUrl As String
    Dim strHtml As String, strTbl As String
    Dim dictInfo As Object
    Dim dtTime As Double
    dtTime = Timer

    Set wsPool = ThisWorkbook.Sheets("样本池")
    Set wsTarget = ThisWorkbook.Sheets(TARGET_SHEET)
    Set objHtml = CreateObject("htmlfile")

    On Error GoTo CleanUp
    Application.ScreenUpdating = False

    lngRow = wsPool.Range("A" & wsPool.Rows.Count).End(xlUp).Row
    If lngRow < POOL_DATA_START_ROW Then
        intFailCnt = 1
        strErrLog = "样本池为空"
        GoTo CleanUp
    End If
    arrPool = wsPool.Range("A" & POOL_DATA_START_ROW & ":H" & lngRow).Value
    If Not IsArray(arrPool) Then
        Dim singleVal As Variant: singleVal = arrPool
        ReDim arrPool(1 To 1, 1 To POOL_LAST_COL)
        arrPool(1, 1) = singleVal
    End If
    numCompanies = UBound(arrPool, 1)

    ' 清旧数据 (保留 Row 1 表头)
    Dim lastUsedRow As Long
    lastUsedRow = wsTarget.UsedRange.Rows.Count + wsTarget.UsedRange.Row - 1
    If lastUsedRow > 1 Then _
        wsTarget.Range("A2:E" & lastUsedRow).Clear

    ' 一次性写入数组
    Dim arrOut() As Variant
    ReDim arrOut(1 To numCompanies, 1 To 5)
    Dim outRow As Long: outRow = 0

    For i = 1 To numCompanies
        strCode = Trim$(CStr(arrPool(i, 1)))
        If Len(strCode) = 0 Then GoTo NextRow
        strName = Trim$(CStr(arrPool(i, 2)))
        ' Phase 4: 跳过非 A 股
        If ResolveMarket(strCode, CStr(arrPool(i, POOL_MARKET_COL))) <> "A" Then GoTo NextRow
        strUrl = Trim$(CStr(arrPool(i, URL_COL)))
        If Len(strUrl) = 0 Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": URL 为空"
            GoTo NextRow
        End If

        Application.StatusBar = "抓取中: 基本资料 (" & i & "/" & numCompanies & ") " & strName
        DoEvents

        On Error Resume Next
        Err.Clear
        strHtml = HttpGet(strUrl)
        If Err.Number = 0 Then strTbl = ExtractTable(strHtml, "comInfo1")
        If Err.Number = 0 Then Set dictInfo = ParseCorpInfoHtml(strTbl, objHtml)
        If Err.Number <> 0 Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": " & Err.Description
            Err.Clear
            On Error GoTo CleanUp
            GoTo NextRow
        End If
        On Error GoTo CleanUp

        outRow = outRow + 1
        arrOut(outRow, 1) = strCode
        arrOut(outRow, 2) = strName
        arrOut(outRow, 3) = LookupCorpInfo(dictInfo, Array("上市日期：", "上市日期", "上市日期:"))
        ' 所属行业 不在 comInfo1, 要拉 CorpOtherInfo/menu_num/2 子页面 (限速另算)
        arrOut(outRow, 4) = FetchIndustry(strCode)
        arrOut(outRow, 5) = LookupCorpInfo(dictInfo, Array("主营业务：", "经营范围：", "主营业务", "经营范围"))
        intSuccessCnt = intSuccessCnt + 1

        ' 上市日期转 Date 类型
        On Error Resume Next
        Dim dStr As String: dStr = CStr(arrOut(outRow, 3))
        If Len(dStr) >= 10 Then arrOut(outRow, 3) = CDate(Left$(dStr, 10))
        Err.Clear
        On Error GoTo CleanUp

        Application.Wait Now + TimeSerial(0, 0, 1)

NextRow:
    Next i

    If outRow > 0 Then
        wsTarget.Range("A2").Resize(outRow, 5).Value = arrOut
        wsTarget.Range("C2:C" & 1 + outRow).NumberFormat = "yyyy-mm-dd"
        wsTarget.Range("E2:E" & 1 + outRow).WrapText = True
        Call SetBorderLine(wsTarget.Range("A1:E" & 1 + outRow))
    End If

CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    g_globalFails = g_globalFails + intFailCnt
    If Len(strErrLog) > 0 Then _
        g_globalLog = g_globalLog & vbCrLf & "[基本资料]" & strErrLog

    If Not blnSilent And Not g_silentMode Then
        Dim msg As String
        msg = "上市公司基本资料 抓取完成" & vbCrLf & _
              "用时: " & Format(Timer - dtTime, "0.0 秒") & vbCrLf & _
              "成功: " & intSuccessCnt & " 家"
        If intFailCnt > 0 Then msg = msg & vbCrLf & vbCrLf & "失败 " & intFailCnt & " 条:" & strErrLog

        Dim style As Long: style = vbInformation
        If intFailCnt > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询"
    End If
End Sub


' --------- 在 corp info 字典里按多个候选 key 查值 ---------
Private Function LookupCorpInfo(ByVal dict As Object, ByVal arrKeys As Variant) As String
    Dim k As Variant
    For Each k In arrKeys
        If dict.Exists(CStr(k)) Then
            LookupCorpInfo = CStr(dict(CStr(k)))
            Exit Function
        End If
    Next k
    LookupCorpInfo = ""
End Function


' --------- 拉公司所属行业 (不在 comInfo1, 在 vCI_CorpOtherInfo 页) ---------
'   失败时返回空字符串, 不抛异常
Private Function FetchIndustry(ByVal strCode As String) As String
    On Error Resume Next
    Dim strUrl As String, strHtml As String
    strUrl = "http://vip.stock.finance.sina.com.cn/corp/go.php/vCI_CorpOtherInfo/stockid/" _
              & strCode & "/menu_num/2.phtml"
    strHtml = HttpGet(strUrl)
    If Err.Number <> 0 Then
        Err.Clear
        FetchIndustry = ""
        Exit Function
    End If

    ' 塌空白方便正则
    Dim objRegx As Object: Set objRegx = CreateObject("VBScript.Regexp")
    objRegx.Global = True
    objRegx.Pattern = "\s+"
    strHtml = objRegx.Replace(strHtml, " ")

    ' Sina 的 HTML 结构:
    '   <td colspan=2>所属行业板块</td></tr>           ← section header
    '   <tr><th>所属行业板块</th><th>同行业个股</th></tr>  ← col headers
    '   <tr><td>其他电子</td>...                        ← data (要拿这一格)
    ' 用 "<th>所属行业板块</th><th>同行业个股</th>" 这条线索锚定到数据 td
    objRegx.Global = False
    objRegx.Pattern = "所属行业板块[^<]*</[a-z]+>\s*<[a-z]+[^>]*>同行业个股[^<]*</[a-z]+>\s*</tr>\s*<tr[^>]*>\s*<td[^>]*>([^<]+)</td>"
    Dim matches As Object
    Set matches = objRegx.Execute(strHtml)
    If matches.Count > 0 Then
        FetchIndustry = Trim$(matches(0).SubMatches(0))
    Else
        FetchIndustry = ""
    End If

    ' 限速 1s
    Application.Wait Now + TimeSerial(0, 0, 1)
End Function


' --------- 报告期降序排序 ---------
'   输入 dictPeriodSet (key=period_str), 返回降序排好的一维数组(1-based)
Public Function SortPeriodsDesc(ByVal dictPeriodSet As Object) As Variant
    Dim arr() As String
    Dim n As Long: n = dictPeriodSet.Count
    If n = 0 Then
        SortPeriodsDesc = Array()
        Exit Function
    End If

    ReDim arr(1 To n)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In dictPeriodSet.Keys
        i = i + 1
        arr(i) = CStr(k)
    Next k

    ' Bubble sort 降序 (n 通常 < 30, 性能不重要)
    Dim j As Long, tmp As String
    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) < arr(j + 1) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i

    SortPeriodsDesc = arr
End Function


' --------- 指标按插入顺序输出 ---------
Public Function IndicatorsByInsertion(ByVal dictIndicatorSet As Object) As Variant
    Dim n As Long: n = dictIndicatorSet.Count
    If n = 0 Then
        IndicatorsByInsertion = Array()
        Exit Function
    End If
    Dim arr() As String
    ReDim arr(1 To n)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In dictIndicatorSet.Keys
        i = i + 1
        arr(i) = CStr(k)
    Next k
    IndicatorsByInsertion = arr
End Function


' --------- 按代码推断交易所前缀 (sh / sz) ---------
Public Function ExchangePrefix(ByVal strCode As String) As String
    Dim s As String: s = Trim$(strCode)
    If Len(s) = 0 Then ExchangePrefix = "": Exit Function

    Dim ch As String: ch = Left$(s, 1)
    If ch = "6" Then
        ExchangePrefix = "sh"
    ElseIf ch = "0" Or ch = "3" Then
        ExchangePrefix = "sz"
    ElseIf ch = "8" Or ch = "4" Then
        ExchangePrefix = "bj"
    Else
        ExchangePrefix = ""
    End If
End Function


' --------- 按代码 + 类型 + 年份 自拼新浪单张报表 URL ---------
'  strType: "balance" / "profit" / "cash" / "indicator"
'  lngYear: 0=取最新季度 (/ctrl/part/), >0=该年报告 (/ctrl/{year}/)
'  注意: 新浪 vFD_* 只认 6 位裸代码 (如 300866), 别加 sz/sh 前缀
Public Function BuildSinaFinancialUrl(ByVal strCode As String, _
                                       ByVal strType As String, _
                                       ByVal lngYear As Long) As String
    Dim strBase As String
    Select Case strType
        Case "balance"
            strBase = "http://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/"
        Case "profit"
            strBase = "http://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/"
        Case "cash"
            strBase = "http://money.finance.sina.com.cn/corp/go.php/vFD_CashFlow/stockid/"
        Case "indicator"
            strBase = "http://vip.stock.finance.sina.com.cn/corp/go.php/vFD_FinancialGuideLine/stockid/"
        Case Else
            Err.Raise vbObjectError + 540, "BuildSinaFinancialUrl", _
                "Unknown statement type: " & strType
    End Select

    Dim strCtrl As String
    If lngYear > 0 Then
        strCtrl = "ctrl/" & lngYear
    Else
        strCtrl = "ctrl/part"
    End If

    BuildSinaFinancialUrl = strBase & strCode & "/" & strCtrl & "/displaytype/4.phtml"
End Function


' --------- 给 Range 加细边框 ---------
Public Sub SetBorderLine(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(128, 128, 128)
    End With
End Sub

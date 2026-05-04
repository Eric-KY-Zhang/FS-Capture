Attribute VB_Name = "模块_抓美股财报"
Option Explicit

' =================================================================
'  抓美股财报 — 共享 helpers (RunUSStatement / Fetch / IsMatch / NzStr)
'
'  数据源: SEC EDGAR companyfacts XBRL JSON
'    https://data.sec.gov/api/xbrl/companyfacts/CIK{10位补零}.json
'
'  字段映射: us-gaap concepts → 中文指标名
'    3 元素: (大类, 标签, concept)            - 默认单位 USD, 缩放 1e6 (= 百万美元)
'    5 元素: (大类, 标签, concept, unit_key, scale)  - 自定义 (用于 EPS 等非美元单位指标)
'  报告期: 默认取最近 6 个; 按 (fy 年份 + fp 季度) 过滤, 每 (concept, fy, fp) 取 end 最大 + start 最小的 entry
'  各市场入口: 模块_抓美股资产负债表 / 模块_抓美股利润表 / 模块_抓美股现金流量表 / 模块_抓美股指标表
' =================================================================


' --------- 通用美股单表抓数主流程 ---------
'   strKind        : "BalanceSheet" / "Income" / "CashFlow" / "Indicator"
'   targetSheet    : 目标 sheet 名
'   conceptMap     : 1-d Variant 数组, 每元素是 Array(大类, 标签, concept[, unit_key, scale])
'   maxPeriods     : 取最近 N 个报告期 (default 6)
Public Sub RunUSStatement(ByVal strKind As String, ByVal targetSheet As String, _
                           ByVal conceptMap As Variant, _
                           ByVal maxPeriods As Long)
    Dim wsPool As Worksheet, wsTarget As Worksheet
    Dim arrPool As Variant
    Dim i As Long, lngRow As Long, numCompanies As Long
    Dim intFailCnt As Long, strErrLog As String
    Dim strCode As String, strName As String
    Dim dtTime As Double: dtTime = Timer

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictPeriodSet As Object: Set dictPeriodSet = CreateObject("Scripting.Dictionary")
    Dim dictIndicatorSet As Object: Set dictIndicatorSet = CreateObject("Scripting.Dictionary")
    Dim dictCategoryMap As Object: Set dictCategoryMap = CreateObject("Scripting.Dictionary")
    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    Dim dictReportingCurrency As Object: Set dictReportingCurrency = CreateObject("Scripting.Dictionary")
    Dim collCodes As New Collection
    Dim collDiagRows As New Collection

    g_diagnosticSheetName = "美股_抓取诊断"

    Set wsPool = ThisWorkbook.Sheets("样本池")
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)

    On Error GoTo CleanUp
    Application.ScreenUpdating = False

    lngRow = wsPool.Cells(wsPool.Rows.Count, POOL_US_CODE_COL).End(xlUp).Row
    If lngRow < POOL_DATA_START_ROW Then
        intFailCnt = 1
        strErrLog = "样本池美股区为空"
        GoTo CleanUp
    End If
    arrPool = wsPool.Range(wsPool.Cells(POOL_DATA_START_ROW, POOL_US_CODE_COL), _
                           wsPool.Cells(lngRow, POOL_US_NAME_COL)).Value
    If Not IsArray(arrPool) Then
        Dim singleVal As Variant: singleVal = arrPool
        ReDim arrPool(1 To 1, 1 To 2)
        arrPool(1, 1) = singleVal
    End If
    numCompanies = UBound(arrPool, 1)

    ' 提前加载 ticker→CIK 映射 (一次会话只下一次)
    Application.StatusBar = "加载 SEC ticker→CIK 映射..."
    DoEvents
    LoadTickerCIKMap

    ' 季度选择 (Bug A 修复: 美股按 form+fp 在 entry 层过滤, 不走字符串后缀)
    Dim strQuarter As String: strQuarter = ReadQuarterSelection()
    ' 年份选择 (A2): 0 = 留空 → 拉最近 N 期; > 0 → 按 EDGAR fy 字段精确过滤
    Dim lngYear As Long: lngYear = ReadYearSelection()
    Dim collFetchYears As Collection: Set collFetchYears = New Collection
    collFetchYears.Add lngYear
    If lngYear > 0 And (strKind = "BalanceSheet" Or strKind = "Income") Then _
        collFetchYears.Add lngYear - 1

    For i = 1 To numCompanies
        strCode = Trim$(CStr(arrPool(i, 1)))
        If Len(strCode) = 0 Then GoTo NextRow
        strName = Trim$(CStr(arrPool(i, 2)))

        Application.StatusBar = "抓取中: " & targetSheet & " (" & i & "/" & numCompanies & ") " & strCode
        DoEvents

        Dim fetchYear As Variant
        Dim mainYearOk As Boolean: mainYearOk = False
        Dim anyYearOk As Boolean: anyYearOk = False
        Dim mainErrNum As Long: mainErrNum = 0
        Dim mainErrSource As String: mainErrSource = ""
        Dim mainErrDesc As String: mainErrDesc = ""

        For Each fetchYear In collFetchYears
            On Error Resume Next
            Err.Clear
            Call FetchAndAccumulateUSCompany(strCode, conceptMap, strQuarter, CLng(fetchYear), _
                                              dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap, _
                                              strKind, collDiagRows, dictReportingCurrency)
            If Err.Number <> 0 Then
                If CLng(fetchYear) = lngYear Then
                    mainErrNum = Err.Number
                    mainErrSource = Err.Source
                    mainErrDesc = Err.Description
                End If
                Err.Clear
            Else
                anyYearOk = True
                If CLng(fetchYear) = lngYear Then mainYearOk = True
            End If
            On Error GoTo CleanUp

            Application.Wait Now + TimeSerial(0, 0, 1)    ' SEC / 雪球限速 1s
        Next fetchYear

        If (lngYear = 0 And Not anyYearOk) Or (lngYear > 0 And Not mainYearOk) Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": " & _
                        "错误号=" & mainErrNum & "; 来源=" & mainErrSource & "; " & mainErrDesc
            AddMissingDiagnosticsForCompany strCode, strKind, conceptMap, collDiagRows, _
                                            "fetch_failed: " & mainErrDesc
        Else
            If Not dictCompanyName.Exists(strCode) Then dictCompanyName.Add strCode, strName
            collCodes.Add strCode
        End If

NextRow:
    Next i

    If collCodes.Count = 0 Then
        ClearUSWideTableOutput wsTarget
        GoTo CleanUp
    End If

    ' 防御: 没匹配到任何 entry (concept 名不对 / fy/fp filter 太严) 就别去 WriteWideTable
    If dictPeriodSet.Count = 0 Or dictIndicatorSet.Count = 0 Then
        intFailCnt = intFailCnt + 1
        strErrLog = strErrLog & vbCrLf & "[" & targetSheet & "] " & _
                    "无匹配数据 (期数=" & dictPeriodSet.Count & _
                    " / 指标=" & dictIndicatorSet.Count & _
                    ", 检查 A2 年份 + A4 季度 + 概念名是否对得上 EDGAR)"
        GoTo CleanUp
    End If

    Dim arrCodes() As String
    ReDim arrCodes(1 To collCodes.Count)
    For i = 1 To collCodes.Count
        arrCodes(i) = collCodes(i)
    Next i

    ' Bug A 修复: 不走字符串后缀过滤 — entry 层已用 form+fp 过滤过

    ' 取最近 maxPeriods 个 (降序后截断)
    Dim arrPeriods As Variant, arrIndicators As Variant
    arrPeriods = SortPeriodsDesc(dictPeriodSet)
    If IsArray(arrPeriods) Then
        On Error Resume Next
        Dim ubP As Long: ubP = -1
        ubP = UBound(arrPeriods)
        On Error GoTo CleanUp
        If lngYear = 0 And ubP - LBound(arrPeriods) + 1 > maxPeriods Then
            ReDim Preserve arrPeriods(LBound(arrPeriods) To LBound(arrPeriods) + maxPeriods - 1)
        End If
    End If

    ' 指标排序 = 按 conceptMap 顺序 (跳过未抓到的)
    arrIndicators = OrderIndicatorsByConceptMap(dictIndicatorSet, conceptMap)

    Application.StatusBar = "写入: " & targetSheet
    DoEvents
    Dim usCode As Variant
    For Each usCode In arrCodes
        If Not dictReportingCurrency.Exists(CStr(usCode)) Then dictReportingCurrency(CStr(usCode)) = "USD"
    Next usCode

    Dim hookKind As String
    Select Case True
        Case InStr(targetSheet, "资产负债") > 0: hookKind = "BalanceSheet"
        Case InStr(targetSheet, "利润") > 0:     hookKind = "Income"
        Case InStr(targetSheet, "现金流") > 0:   hookKind = "CashFlow"
        Case Else:                               hookKind = ""
    End Select

    WriteWideTable wsTarget, arrCodes, dictCompanyName, dictData, _
                    arrPeriods, arrIndicators, dictCategoryMap, _
                    perCompanyPeriods:=True, _
                    dictReportingCurrency:=dictReportingCurrency, _
                    statementKind:=hookKind
    RefreshA1CurrencyComment wsTarget, targetSheet

CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    On Error Resume Next
    WriteDiagnosticForKind strKind, collDiagRows
    Err.Clear
    On Error GoTo 0

    g_globalFails = g_globalFails + intFailCnt
    If Len(strErrLog) > 0 Then _
        g_globalLog = g_globalLog & vbCrLf & "[" & targetSheet & "]" & strErrLog

    If Not g_silentMode Then
        Dim msg As String
        msg = targetSheet & " 抓取完成 (单位: 百万美元)" & vbCrLf & _
              "用时: " & Format(Timer - dtTime, "0.0 秒") & vbCrLf & _
              "公司数: " & collCodes.Count & " / 期数: " & dictPeriodSet.Count
        If intFailCnt > 0 Then msg = msg & vbCrLf & vbCrLf & "失败 " & intFailCnt & " 条:" & strErrLog

        Dim style As Long: style = vbInformation
        If intFailCnt > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询 (US)"
    End If
End Sub


Private Sub ClearUSWideTableOutput(ByVal ws As Worksheet)
    Dim metaCols As Long: metaCols = 2
    If ws.Name = "美股_指标表" Then metaCols = 3

    On Error Resume Next
    ws.UsedRange.UnMerge
    On Error GoTo 0

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
    If lastRow < 2 Then lastRow = 2
    If lastCol < metaCols + 1 Then lastCol = metaCols + 1

    ws.Range(ws.Cells(1, metaCols + 1), ws.Cells(lastRow, lastCol)).Clear
    If lastRow >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, metaCols)).Clear

    If metaCols = 3 Then
        ws.Range("A1").Value = "指标类型"
        ws.Range("C1").Value = "英文指标名"
    Else
        ws.Range("A1").Value = "大类"
    End If
    ws.Range("B1").Value = "指标名称"
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, metaCols))
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Call SetBorderLine(ws.Range(ws.Cells(1, 1), ws.Cells(2, metaCols + 1)))
End Sub


' --------- 抓单家美股公司 + 累计到字典 ---------
'   strQuarter: "全部"/"Q1"/"Q2"/"Q3"/"Q4" — 按 EDGAR form+fp 在 entry 层过滤
'   lngYear   : 0=不过滤; >0=只保留 EDGAR fy 字段匹配的 entry (= 该公司 fiscal year)
'   strKind   : "BalanceSheet"/"Income"/"CashFlow"/"Indicator" — 决定是否走 xueqiu fallback
Private Sub FetchAndAccumulateUSCompany(ByVal strTicker As String, _
                                         ByVal conceptMap As Variant, _
                                         ByVal strQuarter As String, _
                                         ByVal lngYear As Long, _
                                         ByRef dictData As Object, _
                                         ByRef dictPeriodSet As Object, _
                                         ByRef dictIndicatorSet As Object, _
                                         ByRef dictCategoryMap As Object, _
                                         Optional ByVal strKind As String = "BalanceSheet", _
                                         Optional ByVal collDiagRows As Collection = Nothing, _
                                         Optional ByRef dictReportingCurrency As Object = Nothing)
    Dim strCIK As String, strUrl As String, strJson As String

    ' Phase 4b-14a: EDGAR 先写临时字典; 只有确认不用雪球 fallback 后才合并到正式输出。
    On Error Resume Next
    Err.Clear
    strCIK = LookupCIK(strTicker)
    strUrl = "https://data.sec.gov/api/xbrl/companyfacts/CIK" & strCIK & ".json"
    strJson = CachedEdgarHttpGet(strUrl, "edgar_US_" & UCase$(strTicker) & "_companyfacts")
    Dim edgarErrNum As Long: edgarErrNum = Err.Number
    Dim edgarErrDesc As String: edgarErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If edgarErrNum <> 0 Then
        If XueqiuSupportsKind(strKind) Then
            FetchUSFallbackAfterXueqiu strTicker, strKind, conceptMap, strQuarter, lngYear, _
                                       dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap, _
                                       collDiagRows, dictReportingCurrency, edgarErrDesc
            Exit Sub
        Else
            Err.Raise vbObjectError + 526, "FetchUS", edgarErrDesc
        End If
    End If

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(strJson)

    If Not parsed.Exists("facts") Then
        If XueqiuSupportsKind(strKind) Then
            FetchUSFallbackAfterXueqiu strTicker, strKind, conceptMap, strQuarter, lngYear, _
                                       dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap, _
                                       collDiagRows, dictReportingCurrency, "JSON 缺少 facts"
            Exit Sub
        Else
            Err.Raise vbObjectError + 530, "FetchUS", "JSON 缺少 facts: " & strTicker
        End If
    End If
    Dim factsRoot As Object: Set factsRoot = parsed("facts")

    Dim tempData As Object: Set tempData = CreateObject("Scripting.Dictionary")
    Dim tempPeriodSet As Object: Set tempPeriodSet = CreateObject("Scripting.Dictionary")
    Dim tempIndicatorSet As Object: Set tempIndicatorSet = CreateObject("Scripting.Dictionary")
    Dim tempCategoryMap As Object: Set tempCategoryMap = CreateObject("Scripting.Dictionary")
    Dim dictMatchInfo As Object: Set dictMatchInfo = CreateObject("Scripting.Dictionary")
    Dim fuzzyTaxonomy As Object, fuzzyTaxonomyName As String

    If factsRoot.Exists("us-gaap") Then
        Dim usGaap As Object: Set usGaap = factsRoot.Item("us-gaap")
        AccumulateEdgarTaxonomy strTicker, conceptMap, usGaap, "us-gaap", "EDGAR us-gaap", _
                                strQuarter, lngYear, tempData, tempPeriodSet, _
                                tempIndicatorSet, tempCategoryMap, dictMatchInfo
        Set fuzzyTaxonomy = usGaap
        fuzzyTaxonomyName = "us-gaap"
    End If

    If HasAnyCoreLabel(tempData, strTicker, CoreLabelsForKind(strKind)) Then
        CommitTempUSData strTicker, tempData, tempPeriodSet, tempIndicatorSet, tempCategoryMap, _
                         dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap
        If Not dictReportingCurrency Is Nothing Then dictReportingCurrency(strTicker) = "USD"
        AppendDiagnosticsForConceptMap strTicker, strKind, conceptMap, dictMatchInfo, _
                                       fuzzyTaxonomy, fuzzyTaxonomyName, collDiagRows
        Exit Sub
    End If

    If factsRoot.Exists("ifrs-full") Then
        Dim ifrsFull As Object: Set ifrsFull = factsRoot.Item("ifrs-full")
        AccumulateEdgarTaxonomy strTicker, conceptMap, ifrsFull, "ifrs-full", "EDGAR ifrs-full", _
                                strQuarter, lngYear, tempData, tempPeriodSet, _
                                tempIndicatorSet, tempCategoryMap, dictMatchInfo
        Set fuzzyTaxonomy = ifrsFull
        fuzzyTaxonomyName = "ifrs-full"
    End If

    If HasAnyCoreLabel(tempData, strTicker, CoreLabelsForKind(strKind)) Then
        CommitTempUSData strTicker, tempData, tempPeriodSet, tempIndicatorSet, tempCategoryMap, _
                         dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap
        If Not dictReportingCurrency Is Nothing Then dictReportingCurrency(strTicker) = "USD"
        AppendDiagnosticsForConceptMap strTicker, strKind, conceptMap, dictMatchInfo, _
                                       fuzzyTaxonomy, fuzzyTaxonomyName, collDiagRows
        Exit Sub
    End If

    AppendNonUsdDiagnostics strTicker, strKind, conceptMap, dictMatchInfo, collDiagRows
    If XueqiuSupportsKind(strKind) Then
        FetchUSFallbackAfterXueqiu strTicker, strKind, conceptMap, strQuarter, lngYear, _
                                   dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap, _
                                   collDiagRows, dictReportingCurrency, "no core EDGAR fields matched"
    Else
        AddMissingDiagnosticsForCompany strTicker, strKind, conceptMap, collDiagRows, _
                                        "no core EDGAR fields matched"
    End If
End Sub


Private Function AccumulateEdgarTaxonomy(ByVal strTicker As String, _
                                         ByVal conceptMap As Variant, _
                                         ByVal taxonomy As Object, _
                                         ByVal taxonomyName As String, _
                                         ByVal sourceName As String, _
                                         ByVal strQuarter As String, _
                                         ByVal lngYear As Long, _
                                         ByRef dictData As Object, _
                                         ByRef dictPeriodSet As Object, _
                                         ByRef dictIndicatorSet As Object, _
                                         ByRef dictCategoryMap As Object, _
                                         ByRef dictMatchInfo As Object) As Long
    If taxonomy Is Nothing Then Exit Function

    Dim dictCompany As Object
    If dictData.Exists(strTicker) Then
        Set dictCompany = dictData.Item(strTicker)
    Else
        Set dictCompany = CreateObject("Scripting.Dictionary")
        dictData.Add strTicker, dictCompany
    End If

    Dim i As Long, mapEntry As Variant
    For i = LBound(conceptMap) To UBound(conceptMap)
        mapEntry = conceptMap(i)
        Dim strCat As String: strCat = CStr(mapEntry(0))
        Dim strLabel As String: strLabel = CStr(mapEntry(1))
        If dictMatchInfo.Exists(strLabel) Then GoTo NextMapEntry

        Dim conceptCsv As String
        If taxonomyName = "ifrs-full" Then
            conceptCsv = MapEntryIfrsConcepts(mapEntry)
        Else
            conceptCsv = MapEntryUsGaapConcepts(mapEntry)
        End If

        Dim strUnit As String: strUnit = MapEntryUnit(mapEntry)
        Dim dblScale As Double: dblScale = MapEntryScale(mapEntry)
        Dim conceptCandidates As Variant: conceptCandidates = Split(conceptCsv, ",")
        Dim conceptCandidate As Variant, candIdx As Long, totalCand As Long
        candIdx = 0
        totalCand = UBound(conceptCandidates) - LBound(conceptCandidates) + 1

        For Each conceptCandidate In conceptCandidates
            candIdx = candIdx + 1
            Dim conceptName As String: conceptName = Trim$(CStr(conceptCandidate))
            If Len(conceptName) = 0 Then GoTo NextConceptCandidate

            If taxonomy.Exists(conceptName) Then
                Dim conceptObj As Object: Set conceptObj = taxonomy.Item(conceptName)
                If conceptObj.Exists("units") Then
                    Dim units As Object: Set units = conceptObj.Item("units")
                    If units.Exists(strUnit) Then
                        Dim entries As Object: Set entries = units.Item(strUnit)
                        Dim dictCanonical As Object
                        Set dictCanonical = CanonicalEdgarEntries(entries, strQuarter, lngYear)

                        If dictCanonical.Count > 0 Then
                            Dim ck As Variant
                            Dim fxPeriodEnd As String: fxPeriodEnd = ""
                            For Each ck In dictCanonical.Keys
                                Dim pair As Variant: pair = dictCanonical.Item(ck)
                                Dim canonEnd As String: canonEnd = CStr(pair(0))
                                Dim canonVal As Variant: canonVal = CDbl(pair(2)) / dblScale
                                If canonEnd > fxPeriodEnd Then fxPeriodEnd = canonEnd

                                If Not dictPeriodSet.Exists(canonEnd) Then dictPeriodSet.Add canonEnd, True
                                If Not dictIndicatorSet.Exists(strLabel) Then
                                    dictIndicatorSet.Add strLabel, dictIndicatorSet.Count
                                    If Not dictCategoryMap.Exists(strLabel) Then dictCategoryMap.Add strLabel, strCat
                                End If

                                Dim dictPer As Object
                                If dictCompany.Exists(canonEnd) Then
                                    Set dictPer = dictCompany.Item(canonEnd)
                                Else
                                    Set dictPer = CreateObject("Scripting.Dictionary")
                                    dictCompany.Add canonEnd, dictPer
                                End If
                                dictPer(strLabel) = canonVal
                            Next ck

                            Dim statusText As String
                            If taxonomyName = "ifrs-full" Then
                                statusText = "OK_IFRS"
                            Else
                                statusText = "OK"
                            End If
                            Dim scoreText As String
                            Dim noteBase As String
                            If candIdx = 1 Then
                                scoreText = "100"
                                noteBase = "hardcoded_primary"
                            Else
                                scoreText = "85"
                                noteBase = "hardcoded_alt[" & candIdx & "/" & totalCand & "]"
                            End If
                            dictMatchInfo(strLabel) = Array(statusText, sourceName, taxonomyName, _
                                                            conceptName, strUnit, scoreText, noteBase, _
                                                            CLng(dictCanonical.Count), fxPeriodEnd)
                            AccumulateEdgarTaxonomy = AccumulateEdgarTaxonomy + dictCanonical.Count
                            Exit For
                        ElseIf Not dictMatchInfo.Exists("NOPERIOD|" & strLabel) Then
                            Dim noPeriodScore As String
                            Dim noPeriodNote As String
                            If candIdx = 1 Then
                                noPeriodScore = "100"
                                noPeriodNote = "hardcoded_primary; exact concept matched but no period entry"
                            Else
                                noPeriodScore = "85"
                                noPeriodNote = "hardcoded_alt[" & candIdx & "/" & totalCand & _
                                               "]; exact concept matched but no period entry"
                            End If
                            dictMatchInfo.Add "NOPERIOD|" & strLabel, _
                                Array("MISSING", sourceName, taxonomyName, conceptName, _
                                      strUnit, noPeriodScore, noPeriodNote, 0, "")
                        End If
                    ElseIf taxonomyName = "ifrs-full" Then
                        Dim actualUnit As String: actualUnit = FirstUnitKey(units)
                        If Len(actualUnit) > 0 And Not dictMatchInfo.Exists("NONUSD|" & strLabel) Then
                            dictMatchInfo.Add "NONUSD|" & strLabel, _
                                Array("MISSING_NON_USD", sourceName, taxonomyName, conceptName, _
                                      actualUnit, "—", "non_usd_unit", 0, "")
                        End If
                    End If
                End If
            End If
NextConceptCandidate:
        Next conceptCandidate
NextMapEntry:
    Next i
End Function


Private Function CanonicalEdgarEntries(ByVal entries As Object, ByVal strQuarter As String, _
                                       ByVal lngYear As Long) As Object
    Dim dictCanonical As Object: Set dictCanonical = CreateObject("Scripting.Dictionary")
    Dim j As Long
    For j = 1 To entries.Count
        Dim e As Object: Set e = entries.Item(j)
        If IsEdgarEntryMatch(e, strQuarter, lngYear) Then
            Dim periodEnd As String: periodEnd = NzStr(e, "end")
            Dim periodStart As String: periodStart = NzStr(e, "start")
            If Len(periodEnd) > 0 Then
                Dim ckey As String: ckey = NzStr(e, "fy") & "|" & NzStr(e, "fp")
                Dim takeIt As Boolean: takeIt = False
                If Not dictCanonical.Exists(ckey) Then
                    takeIt = True
                Else
                    Dim prev As Variant: prev = dictCanonical.Item(ckey)
                    Dim prevEnd As String: prevEnd = CStr(prev(0))
                    Dim prevStart As String: prevStart = CStr(prev(1))
                    If periodEnd > prevEnd Then
                        takeIt = True
                    ElseIf periodEnd = prevEnd And Len(periodStart) > 0 _
                            And (Len(prevStart) = 0 Or periodStart < prevStart) Then
                        takeIt = True
                    End If
                End If
                If takeIt Then dictCanonical(ckey) = Array(periodEnd, periodStart, e("val"))
            End If
        End If
    Next j
    Set CanonicalEdgarEntries = dictCanonical
End Function


Private Function HasAnyCoreLabel(ByVal dictData As Object, ByVal strTicker As String, _
                                 ByVal coreLabels As Variant) As Boolean
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(coreLabels): ub = UBound(coreLabels)
    If Err.Number <> 0 Or ub < lb Then
        HasAnyCoreLabel = (dictData.Exists(strTicker) And dictData.Item(strTicker).Count > 0)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If Not dictData.Exists(strTicker) Then Exit Function
    Dim dictCompany As Object: Set dictCompany = dictData.Item(strTicker)
    Dim p As Variant, i As Long
    For Each p In dictCompany.Keys
        Dim dictPer As Object: Set dictPer = dictCompany.Item(p)
        For i = lb To ub
            If dictPer.Exists(CStr(coreLabels(i))) Then
                HasAnyCoreLabel = True
                Exit Function
            End If
        Next i
    Next p
End Function


Private Sub CommitTempUSData(ByVal strTicker As String, _
                             ByVal tempData As Object, _
                             ByVal tempPeriodSet As Object, _
                             ByVal tempIndicatorSet As Object, _
                             ByVal tempCategoryMap As Object, _
                             ByRef dictData As Object, _
                             ByRef dictPeriodSet As Object, _
                             ByRef dictIndicatorSet As Object, _
                             ByRef dictCategoryMap As Object)
    If Not tempData.Exists(strTicker) Then Exit Sub

    Dim destCompany As Object
    If dictData.Exists(strTicker) Then
        Set destCompany = dictData.Item(strTicker)
    Else
        Set destCompany = CreateObject("Scripting.Dictionary")
        dictData.Add strTicker, destCompany
    End If

    Dim tempCompany As Object: Set tempCompany = tempData.Item(strTicker)
    Dim p As Variant, label As Variant
    For Each p In tempCompany.Keys
        Dim destPer As Object
        If destCompany.Exists(CStr(p)) Then
            Set destPer = destCompany.Item(CStr(p))
        Else
            Set destPer = CreateObject("Scripting.Dictionary")
            destCompany.Add CStr(p), destPer
        End If

        Dim tempPer As Object: Set tempPer = tempCompany.Item(p)
        For Each label In tempPer.Keys
            destPer(CStr(label)) = tempPer.Item(label)
        Next label
    Next p

    Dim k As Variant
    For Each k In tempPeriodSet.Keys
        If Not dictPeriodSet.Exists(CStr(k)) Then dictPeriodSet.Add CStr(k), True
    Next k
    For Each k In tempIndicatorSet.Keys
        If Not dictIndicatorSet.Exists(CStr(k)) Then dictIndicatorSet.Add CStr(k), dictIndicatorSet.Count
    Next k
    For Each k In tempCategoryMap.Keys
        If Not dictCategoryMap.Exists(CStr(k)) Then dictCategoryMap.Add CStr(k), tempCategoryMap.Item(k)
    Next k
End Sub


Private Sub AppendDiagnosticsForConceptMap(ByVal strTicker As String, _
                                           ByVal strKind As String, _
                                           ByVal conceptMap As Variant, _
                                           ByVal dictMatchInfo As Object, _
                                           ByVal fuzzyTaxonomy As Object, _
                                           ByVal fuzzyTaxonomyName As String, _
                                           ByVal collDiagRows As Collection)
    If collDiagRows Is Nothing Then Exit Sub
    Dim i As Long, mapEntry As Variant, strLabel As String
    For i = LBound(conceptMap) To UBound(conceptMap)
        mapEntry = conceptMap(i)
        strLabel = CStr(mapEntry(1))
        If dictMatchInfo.Exists(strLabel) Then
            AppendMatchDiagnostic strTicker, strKind, strLabel, dictMatchInfo.Item(strLabel), collDiagRows
        ElseIf dictMatchInfo.Exists("NOPERIOD|" & strLabel) Then
            AppendMatchDiagnostic strTicker, strKind, strLabel, dictMatchInfo.Item("NOPERIOD|" & strLabel), collDiagRows
        ElseIf dictMatchInfo.Exists("NONUSD|" & strLabel) Then
            AppendMatchDiagnostic strTicker, strKind, strLabel, dictMatchInfo.Item("NONUSD|" & strLabel), collDiagRows
        Else
            AddDiagnosticRow collDiagRows, strTicker, strKind, strLabel, "MISSING", "—", "—", "—", "—", "—", _
                             "no exact candidate matched"
            AddFuzzyDiagnosticCandidates collDiagRows, strTicker, strKind, strLabel, fuzzyTaxonomyName, _
                                         fuzzyTaxonomy, MapEntryFuzzyHint(mapEntry), MapEntryUnit(mapEntry)
        End If
    Next i
End Sub


Private Sub AppendNonUsdDiagnostics(ByVal strTicker As String, _
                                    ByVal strKind As String, _
                                    ByVal conceptMap As Variant, _
                                    ByVal dictMatchInfo As Object, _
                                    ByVal collDiagRows As Collection)
    If collDiagRows Is Nothing Then Exit Sub
    Dim i As Long, strLabel As String
    For i = LBound(conceptMap) To UBound(conceptMap)
        strLabel = CStr(conceptMap(i)(1))
        If dictMatchInfo.Exists("NONUSD|" & strLabel) Then _
            AppendMatchDiagnostic strTicker, strKind, strLabel, dictMatchInfo.Item("NONUSD|" & strLabel), collDiagRows
    Next i
End Sub


Private Sub AppendMatchDiagnostic(ByVal strTicker As String, ByVal strKind As String, _
                                  ByVal strLabel As String, ByVal info As Variant, _
                                  ByVal collDiagRows As Collection)
    Dim noteText As String: noteText = CStr(info(6))
    If CLng(info(7)) > 0 Then noteText = noteText & "; periods_written=" & CStr(info(7))
    Dim periodEnd As String: periodEnd = ""
    If UBound(info) >= 8 Then periodEnd = CStr(info(8))
    Dim fxText As String: fxText = FxRateTextForDiagnostic("USD", periodEnd, strKind)
    AddDiagnosticRow collDiagRows, strTicker, strKind, strLabel, CStr(info(0)), CStr(info(1)), _
                     CStr(info(2)), CStr(info(3)), CStr(info(4)), CStr(info(5)), noteText, fxText
End Sub


Private Function FirstUnitKey(ByVal units As Object) As String
    Dim k As Variant
    For Each k In units.Keys
        FirstUnitKey = CStr(k)
        Exit Function
    Next k
End Function


' --------- null-safe 取 Dictionary string ---------
Private Function NzStr(ByVal dict As Object, ByVal key As String) As String
    If Not dict.Exists(key) Then NzStr = "": Exit Function
    Dim v As Variant: v = dict.Item(key)
    If IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function


' --------- 按 季度 + 财年 过滤 EDGAR entry ---------
'   全部 → 10-K (含/A) + 10-Q (含/A)
'   Q1/Q2/Q3 → 10-Q + fp 匹配
'   Q4 → 10-K (年报)
'   lngYear > 0 时, 进一步要求 fy 字段精确匹配
'   AAPL 财年 9 月底 — fy=2025 fp=FY 是 fiscal year 2025 年报 (end ≈ 2025-09-27)
Private Function IsEdgarEntryMatch(ByVal e As Object, ByVal strQuarter As String, _
                                     ByVal lngYear As Long) As Boolean
    If Not e.Exists("form") Then IsEdgarEntryMatch = False: Exit Function
    If Not e.Exists("end") Then IsEdgarEntryMatch = False: Exit Function
    If Not e.Exists("val") Then IsEdgarEntryMatch = False: Exit Function

    ' val 不能是 null
    Dim varVal As Variant: varVal = e("val")
    If IsNull(varVal) Then IsEdgarEntryMatch = False: Exit Function

    ' 如果用户在 A2 填了年份, 严格按 fy 过滤
    If lngYear > 0 Then
        If Not e.Exists("fy") Then IsEdgarEntryMatch = False: Exit Function
        Dim varFy As Variant: varFy = e("fy")
        If IsNull(varFy) Then IsEdgarEntryMatch = False: Exit Function
        If CLng(varFy) <> lngYear Then IsEdgarEntryMatch = False: Exit Function
    End If

    Dim strForm As String: strForm = NzStr(e, "form")
    Dim strFp As String: strFp = NzStr(e, "fp")

    Dim isAnnual As Boolean: isAnnual = (strForm = "10-K" Or strForm = "10-K/A")
    Dim isQuarter As Boolean: isQuarter = (strForm = "10-Q" Or strForm = "10-Q/A")

    Select Case strQuarter
        Case "全部", ""
            IsEdgarEntryMatch = (isAnnual Or isQuarter)
        Case "Q1"
            IsEdgarEntryMatch = (isQuarter And strFp = "Q1")
        Case "Q2"
            IsEdgarEntryMatch = (isQuarter And strFp = "Q2")
        Case "Q3"
            IsEdgarEntryMatch = (isQuarter And strFp = "Q3")
        Case "Q4"
            IsEdgarEntryMatch = isAnnual
        Case Else
            IsEdgarEntryMatch = (isAnnual Or isQuarter)
    End Select
End Function


' --------- 按 conceptMap 顺序输出已抓到的指标 ---------
Public Function OrderIndicatorsByConceptMap(ByVal dictIndicatorSet As Object, _
                                              ByVal conceptMap As Variant) As Variant
    Dim arr() As String, n As Long: n = 0
    Dim i As Long
    ReDim arr(1 To dictIndicatorSet.Count + 1)
    For i = LBound(conceptMap) To UBound(conceptMap)
        Dim label As String: label = CStr(conceptMap(i)(1))
        If dictIndicatorSet.Exists(label) Then
            n = n + 1
            arr(n) = label
        End If
    Next i
    If n = 0 Then
        OrderIndicatorsByConceptMap = Array()
    Else
        ReDim Preserve arr(1 To n)
        OrderIndicatorsByConceptMap = arr
    End If
End Function


' --------- 美股 指标表 追加 ratio 行 (Excel INDEX/MATCH 公式实时计算) ---------
'   依赖: 美股_资产负债表 + 美股_利润表 已经填好数据
'   公式按 当前列的公司表头 + 报告期 去 BS / IS 定位对应列, 避免跨表列错配
Public Sub AppendUSRatios(ByVal ws As Worksheet)
    ' 数据区列数 = R2 上最后一个非空 cell
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then Exit Sub        ' 没有抓到数据, 跳过

    ' 起始 row = B 列最后非空行 + 1 (在 EDGAR raw indicators 之后)
    Dim startRow As Long
    startRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1

    Dim ratios(0 To 7) As Variant
    ratios(0) = MakeRatio("Liquidity", "Current Ratio", _
        "=IFERROR({BS_TCA}/{BS_TCL},"""")", "0.00")
    ratios(1) = MakeRatio("Liquidity", "Quick Ratio", _
        "=IFERROR(({BS_TCA}-{BS_INV})/{BS_TCL},"""")", "0.00")
    ratios(2) = MakeRatio("Leverage", "Debt Ratio", _
        "=IFERROR({BS_TL}/{BS_TA},"""")", "0.00%")
    ratios(3) = MakeRatio("Profitability", "Gross Margin", _
        "=IFERROR({IS_GP}/{IS_REV},"""")", "0.00%")
    ratios(4) = MakeRatio("Profitability", "Operating Margin", _
        "=IFERROR({IS_OP}/{IS_REV},"""")", "0.00%")
    ratios(5) = MakeRatio("Profitability", "Net Margin", _
        "=IFERROR({IS_NI}/{IS_REV},"""")", "0.00%")
    ratios(6) = MakeRatio("Profitability", "ROA", _
        "=IFERROR({IS_NI}/{BS_TA},"""")", "0.00%")
    ratios(7) = MakeRatio("Profitability", "ROE", _
        "=IFERROR({IS_NI}/{BS_TSE},"""")", "0.00%")

    ' 解析 BS / IS sheet 上各指标的实际行号 (一次性查找, 之后填进公式)
    Dim rowMap As Object: Set rowMap = CreateObject("Scripting.Dictionary")
    rowMap.Add "BS_TCA", Array("美股_资产负债表", FindIndicatorRow("美股_资产负债表", "Total current assets"))
    rowMap.Add "BS_TCL", Array("美股_资产负债表", FindIndicatorRow("美股_资产负债表", "Total current liabilities"))
    rowMap.Add "BS_INV", Array("美股_资产负债表", FindIndicatorRow("美股_资产负债表", "Inventory"))
    rowMap.Add "BS_TA", Array("美股_资产负债表", FindIndicatorRow("美股_资产负债表", "Total assets"))
    rowMap.Add "BS_TL", Array("美股_资产负债表", FindIndicatorRow("美股_资产负债表", "Total liabilities"))
    rowMap.Add "BS_TSE", Array("美股_资产负债表", FindIndicatorRow("美股_资产负债表", "Total stockholders' equity"))
    rowMap.Add "IS_REV", Array("美股_利润表", FindIndicatorRow("美股_利润表", "Revenue"))
    rowMap.Add "IS_GP", Array("美股_利润表", FindIndicatorRow("美股_利润表", "Gross profit"))
    rowMap.Add "IS_OP", Array("美股_利润表", FindIndicatorRow("美股_利润表", "Operating income"))
    rowMap.Add "IS_NI", Array("美股_利润表", FindIndicatorRow("美股_利润表", "Net income"))

    Dim r As Long
    For r = LBound(ratios) To UBound(ratios)
        Dim rowI As Long: rowI = startRow + (r - LBound(ratios))
        Dim cat As String: cat = CStr(ratios(r)(0))
        Dim label As String: label = CStr(ratios(r)(1))
        Dim formulaTpl As String: formulaTpl = CStr(ratios(r)(2))
        Dim numFmt As String: numFmt = CStr(ratios(r)(3))

        ws.Cells(rowI, 1).Value = cat
        ws.Cells(rowI, 2).Value = label
        ws.Cells(rowI, 1).Font.Bold = True
        ws.Cells(rowI, 2).Font.Bold = True

        Dim c As Long
        For c = 3 To lastCol
            Dim companyHeader As String: companyHeader = HeaderTextAt(ws, c)
            Dim periodText As String: periodText = PeriodKey(ws.Cells(2, c).Value)
            Dim bsCol As Long: bsCol = FindStatementColumn("美股_资产负债表", companyHeader, periodText)
            Dim isCol As Long: isCol = FindStatementColumn("美股_利润表", companyHeader, periodText)
            Dim formula As String: formula = formulaTpl

            ' 替换占位符为同公司、同报告期、同指标的真实单元格引用
            Dim k As Variant
            Dim missingRef As Boolean: missingRef = False
            For Each k In rowMap.Keys
                If InStr(formula, "{" & k & "}") > 0 Then
                    Dim refMeta As Variant: refMeta = rowMap.Item(k)
                    Dim refSheet As String: refSheet = CStr(refMeta(0))
                    Dim refRow As Long: refRow = CLng(refMeta(1))
                    Dim refCol As Long
                    If refSheet = "美股_资产负债表" Then
                        refCol = bsCol
                    Else
                        refCol = isCol
                    End If

                    If refRow <= 0 Or refCol <= 0 Then
                        missingRef = True
                    Else
                        formula = Replace(formula, "{" & k & "}", SheetCellRef(refSheet, refRow, refCol))
                    End If
                End If
            Next k

            If missingRef Or InStr(formula, "{") > 0 Then
                ws.Cells(rowI, c).Value = ""
            Else
                ws.Cells(rowI, c).Formula = formula
                ws.Cells(rowI, c).NumberFormat = numFmt
            End If
        Next c
    Next r
End Sub


Private Function MakeRatio(ByVal cat As String, ByVal label As String, _
                            ByVal formulaTpl As String, ByVal numFmt As String) As Variant
    MakeRatio = Array(cat, label, formulaTpl, numFmt)
End Function


' --------- 在某 sheet 的 B 列里找指定指标名, 返回行号; 找不到返回 0 ---------
Private Function FindIndicatorRow(ByVal sheetName As String, ByVal indName As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then FindIndicatorRow = 0: Exit Function

    Dim found As Range
    Set found = ws.Range("B:B").Find(What:=indName, LookIn:=xlValues, _
                                       LookAt:=xlWhole, MatchCase:=True)
    If found Is Nothing Then
        FindIndicatorRow = 0
    Else
        FindIndicatorRow = found.Row
    End If
    On Error GoTo 0
End Function


' --------- 取某列的公司表头; 兼容 R1 合并单元格 ---------
Private Function HeaderTextAt(ByVal ws As Worksheet, ByVal col As Long) As String
    On Error Resume Next
    If ws.Cells(1, col).MergeCells Then
        HeaderTextAt = Trim$(CStr(ws.Cells(1, col).MergeArea.Cells(1, 1).Value))
    Else
        HeaderTextAt = Trim$(CStr(ws.Cells(1, col).Value))
    End If
    Err.Clear
    On Error GoTo 0
End Function


' --------- 统一报告期比较格式 ---------
Private Function PeriodKey(ByVal v As Variant) As String
    On Error Resume Next
    If IsDate(v) Then
        PeriodKey = Format$(CDate(v), "yyyy-mm-dd")
    Else
        PeriodKey = Trim$(CStr(v))
        If Len(PeriodKey) >= 10 And Mid$(PeriodKey, 5, 1) = "-" Then _
            PeriodKey = Left$(PeriodKey, 10)
    End If
    Err.Clear
    On Error GoTo 0
End Function


' --------- 在目标表中按 公司表头 + 报告期 找数据列 ---------
Private Function FindStatementColumn(ByVal sheetName As String, _
                                     ByVal companyHeader As String, _
                                     ByVal periodText As String) As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then GoTo CleanExit
    If Len(companyHeader) = 0 Or Len(periodText) = 0 Then GoTo CleanExit

    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 3 To lastCol
        If HeaderTextAt(ws, c) = companyHeader _
           And PeriodKey(ws.Cells(2, c).Value) = periodText Then
            FindStatementColumn = c
            GoTo CleanExit
        End If
    Next c

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


' --------- 生成跨 sheet 单元格引用 ---------
Private Function SheetCellRef(ByVal sheetName As String, ByVal rowNum As Long, ByVal colNum As Long) As String
    SheetCellRef = "'" & sheetName & "'!" & ColumnLetter(colNum) & rowNum
End Function


' --------- 列号 → 字母 (1 → A, 27 → AA, ...) ---------
Public Function ColumnLetter(ByVal col As Long) As String
    ColumnLetter = Split(ThisWorkbook.Sheets(1).Cells(1, col).Address, "$")(1)
End Function


' =================================================================
'  雪球 fallback (Phase 4b-5): EDGAR 404 时给 BS 用
'  数据源: stock.xueqiu.com/v5/stock/finance/us/balance.json
'  鉴权: 用户在 样本池!B5 提供 cookie (xq_a_token=... 或 完整 Cookie 头)
'  字段名: snake_case, 值通常为 [absolute_value, yoy_pct] 数组
'    多候选名 (匹配 us-gaap concept 的多个 xueqiu 别名), 取第一个非空
' =================================================================

Private Sub FetchUSFromXueqiu(ByVal strTicker As String, _
                              ByVal strKind As String, _
                              ByVal conceptMap As Variant, _
                              ByVal strQuarter As String, _
                              ByVal lngYear As Long, _
                              ByRef dictData As Object, _
                              ByRef dictPeriodSet As Object, _
                              ByRef dictIndicatorSet As Object, _
                              ByRef dictCategoryMap As Object, _
                              Optional ByVal collDiagRows As Collection = Nothing)
    ' --- stage tracker: 异常时附在 description 里, 帮助定位 ---
    Dim stage As String: stage = "init"
    On Error GoTo XqErr

    stage = "ReadCookie"
    Dim strCookie As String: strCookie = ReadXueqiuCookie()
    If Len(strCookie) = 0 Then
        Err.Raise vbObjectError + 542, "FetchUSFromXueqiu", _
            "雪球 fallback 需要 cookie. 请在 样本池!B5 填入 xq_a_token=... (浏览器登录 xueqiu.com 后 F12 拷)"
    End If

    stage = "BuildUrl"
    Dim strType As String: strType = XueqiuTypeForQuarter(strQuarter)
    Dim strEndpoint As String: strEndpoint = XueqiuEndpointForKind(strKind)
    Dim strUrl As String
    strUrl = "https://stock.xueqiu.com/v5/stock/finance/us/" & strEndpoint & ".json?" & _
             "symbol=" & strTicker & _
             "&type=" & strType & _
             "&is_detail=true&count=8"

    stage = "HttpGet"
    Dim strJson As String
    strJson = CachedXueqiuHttpGet(strUrl, strCookie, _
        "xueqiu_US_" & UCase$(strTicker) & "_" & strKind & "_" & _
        strType & "_" & CStr(lngYear))

    ' Debug: 把 raw JSON 写到 samples/xueqiu_<ticker>_bs.json 方便调字段映射
    stage = "DumpJson"
    On Error Resume Next
    DumpXueqiuJson strTicker, strKind, strJson
    On Error GoTo XqErr

    stage = "ParseJson"
    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(strJson)

    stage = "CheckErrorCode"
    If parsed.Exists("error_code") Then
        Dim errCode As Variant: errCode = parsed.Item("error_code")
        ' error_code 可能是 0 (Number) 或 "0" (String) 或 Null
        If Not IsNull(errCode) And Not IsEmpty(errCode) Then
            If CStr(errCode) <> "0" Then
                Err.Raise vbObjectError + 543, "FetchUSFromXueqiu", _
                    "雪球 API 错误 (" & errCode & "): " & NzStr(parsed, "error_description") & _
                    ". cookie 失效请更新 B5"
            End If
        End If
    End If

    stage = "GetData"
    If Not parsed.Exists("data") Then
        Err.Raise vbObjectError + 544, "FetchUSFromXueqiu", _
            "雪球响应缺 data 字段: " & strTicker
    End If
    Dim dataRoot As Object: Set dataRoot = parsed.Item("data")

    stage = "GetList"
    If Not dataRoot.Exists("list") Then
        Err.Raise vbObjectError + 545, "FetchUSFromXueqiu", _
            "雪球响应缺 list 字段: " & strTicker
    End If
    Dim listColl As Object: Set listColl = dataRoot.Item("list")

    stage = "CheckListEmpty"
    If listColl Is Nothing Then
        Err.Raise vbObjectError + 546, "FetchUSFromXueqiu", _
            "雪球 list 为 null (可能 cookie 失效或股票不收录): " & strTicker
    End If
    If listColl.Count = 0 Then
        Err.Raise vbObjectError + 547, "FetchUSFromXueqiu", _
            "雪球 list 为空 (该股票雪球未收录 " & strKind & " 数据): " & strTicker
    End If

    ' 公司子字典
    stage = "GetCompanyDict"
    Dim dictCompany As Object
    If dictData.Exists(strTicker) Then
        Set dictCompany = dictData.Item(strTicker)
    Else
        Set dictCompany = CreateObject("Scripting.Dictionary")
        dictData.Add strTicker, dictCompany
    End If

    ' 概念名 → 雪球字段候选 (基于 POM 真实 JSON 字段名校准)
    stage = "BuildXqMap"
    Dim mapXq As Object: Set mapXq = XueqiuFieldMapForKind(strKind)
    Dim dictMatchInfo As Object: Set dictMatchInfo = CreateObject("Scripting.Dictionary")

    ' 遍历 list 里的每期, 写到 dictData
    Dim recIdx As Long, periodEnd As String
    Dim ci As Long, mapEntry As Variant
    Dim strCat As String, strLabel As String
    Dim candidates As Variant, cand As Variant
    Dim candIdx As Long, totalCand As Long
    Dim val As Variant, valM As Double
    Dim dictPer As Object

    For recIdx = 1 To listColl.Count
        stage = "Record#" & recIdx & ":start"
        Dim record As Object: Set record = listColl.Item(recIdx)

        ' 优先用 ed 字段 ("YYYY-MM-DD" 字符串); 兜底用 report_date (unix ms)
        stage = "Record#" & recIdx & ":readEd"
        periodEnd = NzStr(record, "ed")
        If Len(periodEnd) < 10 Then periodEnd = ParseXueqiuReportDate(record)
        If Len(periodEnd) = 0 Then GoTo NextRecord

        ' 季度过滤 (用 ed 字符串后缀; 雪球美股财报只有 ed 没有 fp)
        stage = "Record#" & recIdx & ":matchPeriod(" & periodEnd & ")"
        If Not MatchXueqiuPeriod(record, periodEnd, strQuarter, lngYear) Then GoTo NextRecord

        ' 给本期建子字典
        stage = "Record#" & recIdx & ":dictPer"
        If dictCompany.Exists(periodEnd) Then
            Set dictPer = dictCompany.Item(periodEnd)
        Else
            Set dictPer = CreateObject("Scripting.Dictionary")
            dictCompany.Add periodEnd, dictPer
        End If

        ' 按 conceptMap 顺序找雪球字段, 第一个非空的填入
        For ci = LBound(conceptMap) To UBound(conceptMap)
            stage = "Record#" & recIdx & ":concept#" & ci
            mapEntry = conceptMap(ci)
            strCat = CStr(mapEntry(0))
            strLabel = CStr(mapEntry(1))
            If Not mapXq.Exists(strLabel) Then GoTo NextConcept

            candidates = Split(CStr(mapXq.Item(strLabel)), ",")
            candIdx = 0
            totalCand = UBound(candidates) - LBound(candidates) + 1
            For Each cand In candidates
                candIdx = candIdx + 1
                stage = "Record#" & recIdx & ":concept#" & ci & ":cand=" & cand
                val = XueqiuValue(record, Trim$(CStr(cand)))
                If Not IsEmpty(val) And Not IsNull(val) Then
                    ' 雪球美股财报金额通常是 USD 原值; EPS/比率由 conceptMap 的 scale 控制
                    stage = "Record#" & recIdx & ":concept#" & ci & ":cdbl(" & cand & ")"
                    valM = CDbl(val) / MapEntryScale(mapEntry)

                    stage = "Record#" & recIdx & ":concept#" & ci & ":write"
                    If Not dictPeriodSet.Exists(periodEnd) Then dictPeriodSet.Add periodEnd, True
                    If Not dictIndicatorSet.Exists(strLabel) Then
                        dictIndicatorSet.Add strLabel, dictIndicatorSet.Count
                        If Not dictCategoryMap.Exists(strLabel) Then _
                            dictCategoryMap.Add strLabel, strCat
                    End If
                    If dictPer.Exists(strLabel) Then
                        dictPer.Item(strLabel) = valM
                    Else
                        dictPer.Add strLabel, valM
                    End If

                    Dim scoreText As String, noteBase As String
                    If candIdx = 1 Then
                        scoreText = "100"
                        noteBase = "hardcoded_primary"
                    Else
                        scoreText = "85"
                        noteBase = "hardcoded_alt[" & candIdx & "/" & totalCand & "]"
                    End If

                    If dictMatchInfo.Exists(strLabel) Then
                        Dim info As Variant: info = dictMatchInfo.Item(strLabel)
                        info(7) = CLng(info(7)) + 1
                        If periodEnd > CStr(info(8)) Then info(8) = periodEnd
                        dictMatchInfo.Item(strLabel) = info
                    Else
                        dictMatchInfo.Add strLabel, Array("OK_XUEQIU", "Xueqiu", "xueqiu", _
                                                         Trim$(CStr(cand)), MapEntryUnit(mapEntry), _
                                                         scoreText, noteBase, 1, periodEnd)
                    End If
                    Exit For    ' 找到一个匹配的就够
                End If
            Next cand
NextConcept:
        Next ci

NextRecord:
    Next recIdx

    stage = "WriteDiagnostics"
    AppendDiagnosticsForConceptMap strTicker, strKind, conceptMap, dictMatchInfo, _
                                   Nothing, "xueqiu", collDiagRows

    Err.Clear
    Exit Sub

XqErr:
    ' 用自定义错误号向上重抛, 避免原生 380 等错误吞掉 stage 描述
    Dim origNum As Long: origNum = Err.Number
    Dim origSource As String: origSource = Err.Source
    Dim origDesc As String: origDesc = Err.Description
    Err.Clear
    Err.Raise vbObjectError + 590, "FetchUSFromXueqiu", _
        "[stage=" & stage & "] 原始错误号=" & origNum & _
        "; 原始来源=" & origSource & "; " & origDesc
End Sub


Private Sub FetchUSFallbackAfterXueqiu(ByVal strTicker As String, _
                                       ByVal strKind As String, _
                                       ByVal conceptMap As Variant, _
                                       ByVal strQuarter As String, _
                                       ByVal lngYear As Long, _
                                       ByRef dictData As Object, _
                                       ByRef dictPeriodSet As Object, _
                                       ByRef dictIndicatorSet As Object, _
                                       ByRef dictCategoryMap As Object, _
                                       Optional ByVal collDiagRows As Collection = Nothing, _
                                       Optional ByRef dictReportingCurrency As Object = Nothing, _
                                       Optional ByVal triggerReason As String = "")
    On Error Resume Next
    Err.Clear
    FetchUSFromXueqiu strTicker, strKind, conceptMap, strQuarter, lngYear, _
                      dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap, collDiagRows
    If Err.Number = 0 Then
        If Not dictReportingCurrency Is Nothing Then dictReportingCurrency(strTicker) = "USD"
        On Error GoTo 0
        Exit Sub
    End If

    Dim xqErrNum As Long: xqErrNum = Err.Number
    Dim xqErrSource As String: xqErrSource = Err.Source
    Dim xqErrDesc As String: xqErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If ReadStockAnalysisFallbackEnabled() = "开" And StockAnalysisUSSupportsKind(strKind) _
            And StockAnalysisUSSupportsTicker(strTicker) Then
        On Error Resume Next
        Err.Clear
        FetchUSFromStockAnalysis strTicker, strKind, conceptMap, strQuarter, lngYear, _
                                 dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap, _
                                 collDiagRows, dictReportingCurrency, triggerReason & "; xueqiu_failed=" & xqErrDesc
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If

        Dim saErrDesc As String: saErrDesc = Err.Description
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 592, "FetchUSFallbackAfterXueqiu", _
            "雪球 fallback 失败: " & xqErrDesc & "; stockanalysis fallback 失败: " & saErrDesc
    Else
        Err.Raise xqErrNum, xqErrSource, xqErrDesc
    End If
End Sub


Private Sub FetchUSFromStockAnalysis(ByVal strTicker As String, _
                                     ByVal strKind As String, _
                                     ByVal conceptMap As Variant, _
                                     ByVal strQuarter As String, _
                                     ByVal lngYear As Long, _
                                     ByRef dictData As Object, _
                                     ByRef dictPeriodSet As Object, _
                                     ByRef dictIndicatorSet As Object, _
                                     ByRef dictCategoryMap As Object, _
                                     Optional ByVal collDiagRows As Collection = Nothing, _
                                     Optional ByRef dictReportingCurrency As Object = Nothing, _
                                     Optional ByVal triggerReason As String = "")
    Dim stage As String: stage = "init"
    On Error GoTo SaErr

    stage = "BuildUrl"
    Dim strUrl As String: strUrl = StockAnalysisUSUrl(strTicker, strKind)
    Dim cacheKey As String
    cacheKey = "stockanalysis_US_" & UCase$(strTicker) & "_" & strKind & "_annual"

    stage = "HttpGet"
    Dim strHtml As String: strHtml = StockAnalysisUSHttpGet(strUrl, cacheKey)
    Dim reportingCurrency As String: reportingCurrency = StockAnalysisUSFinancialCurrency(strHtml)
    If Len(reportingCurrency) = 0 Then reportingCurrency = "USD"
    If reportingCurrency = "CNY" Then reportingCurrency = "RMB"

    stage = "ParseHtml"
    Dim objHtml As Object: Set objHtml = CreateObject("htmlfile")
    objHtml.Open
    objHtml.Write strHtml
    objHtml.Close

    Dim tables As Object: Set tables = objHtml.getElementsByTagName("table")
    If tables Is Nothing Or tables.Length = 0 Then
        Err.Raise vbObjectError + 593, "FetchUSFromStockAnalysis", _
            "stockanalysis 页面未找到 HTML table: " & strUrl
    End If

    Dim objTb As Object: Set objTb = tables.Item(0)
    If objTb.Rows.Length < 2 Then
        Err.Raise vbObjectError + 594, "FetchUSFromStockAnalysis", _
            "stockanalysis table 行数不足: " & strUrl
    End If

    Dim periods As Object: Set periods = CreateObject("Scripting.Dictionary")
    Dim headerYearRow As Object: Set headerYearRow = objTb.Rows.Item(0)
    Dim headerDateRow As Object: Set headerDateRow = headerYearRow
    Dim dataStartRow As Long: dataStartRow = 1
    If objTb.Rows.Length > 1 Then
        If InStr(1, StockAnalysisUSCellText(objTb.Rows.Item(1), 0), "Period Ending", vbTextCompare) > 0 Then
            Set headerDateRow = objTb.Rows.Item(1)
            dataStartRow = 2
        End If
    End If
    Dim j As Long, headerText As String, periodEnd As String
    For j = 1 To headerDateRow.Cells.Length - 1
        headerText = StockAnalysisUSCellText(headerYearRow, j) & " " & StockAnalysisUSCellText(headerDateRow, j)
        periodEnd = StockAnalysisUSPeriodFromHeader(headerText)
        If Len(periodEnd) > 0 Then
            If StockAnalysisUSMatchPeriod(periodEnd, strQuarter, lngYear, headerText) Then _
                periods.Item(CStr(j)) = periodEnd
        End If
    Next j
    If periods.Count = 0 Then
        Err.Raise vbObjectError + 595, "FetchUSFromStockAnalysis", _
            "stockanalysis 无匹配期间: " & strTicker & " / " & strQuarter & " / " & lngYear
    End If

    Dim dictRows As Object: Set dictRows = CreateObject("Scripting.Dictionary")
    dictRows.CompareMode = vbTextCompare
    Dim r As Long, rowLabel As String
    For r = dataStartRow To objTb.Rows.Length - 1
        rowLabel = StockAnalysisUSCellText(objTb.Rows.Item(r), 0)
        If Len(rowLabel) > 0 Then
            If Not dictRows.Exists(rowLabel) Then dictRows.Add rowLabel, objTb.Rows.Item(r)
        End If
    Next r

    Dim mapSA As Object: Set mapSA = StockAnalysisUSFieldMap(strKind)
    Dim dictMatchInfo As Object: Set dictMatchInfo = CreateObject("Scripting.Dictionary")
    dictMatchInfo.CompareMode = vbTextCompare

    Dim dictCompany As Object
    If dictData.Exists(strTicker) Then
        Set dictCompany = dictData.Item(strTicker)
    Else
        Set dictCompany = CreateObject("Scripting.Dictionary")
        dictData.Add strTicker, dictCompany
    End If

    Dim ci As Long, mapEntry As Variant, strCat As String, strLabel As String
    Dim candidates As Variant, cand As Variant, candIdx As Long, totalCand As Long
    Dim objRow As Object, key As Variant, rawValue As String, scaledValue As Double
    Dim dictPer As Object, matchPeriodEnd As String

    For ci = LBound(conceptMap) To UBound(conceptMap)
        mapEntry = conceptMap(ci)
        strCat = CStr(mapEntry(0))
        strLabel = CStr(mapEntry(1))
        If Not mapSA.Exists(strLabel) Then GoTo NextSaConcept

        candidates = mapSA.Item(strLabel)
        totalCand = UBound(candidates) - LBound(candidates) + 1
        candIdx = 0
        For Each cand In candidates
            candIdx = candIdx + 1
            If dictRows.Exists(CStr(cand)) Then
                Set objRow = dictRows.Item(CStr(cand))
                matchPeriodEnd = ""

                For Each key In periods.Keys
                    j = CLng(key)
                    If j < objRow.Cells.Length Then
                        rawValue = StockAnalysisUSCellText(objRow, j)
                        If StockAnalysisUSTryValue(rawValue, scaledValue) Then
                            periodEnd = CStr(periods.Item(key))
                            If periodEnd > matchPeriodEnd Then matchPeriodEnd = periodEnd
                            If Not dictPeriodSet.Exists(periodEnd) Then dictPeriodSet.Add periodEnd, True
                            If Not dictIndicatorSet.Exists(strLabel) Then
                                dictIndicatorSet.Add strLabel, dictIndicatorSet.Count
                                If Not dictCategoryMap.Exists(strLabel) Then dictCategoryMap.Add strLabel, strCat
                            End If
                            If dictCompany.Exists(periodEnd) Then
                                Set dictPer = dictCompany.Item(periodEnd)
                            Else
                                Set dictPer = CreateObject("Scripting.Dictionary")
                                dictCompany.Add periodEnd, dictPer
                            End If
                            dictPer(strLabel) = scaledValue
                        End If
                    End If
                Next key

                If Len(matchPeriodEnd) > 0 Then
                    dictMatchInfo(strLabel) = Array(CStr(cand), candIdx, totalCand, matchPeriodEnd)
                    Exit For
                End If
            End If
        Next cand
NextSaConcept:
    Next ci

    If Not HasAnyCoreLabel(dictData, strTicker, CoreLabelsForKind(strKind)) Then
        Err.Raise vbObjectError + 596, "FetchUSFromStockAnalysis", _
            "stockanalysis 未匹配核心字段: " & strTicker & " / " & strKind
    End If

    AppendStockAnalysisUSDiagnostics strTicker, strKind, conceptMap, dictMatchInfo, _
                                     reportingCurrency, triggerReason, collDiagRows
    If Not dictReportingCurrency Is Nothing Then dictReportingCurrency(strTicker) = reportingCurrency
    Application.Wait Now + TimeSerial(0, 0, 2)
    Exit Sub

SaErr:
    Dim origNum As Long: origNum = Err.Number
    Dim origSource As String: origSource = Err.Source
    Dim origDesc As String: origDesc = Err.Description
    Err.Clear
    Err.Raise vbObjectError + 597, "FetchUSFromStockAnalysis", _
        "[stage=" & stage & "] 原始错误号=" & origNum & _
        "; 原始来源=" & origSource & "; " & origDesc
End Sub


Private Sub AppendStockAnalysisUSDiagnostics(ByVal strTicker As String, _
                                             ByVal strKind As String, _
                                             ByVal conceptMap As Variant, _
                                             ByVal dictMatchInfo As Object, _
                                             ByVal reportingCurrency As String, _
                                             ByVal triggerReason As String, _
                                             ByVal collDiagRows As Collection)
    If collDiagRows Is Nothing Then Exit Sub
    Dim i As Long, strLabel As String, info As Variant, noteText As String
    For i = LBound(conceptMap) To UBound(conceptMap)
        strLabel = CStr(conceptMap(i)(1))
        If dictMatchInfo.Exists(strLabel) Then
            info = dictMatchInfo.Item(strLabel)
            If CLng(info(1)) = 1 Then
                noteText = "stockanalysis_primary"
            Else
                noteText = "stockanalysis_alt[" & CStr(info(1)) & "/" & CStr(info(2)) & "]"
            End If
            If Len(triggerReason) > 0 Then noteText = noteText & "; trigger=" & triggerReason
            AddDiagnosticRow collDiagRows, strTicker, strKind, strLabel, "OK_STOCKANALYSIS", _
                             "stockanalysis (fallback)", "stockanalysis", CStr(info(0)), _
                             reportingCurrency, "90", noteText, _
                             FxRateTextForDiagnostic(reportingCurrency, CStr(info(3)), strKind)
        Else
            AddDiagnosticRow collDiagRows, strTicker, strKind, strLabel, "MISSING", _
                             "stockanalysis (fallback)", "stockanalysis", "—", _
                             reportingCurrency, "—", "stockanalysis field not mapped"
        End If
    Next i
End Sub


Private Function StockAnalysisUSSupportsKind(ByVal strKind As String) As Boolean
    Select Case strKind
        Case "BalanceSheet", "Income", "CashFlow"
            StockAnalysisUSSupportsKind = True
        Case Else
            StockAnalysisUSSupportsKind = False
    End Select
End Function


Private Function StockAnalysisUSSupportsTicker(ByVal strTicker As String) As Boolean
    Select Case UCase$(Trim$(strTicker))
        Case "BABA", "JD", "PDD"
            StockAnalysisUSSupportsTicker = True
        Case Else
            StockAnalysisUSSupportsTicker = False
    End Select
End Function


Private Function StockAnalysisUSUrl(ByVal strTicker As String, ByVal strKind As String) As String
    Dim pathPart As String
    Select Case strKind
        Case "BalanceSheet": pathPart = "financials/balance-sheet/"
        Case "Income":       pathPart = "financials/"
        Case "CashFlow":     pathPart = "financials/cash-flow-statement/"
        Case Else
            Err.Raise vbObjectError + 598, "StockAnalysisUSUrl", _
                "stockanalysis US 不支持报表类型: " & strKind
    End Select
    StockAnalysisUSUrl = "https://stockanalysis.com/stocks/" & LCase$(Trim$(strTicker)) & "/" & pathPart
End Function


Private Function StockAnalysisUSHttpGet(ByVal strUrl As String, ByVal cacheKey As String) As String
    Dim cached As String: cached = ReadLocalHttpCache(cacheKey)
    If Len(cached) > 0 Then
        StockAnalysisUSHttpGet = cached
        Exit Function
    End If

    Dim objWinHttp As Object, arrByte() As Byte
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With objWinHttp
        .SetTimeouts 30000, 60000, 60000, 60000
        .Open "GET", strUrl, False
        .SetRequestHeader "User-Agent", _
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " & _
            "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        .SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        .SetRequestHeader "Accept-Language", "en-US,en;q=0.9,zh-CN;q=0.8"
        .SetRequestHeader "Accept-Encoding", "identity"
        .Send
        .WaitForResponse 60
        If .Status < 200 Or .Status >= 300 Then
            Err.Raise vbObjectError + 599, "StockAnalysisUSHttpGet", _
                "HTTP " & .Status & " for " & strUrl
        End If
        arrByte = .ResponseBody
    End With
    StockAnalysisUSHttpGet = USByteToStr(arrByte, "utf-8")
    WriteLocalHttpCache cacheKey, StockAnalysisUSHttpGet
End Function


Private Function StockAnalysisUSFieldMap(ByVal strKind As String) As Object
    Dim mapSA As Object: Set mapSA = CreateObject("Scripting.Dictionary")
    mapSA.CompareMode = vbTextCompare

    Select Case strKind
        Case "BalanceSheet"
            mapSA.Add "Cash & equivalents", Array("Cash & Equivalents")
            mapSA.Add "Marketable securities (current)", Array("Short-Term Investments")
            mapSA.Add "Accounts receivable, net", Array("Accounts Receivable", "Total Trade Receivables", "Other Receivables")
            mapSA.Add "Inventory", Array("Inventory")
            mapSA.Add "Other current assets", Array("Other Current Assets")
            mapSA.Add "Total current assets", Array("Total Current Assets")
            mapSA.Add "Property, plant & equipment, net", Array("Net Property, Plant & Equipment")
            mapSA.Add "Goodwill", Array("Goodwill")
            mapSA.Add "Intangible assets", Array("Other Intangible Assets")
            mapSA.Add "Other non-current assets", Array("Other Long-Term Assets")
            mapSA.Add "Total assets", Array("Total Assets")
            mapSA.Add "Accounts payable", Array("Accounts Payable")
            mapSA.Add "Short-term debt", Array("Short-Term Debt", "Current Portion of Long-Term Debt")
            mapSA.Add "Other current liabilities", Array("Other Current Liabilities", "Accrued Expenses")
            mapSA.Add "Total current liabilities", Array("Total Current Liabilities")
            mapSA.Add "Long-term debt", Array("Long-Term Debt")
            mapSA.Add "Other non-current liabilities", Array("Other Long-Term Liabilities")
            mapSA.Add "Total non-current liabilities", Array("Total Long-Term Liabilities")
            mapSA.Add "Total liabilities", Array("Total Liabilities")
            mapSA.Add "Common stock", Array("Common Stock", "Additional Paid-in Capital")
            mapSA.Add "Retained earnings", Array("Retained Earnings")
            mapSA.Add "Accumulated OCI", Array("Accumulated Other Comprehensive Income")
            mapSA.Add "Total stockholders' equity", Array("Total Common Shareholders' Equity", "Shareholders' Equity")
            mapSA.Add "Total liabilities & equity", Array("Total Liabilities & Equity")
        Case "Income"
            mapSA.Add "Revenue", Array("Revenue")
            mapSA.Add "Cost of goods & services sold", Array("Cost of Revenue")
            mapSA.Add "Gross profit", Array("Gross Profit")
            mapSA.Add "R&D expense", Array("Research & Development")
            mapSA.Add "SG&A expense", Array("Selling, General & Admin")
            mapSA.Add "Total operating expenses", Array("Total Operating Expenses")
            mapSA.Add "Operating income", Array("Operating Income")
            mapSA.Add "Non-operating income / (expense)", Array("Other Non-Operating Income (Expense)", "Total Non-Operating Income (Expense)")
            mapSA.Add "Interest expense", Array("Interest Expense")
            mapSA.Add "Pre-tax income", Array("Pretax Income")
            mapSA.Add "Income tax expense", Array("Provision for Income Taxes")
            mapSA.Add "Net income", Array("Net Income", "Net Income to Common")
            mapSA.Add "Basic EPS (USD/share)", Array("EPS (Basic)")
            mapSA.Add "Diluted EPS (USD/share)", Array("EPS (Diluted)")
        Case "CashFlow"
            mapSA.Add "Net income", Array("Net Income")
            mapSA.Add "Depreciation & amortization", Array("Depreciation & Amortization")
            mapSA.Add "Stock-based compensation", Array("Stock-Based Compensation")
            mapSA.Add "Other non-cash items", Array("Other Adjustments")
            mapSA.Add "Change in AR", Array("Change in Receivables")
            mapSA.Add "Change in inventory", Array("Changes in Inventories")
            mapSA.Add "Change in AP", Array("Changes in Accounts Payable")
            mapSA.Add "Change in deferred revenue", Array("Changes in Unearned Revenue")
            mapSA.Add "Change in other operating liabilities", Array("Changes in Other Operating Activities")
            mapSA.Add "Cash from operations", Array("Operating Cash Flow")
            mapSA.Add "Capex", Array("Capital Expenditures")
            mapSA.Add "Business acquisitions", Array("Payments for Business Acquisitions")
            mapSA.Add "Proceeds from sale of PP&E", Array("Sale of Property, Plant & Equipment")
            mapSA.Add "Other investing", Array("Other Investing Activities")
            mapSA.Add "Cash from investing", Array("Investing Cash Flow")
            mapSA.Add "Dividends paid", Array("Common Dividends Paid")
            mapSA.Add "Stock repurchases", Array("Repurchase of Common Stock")
            mapSA.Add "Stock issuance", Array("Issuance of Common Stock")
            mapSA.Add "Long-term debt issued", Array("Long-Term Debt Issued")
            mapSA.Add "Long-term debt repaid", Array("Long-Term Debt Repaid")
            mapSA.Add "Other financing", Array("Other Financing Activities")
            mapSA.Add "Cash from financing", Array("Financing Cash Flow")
            mapSA.Add "FX effect on cash", Array("Effect of Exchange Rate Changes on Cash and Cash Equivalents")
            mapSA.Add "Net change in cash (incl FX)", Array("Net Cash Flow")
            mapSA.Add "Cash at end of period", Array("Cash & Equivalents")
        Case Else
            Err.Raise vbObjectError + 600, "StockAnalysisUSFieldMap", _
                "stockanalysis US 不支持报表类型: " & strKind
    End Select

    Set StockAnalysisUSFieldMap = mapSA
End Function


Private Function StockAnalysisUSFinancialCurrency(ByVal strHtml As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "financial:""([A-Z]{3})"""
    re.IgnoreCase = True
    If re.Test(strHtml) Then
        StockAnalysisUSFinancialCurrency = UCase$(re.Execute(strHtml)(0).SubMatches(0))
    Else
        StockAnalysisUSFinancialCurrency = "USD"
    End If
End Function


Private Function StockAnalysisUSPeriodFromHeader(ByVal headerText As String) As String
    Dim s As String: s = UCase$(StockAnalysisUSCleanText(headerText))
    If Len(s) = 0 Or Left$(s, 3) = "TTM" Then Exit Function

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2}),\s+(\d{4})"
    re.Global = True
    If Not re.Test(s) Then Exit Function

    Dim matches As Object: Set matches = re.Execute(s)
    Dim m As Object: Set m = matches.Item(matches.Count - 1)
    Dim monthNum As Long: monthNum = StockAnalysisUSMonthNumber(CStr(m.SubMatches(0)))
    If monthNum = 0 Then Exit Function
    StockAnalysisUSPeriodFromHeader = Format$(DateSerial(CLng(m.SubMatches(2)), monthNum, CLng(m.SubMatches(1))), "yyyy-mm-dd")
End Function


Private Function StockAnalysisUSMatchPeriod(ByVal periodEnd As String, _
                                            ByVal strQuarter As String, _
                                            ByVal lngYear As Long, _
                                            ByVal headerText As String) As Boolean
    If Len(periodEnd) < 10 Then Exit Function
    If lngYear > 0 Then
        If CLng(Left$(periodEnd, 4)) <> lngYear Then Exit Function
    End If

    Select Case UCase$(Trim$(strQuarter))
        Case "全部", ""
            StockAnalysisUSMatchPeriod = True
        Case "Q1"
            StockAnalysisUSMatchPeriod = (Right$(periodEnd, 5) = "03-31")
        Case "Q2"
            StockAnalysisUSMatchPeriod = (Right$(periodEnd, 5) = "06-30")
        Case "Q3"
            StockAnalysisUSMatchPeriod = (Right$(periodEnd, 5) = "09-30")
        Case "Q4"
            StockAnalysisUSMatchPeriod = (Left$(UCase$(Trim$(headerText)), 2) = "FY")
        Case Else
            StockAnalysisUSMatchPeriod = True
    End Select
End Function


Private Function StockAnalysisUSCellText(ByVal objRow As Object, ByVal colIndex As Long) As String
    On Error Resume Next
    StockAnalysisUSCellText = StockAnalysisUSCleanText(CStr(objRow.Cells.Item(colIndex).innerText))
    Err.Clear
    On Error GoTo 0
End Function


Private Function StockAnalysisUSCleanText(ByVal rawText As String) As String
    Dim s As String: s = CStr(rawText)
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    StockAnalysisUSCleanText = Trim$(s)
End Function


Private Function StockAnalysisUSTryValue(ByVal rawValue As String, ByRef scaledValue As Double) As Boolean
    Dim s As String: s = StockAnalysisUSCleanText(rawValue)
    If Len(s) = 0 Or s = "-" Or s = "--" Then Exit Function
    s = Replace(s, ",", "")
    s = Replace(s, "−", "-")
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then s = "-" & Mid$(s, 2, Len(s) - 2)

    Dim isPct As Boolean
    If Right$(s, 1) = "%" Then
        isPct = True
        s = Left$(s, Len(s) - 1)
    End If
    If Not IsNumeric(s) Then Exit Function

    scaledValue = CDbl(s)
    If isPct Then scaledValue = scaledValue / 100#
    StockAnalysisUSTryValue = True
End Function


Private Function StockAnalysisUSMonthNumber(ByVal mon As String) As Long
    Select Case UCase$(mon)
        Case "JAN": StockAnalysisUSMonthNumber = 1
        Case "FEB": StockAnalysisUSMonthNumber = 2
        Case "MAR": StockAnalysisUSMonthNumber = 3
        Case "APR": StockAnalysisUSMonthNumber = 4
        Case "MAY": StockAnalysisUSMonthNumber = 5
        Case "JUN": StockAnalysisUSMonthNumber = 6
        Case "JUL": StockAnalysisUSMonthNumber = 7
        Case "AUG": StockAnalysisUSMonthNumber = 8
        Case "SEP": StockAnalysisUSMonthNumber = 9
        Case "OCT": StockAnalysisUSMonthNumber = 10
        Case "NOV": StockAnalysisUSMonthNumber = 11
        Case "DEC": StockAnalysisUSMonthNumber = 12
    End Select
End Function


Private Function USByteToStr(arrByte, ByVal strCharSet As String) As String
    With CreateObject("Adodb.Stream")
        .Type = 1
        .Open
        .Write arrByte
        .Position = 0
        .Type = 2
        .Charset = strCharSet
        USByteToStr = .ReadText
        .Close
    End With
End Function


' --------- 雪球 fallback 支持的美股报表类型 ---------
Private Function XueqiuSupportsKind(ByVal strKind As String) As Boolean
    Select Case strKind
        Case "BalanceSheet", "Income", "CashFlow", "Indicator"
            XueqiuSupportsKind = True
        Case Else
            XueqiuSupportsKind = False
    End Select
End Function


' --------- 美股报表类型 → 雪球 us finance endpoint ---------
Private Function XueqiuEndpointForKind(ByVal strKind As String) As String
    Select Case strKind
        Case "BalanceSheet": XueqiuEndpointForKind = "balance"
        Case "Income":       XueqiuEndpointForKind = "income"
        Case "CashFlow":     XueqiuEndpointForKind = "cash_flow"
        Case "Indicator":    XueqiuEndpointForKind = "indicator"
        Case Else
            Err.Raise vbObjectError + 548, "XueqiuEndpointForKind", _
                "雪球 fallback 不支持报表类型: " & strKind
    End Select
End Function


' --------- 英文指标名 → 雪球字段候选 ---------
Private Function XueqiuFieldMapForKind(ByVal strKind As String) As Object
    Dim mapXq As Object: Set mapXq = CreateObject("Scripting.Dictionary")

    Select Case strKind
        Case "BalanceSheet"
            mapXq.Add "Cash & equivalents", "cce,total_cash"
            mapXq.Add "Marketable securities (current)", "st_invest"
            mapXq.Add "Accounts receivable, net", "net_receivables"
            mapXq.Add "Inventory", "inventory"
            mapXq.Add "Other current assets", "current_assets_special_subject,prepaid_expense,dt_assets_current_assets"
            mapXq.Add "Total current assets", "total_current_assets"
            mapXq.Add "Marketable securities (non-current)", "lt_invest,equity_and_othr_invest"
            mapXq.Add "Property, plant & equipment, net", "net_property_plant_and_equip"
            mapXq.Add "Goodwill", "goodwill"
            mapXq.Add "Intangible assets", "net_intangible_assets"
            mapXq.Add "Other non-current assets", "nca_si,dt_assets_noncurrent_assets"
            mapXq.Add "Total non-current assets", "total_noncurrent_assets"
            mapXq.Add "Total assets", "total_assets"
            mapXq.Add "Accounts payable", "accounts_payable"
            mapXq.Add "Short-term debt", "st_debt"
            mapXq.Add "Compensation & benefits", "accrued_liab"
            mapXq.Add "Other current liabilities", "current_liab_si"
            mapXq.Add "Total current liabilities", "total_current_liab"
            mapXq.Add "Long-term debt", "lt_debt"
            mapXq.Add "Other non-current liabilities", "noncurrent_liab_si,dr_noncurrent_liab,deferred_tax_liab"
            mapXq.Add "Total non-current liabilities", "total_noncurrent_liab"
            mapXq.Add "Total liabilities", "total_liab"
            mapXq.Add "Common stock", "common_stock"
            mapXq.Add "Retained earnings", "retained_earning"
            mapXq.Add "Accumulated OCI", "accum_othr_compre_income"
            mapXq.Add "Total stockholders' equity", "total_holders_equity,total_equity"
            mapXq.Add "Total liabilities & equity", "total_assets"     ' 会计恒等式 (雪球未单独提供)

        Case "Income"
            mapXq.Add "Revenue", "revenue,total_revenue"
            mapXq.Add "Cost of goods & services sold", "sales_cost"
            mapXq.Add "Gross profit", "gross_profit"
            mapXq.Add "R&D expense", "rad_expenses"
            mapXq.Add "SG&A expense", "marketing_selling_etc"
            mapXq.Add "Total operating expenses", "total_operate_expenses,total_operate_expenses_si"
            mapXq.Add "Operating income", "operating_income"
            mapXq.Add "Non-operating income / (expense)", "income_from_co_before_tax_si"
            mapXq.Add "Interest expense", "interest_expense,net_interest_expense"
            mapXq.Add "Pre-tax income", "income_from_co_before_it"
            mapXq.Add "Income tax expense", "income_tax"
            mapXq.Add "Net income", "net_income_atcss,total_net_income_atcss,net_income"
            mapXq.Add "Basic EPS (USD/share)", "total_basic_earning_common_ps,basic_eps"
            mapXq.Add "Diluted EPS (USD/share)", "total_dlt_earnings_common_ps,eps_dlt"

        Case "CashFlow"
            mapXq.Add "Cash from operations", "net_cash_provided_by_oa"
            mapXq.Add "Depreciation & amortization", "depreciation_and_amortization"
            mapXq.Add "Purchases of investments", "purs_of_invest"
            mapXq.Add "Cash from investing", "net_cash_used_in_ia"
            mapXq.Add "Capex", "payment_for_property_and_equip"
            mapXq.Add "Cash from financing", "net_cash_used_in_fa"
            mapXq.Add "Dividends paid", "dividend_paid"
            mapXq.Add "Stock repurchases", "repur_of_common_stock"
            mapXq.Add "Stock issuance", "common_stock_issue"
            mapXq.Add "FX effect on cash", "effect_of_exchange_chg_on_cce"
            mapXq.Add "Net change in cash (incl FX)", "increase_in_cce"
            mapXq.Add "Cash at beginning of period", "cce_at_boy"
            mapXq.Add "Cash at end of period", "cce_at_eoy"

        Case "Indicator"
            mapXq.Add "Basic EPS (USD/share)", "basic_eps"
            mapXq.Add "Diluted EPS (USD/share)", "eps_dlt"
            mapXq.Add "Dividends declared per share (USD)", "dividend_ps"

        Case Else
            Err.Raise vbObjectError + 549, "XueqiuFieldMapForKind", _
                "雪球 fallback 不支持报表类型: " & strKind
    End Select

    Set XueqiuFieldMapForKind = mapXq
End Function


' --------- 雪球 record 里的 report_date (unix ms) → "YYYY-MM-DD" 字符串 ---------
Private Function ParseXueqiuReportDate(ByVal record As Object) As String
    On Error Resume Next
    ParseXueqiuReportDate = ""
    If Not record.Exists("report_date") Then GoTo CleanExit
    Dim v As Variant: v = record.Item("report_date")
    If IsNull(v) Or IsEmpty(v) Then GoTo CleanExit

    ' Unix ms → VBA Date
    Dim dt As Date
    dt = DateAdd("s", CDbl(v) / 1000#, DateSerial(1970, 1, 1))
    ParseXueqiuReportDate = Format(dt, "yyyy-mm-dd")
CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


' --------- 取雪球 record 某字段值 (处理 [absolute, yoy_pct] 数组形式) ---------
Private Function XueqiuValue(ByVal record As Object, ByVal key As String) As Variant
    On Error Resume Next
    XueqiuValue = Empty
    If Not record.Exists(key) Then GoTo CleanExit

    ' JSON array 在 VBA-JSON 里是 Collection, 需要用 Set 接对象;
    ' 直接塞进 Variant 会在部分 Office/VBA 版本触发 450 并被 Resume Next 吞掉。
    If IsObject(record.Item(key)) Then
        Dim objVal As Object: Set objVal = record.Item(key)
        If TypeName(objVal) = "Collection" Then
            If objVal.Count >= 1 Then
                Dim inner As Variant: inner = objVal.Item(1)
                If Not IsNull(inner) And Not IsEmpty(inner) Then XueqiuValue = inner
            End If
        End If
    Else
        Dim v As Variant: v = record.Item(key)
        If IsNull(v) Or IsEmpty(v) Then GoTo CleanExit
        XueqiuValue = v
    End If

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


' --------- 把 raw 雪球 JSON 写到 samples/xueqiu_{ticker}_{kind}.json (调试用) ---------
Private Sub DumpXueqiuJson(ByVal strTicker As String, ByVal strKind As String, ByVal strJson As String)
    On Error Resume Next
    Dim wbPath As String: wbPath = ThisWorkbook.Path
    Dim sampleDir As String: sampleDir = wbPath & Application.PathSeparator & "samples"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sampleDir) Then fso.CreateFolder sampleDir

    Dim fname As String
    fname = sampleDir & Application.PathSeparator & "xueqiu_" & strTicker & "_" & _
            XueqiuDumpSuffixForKind(strKind) & ".json"
    Dim ts As Object: Set ts = fso.CreateTextFile(fname, True, True)    ' Unicode
    ts.Write strJson
    ts.Close
    Err.Clear
    On Error GoTo 0
End Sub


Private Function XueqiuDumpSuffixForKind(ByVal strKind As String) As String
    Select Case strKind
        Case "BalanceSheet": XueqiuDumpSuffixForKind = "bs"
        Case "Income":       XueqiuDumpSuffixForKind = "income"
        Case "CashFlow":     XueqiuDumpSuffixForKind = "cash_flow"
        Case "Indicator":    XueqiuDumpSuffixForKind = "indicator"
        Case Else:           XueqiuDumpSuffixForKind = LCase$(strKind)
    End Select
End Function


' --------- 雪球 record 按 季度 + 年份 过滤 (按 ed 字符串日期匹配) ---------
'   全部 + 0年份 → 全部留
'   Q1/Q2/Q3 → 按月末后缀匹配; Q4 → 按雪球 FY 年报标记匹配
'   lngYear>0 → 优先按 report_annual (= fiscal year) 匹配, 再兜底 ed 年份
Private Function MatchXueqiuPeriod(ByVal record As Object, _
                                    ByVal periodEnd As String, _
                                    ByVal strQuarter As String, _
                                    ByVal lngYear As Long) As Boolean
    MatchXueqiuPeriod = False
    If Len(periodEnd) < 10 Then Exit Function

    ' 年份检查: 优先用雪球 report_annual (= fiscal year), 再兜底用 ed 年份。
    ' BABA 等 3 月财年公司 Q4/FY 的 ed=2024-03-31, report_annual=2024。
    If lngYear > 0 Then
        Dim yr As Long: yr = 0
        On Error Resume Next
        If Not record Is Nothing Then
            If record.Exists("report_annual") Then yr = CLng(record.Item("report_annual"))
        End If
        Err.Clear
        On Error GoTo 0
        If yr = 0 Then yr = CLng(Left$(periodEnd, 4))
        If yr <> lngYear Then Exit Function
    End If

    ' 季度检查。Q4 对应雪球 FY 年报, 不能写死 -12-31。
    Dim suffix As String
    Select Case strQuarter
        Case "全部", "":   MatchXueqiuPeriod = True: Exit Function
        Case "Q1":         suffix = "-03-31"
        Case "Q2":         suffix = "-06-30"
        Case "Q3":         suffix = "-09-30"
        Case "Q4"
            Dim reportType As String: reportType = NzStr(record, "report_type_code")
            Dim reportName As String: reportName = NzStr(record, "report_name")
            If reportType = "596001" Or InStr(1, reportName, "FY", vbTextCompare) > 0 Then _
                MatchXueqiuPeriod = True
            Exit Function
        Case Else:         MatchXueqiuPeriod = True: Exit Function
    End Select

    If Right$(periodEnd, Len(suffix)) = suffix Then MatchXueqiuPeriod = True
End Function


' --------- 把 A4 季度选择 映射成雪球 type 参数 ---------
'   全部=all, Q1/Q2/Q3=对应季报, Q4=年报 (Q4)
Private Function XueqiuTypeForQuarter(ByVal strQuarter As String) As String
    Select Case strQuarter
        Case "全部", "":   XueqiuTypeForQuarter = "all"
        Case "Q1":         XueqiuTypeForQuarter = "Q1"
        Case "Q2":         XueqiuTypeForQuarter = "Q2"
        Case "Q3":         XueqiuTypeForQuarter = "Q3"
        Case "Q4":         XueqiuTypeForQuarter = "Q4"
        Case Else:         XueqiuTypeForQuarter = "all"
    End Select
End Function

Attribute VB_Name = "模块_抓韩股财报"
Option Explicit

' =================================================================
'  抓韩股财报 — 共享 helpers (RunKRStatement / FetchKRFromStockAnalysis)
'
'  数据源: stockanalysis.com KRX HTML tables
'    https://stockanalysis.com/quote/krx/{ticker}/financials/
'    https://stockanalysis.com/quote/krx/{ticker}/financials/balance-sheet/
'    https://stockanalysis.com/quote/krx/{ticker}/financials/cash-flow-statement/
'
'  约定:
'    - 韩股只走 stockanalysis.com, 不走雪球 / EDGAR / fuzzy。
'    - HTML 解析使用 htmlfile DOM, 不用正则解析表格。
'    - stockanalysis 表格金额单位为 millions KRW; 写入值 = 表格值 / 1,000 (= KRW billions)。
' =================================================================


' --------- 通用韩股单表抓数主流程 ---------
'   strKind     : "BalanceSheet" / "Income" / "CashFlow"
'   targetSheet : 目标 sheet 名
'   conceptMap  : 1-d Variant 数组, 每元素至少是 Array(大类, 标签)
'   maxPeriods  : 最新期数截断
Public Sub RunKRStatement(ByVal strKind As String, ByVal targetSheet As String, _
                          ByVal conceptMap As Variant, _
                          ByVal maxPeriods As Long)
    Dim wsPool As Worksheet, wsTarget As Worksheet
    Dim arrPool As Variant
    Dim i As Long, lngRow As Long, numCompanies As Long
    Dim intFailCnt As Long, strErrLog As String
    Dim strCodeRaw As String, strCode As String, strName As String
    Dim dtTime As Double: dtTime = Timer

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictPeriodSet As Object: Set dictPeriodSet = CreateObject("Scripting.Dictionary")
    Dim dictIndicatorSet As Object: Set dictIndicatorSet = CreateObject("Scripting.Dictionary")
    Dim dictCategoryMap As Object: Set dictCategoryMap = CreateObject("Scripting.Dictionary")
    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    Dim collCodes As New Collection
    Dim collDiagRows As New Collection

    g_diagnosticSheetName = "韩股_抓取诊断"

    Set wsPool = ThisWorkbook.Sheets("样本池")
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)

    On Error GoTo CleanUp
    Application.ScreenUpdating = False

    lngRow = wsPool.Cells(wsPool.Rows.Count, POOL_KR_CODE_COL).End(xlUp).Row
    If lngRow < POOL_DATA_START_ROW Then
        intFailCnt = 1
        strErrLog = "样本池韩股区为空"
        GoTo CleanUp
    End If

    arrPool = wsPool.Range(wsPool.Cells(POOL_DATA_START_ROW, POOL_KR_CODE_COL), _
                           wsPool.Cells(lngRow, POOL_KR_NAME_COL)).Value
    If Not IsArray(arrPool) Then
        Dim singleVal As Variant: singleVal = arrPool
        ReDim arrPool(1 To 1, 1 To 2)
        arrPool(1, 1) = singleVal
    End If
    numCompanies = UBound(arrPool, 1)

    Dim strQuarter As String: strQuarter = ReadQuarterSelection()
    Dim lngYear As Long: lngYear = ReadYearSelection()

    For i = 1 To numCompanies
        strCodeRaw = Trim$(CStr(arrPool(i, 1)))
        If Len(strCodeRaw) = 0 Then GoTo NextRow
        strName = Trim$(CStr(arrPool(i, 2)))
        strCode = NormalizeKRTicker(strCodeRaw)

        Application.StatusBar = "抓取中: " & targetSheet & " (" & i & "/" & numCompanies & ") " & strCode
        DoEvents

        Dim tempData As Object: Set tempData = CreateObject("Scripting.Dictionary")
        Dim tempPeriodSet As Object: Set tempPeriodSet = CreateObject("Scripting.Dictionary")
        Dim tempIndicatorSet As Object: Set tempIndicatorSet = CreateObject("Scripting.Dictionary")
        Dim tempCategoryMap As Object: Set tempCategoryMap = CreateObject("Scripting.Dictionary")

        On Error Resume Next
        Err.Clear
        FetchKRFromStockAnalysis strCode, strKind, conceptMap, strQuarter, lngYear, maxPeriods, _
                                 tempData, tempPeriodSet, tempIndicatorSet, tempCategoryMap, collDiagRows
        Dim fetchErrNum As Long: fetchErrNum = Err.Number
        Dim fetchErrSource As String: fetchErrSource = Err.Source
        Dim fetchErrDesc As String: fetchErrDesc = Err.Description
        Err.Clear
        On Error GoTo CleanUp

        If fetchErrNum <> 0 Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": " & _
                        "错误号=" & fetchErrNum & "; 来源=" & fetchErrSource & "; " & fetchErrDesc
            AddMissingDiagnosticsForCompany strCode, strKind, conceptMap, collDiagRows, _
                                            "fetch_failed: " & fetchErrDesc
        ElseIf Not HasAnyCoreLabelKR(tempData, strCode, CoreLabelsForKind(strKind)) Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": 核心字段未命中"
        Else
            CommitTempKRData strCode, tempData, tempPeriodSet, tempIndicatorSet, tempCategoryMap, _
                             dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap
            If Not dictCompanyName.Exists(strCode) Then dictCompanyName.Add strCode, strName
            collCodes.Add strCode
        End If

        Application.Wait Now + TimeSerial(0, 0, 1)

NextRow:
    Next i

    If collCodes.Count = 0 Then
        ClearKRWideTableOutput wsTarget
        GoTo CleanUp
    End If

    If dictPeriodSet.Count = 0 Or dictIndicatorSet.Count = 0 Then
        intFailCnt = intFailCnt + 1
        strErrLog = strErrLog & vbCrLf & "[" & targetSheet & "] 无匹配数据" & _
                    "(期数=" & dictPeriodSet.Count & " / 指标=" & dictIndicatorSet.Count & ")"
        GoTo CleanUp
    End If

    Dim arrCodes() As String
    ReDim arrCodes(1 To collCodes.Count)
    For i = 1 To collCodes.Count
        arrCodes(i) = collCodes.Item(i)
    Next i

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

    arrIndicators = OrderIndicatorsByKRConceptMap(dictIndicatorSet, conceptMap)

    Application.StatusBar = "写入: " & targetSheet
    DoEvents
    Dim dictReportingCurrency As Object: Set dictReportingCurrency = CreateObject("Scripting.Dictionary")
    Dim krCode As Variant
    For Each krCode In arrCodes
        dictReportingCurrency(CStr(krCode)) = "KRW"
    Next krCode

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
    Dim cleanErrNum As Long: cleanErrNum = Err.Number
    Dim cleanErrSource As String: cleanErrSource = Err.Source
    Dim cleanErrDesc As String: cleanErrDesc = Err.Description
    If cleanErrNum <> 0 Then
        intFailCnt = intFailCnt + 1
        strErrLog = strErrLog & vbCrLf & "[RunKRStatement] 错误号=" & cleanErrNum & _
                    "; 来源=" & cleanErrSource & "; " & cleanErrDesc
        If Len(strCode) > 0 Then
            AddMissingDiagnosticsForCompany strCode, strKind, conceptMap, collDiagRows, _
                                            "run_failed: " & cleanErrDesc
        Else
            AddDiagnosticRow collDiagRows, "(RunKRStatement)", strKind, targetSheet, _
                             "MISSING", "stockanalysis.com", "HTML", "—", _
                             "KRW billions", "—", "run_failed: " & cleanErrDesc
        End If
        Err.Clear
    End If

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
        msg = targetSheet & " 抓取完成 (单位: 十亿韩元)" & vbCrLf & _
              "用时: " & Format(Timer - dtTime, "0.0 秒") & vbCrLf & _
              "公司数: " & collCodes.Count & " / 期数: " & dictPeriodSet.Count
        If intFailCnt > 0 Then msg = msg & vbCrLf & vbCrLf & "失败 " & intFailCnt & " 条:" & strErrLog

        Dim style As Long: style = vbInformation
        If intFailCnt > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询 (KR)"
    End If
End Sub


Private Sub FetchKRFromStockAnalysis(ByVal strTicker As String, _
                                     ByVal strKind As String, _
                                     ByVal conceptMap As Variant, _
                                     ByVal strQuarter As String, _
                                     ByVal lngYear As Long, _
                                     ByVal maxPeriods As Long, _
                                     ByRef dictData As Object, _
                                     ByRef dictPeriodSet As Object, _
                                     ByRef dictIndicatorSet As Object, _
                                     ByRef dictCategoryMap As Object, _
                                     Optional ByVal collDiagRows As Collection = Nothing)
    Dim stage As String: stage = "init"
    On Error GoTo KrErr

    Dim mapSA As Object: Set mapSA = StockAnalysisFieldMapForKindKR(strKind)
    Dim dictMatchInfo As Object: Set dictMatchInfo = CreateObject("Scripting.Dictionary")
    dictMatchInfo.CompareMode = vbTextCompare

    Dim fetchAnnual As Boolean, fetchQuarterly As Boolean
    Select Case UCase$(Trim$(strQuarter))
        Case "Q1", "Q2", "Q3"
            fetchQuarterly = True
        Case "Q4"
            fetchAnnual = True
        Case Else
            ' 全部: 同时取季度页 + 年度页; 年度页后写, 让 12-31 保留 FY 年报口径。
            fetchQuarterly = True
            fetchAnnual = True
    End Select

    Dim includePriorYear As Boolean
    includePriorYear = (lngYear > 0 And (strKind = "BalanceSheet" Or strKind = "Income"))

    If fetchQuarterly Then
        stage = "QuarterlyPage"
        ParseKRStockAnalysisPage strTicker, strKind, conceptMap, mapSA, True, _
                                 strQuarter, lngYear, includePriorYear, dictData, dictPeriodSet, _
                                 dictIndicatorSet, dictCategoryMap, dictMatchInfo
    End If

    If fetchAnnual Then
        stage = "AnnualPage"
        ParseKRStockAnalysisPage strTicker, strKind, conceptMap, mapSA, False, _
                                 strQuarter, lngYear, includePriorYear, dictData, dictPeriodSet, _
                                 dictIndicatorSet, dictCategoryMap, dictMatchInfo
    End If

    stage = "WriteDiagnostics"
    AppendKRDiagnosticsForConceptMap strTicker, strKind, conceptMap, dictMatchInfo, collDiagRows

    If dictPeriodSet.Count = 0 Then
        Err.Raise vbObjectError + 742, "FetchKRFromStockAnalysis", _
            "stockanalysis KR 无匹配期间: " & strTicker & " / " & strQuarter & " / " & lngYear
    End If

    Err.Clear
    Exit Sub

KrErr:
    Dim origNum As Long: origNum = Err.Number
    Dim origSource As String: origSource = Err.Source
    Dim origDesc As String: origDesc = Err.Description
    Err.Clear
    Err.Raise vbObjectError + 790, "FetchKRFromStockAnalysis", _
        "[stage=" & stage & "] 原始错误号=" & origNum & _
        "; 原始来源=" & origSource & "; " & origDesc
End Sub


Private Sub ParseKRStockAnalysisPage(ByVal strTicker As String, _
                                     ByVal strKind As String, _
                                     ByVal conceptMap As Variant, _
                                     ByVal mapSA As Object, _
                                     ByVal quarterly As Boolean, _
                                     ByVal strQuarter As String, _
                                     ByVal lngYear As Long, _
                                     ByVal includePriorYear As Boolean, _
                                     ByRef dictData As Object, _
                                     ByRef dictPeriodSet As Object, _
                                     ByRef dictIndicatorSet As Object, _
                                     ByRef dictCategoryMap As Object, _
                                     ByRef dictMatchInfo As Object)
    Dim strUrl As String: strUrl = StockAnalysisKRUrl(strTicker, strKind, quarterly)
    Dim cacheKey As String
    cacheKey = "stockanalysis_KR_" & NormalizeKRTicker(strTicker) & "_" & strKind & "_" & _
               IIf(quarterly, "quarterly", "annual")
    Dim strHtml As String: strHtml = StockAnalysisHttpGet(strUrl, cacheKey)

    Dim objHtml As Object: Set objHtml = CreateObject("htmlfile")
    objHtml.Open
    objHtml.Write strHtml
    objHtml.Close

    Dim tables As Object: Set tables = objHtml.getElementsByTagName("table")
    If tables Is Nothing Or tables.Length = 0 Then
        Err.Raise vbObjectError + 743, "ParseKRStockAnalysisPage", _
            "stockanalysis 页面未找到 HTML table: " & strUrl
    End If

    Dim objTb As Object: Set objTb = tables.Item(0)
    If objTb.Rows.Length < 3 Then
        Err.Raise vbObjectError + 744, "ParseKRStockAnalysisPage", _
            "stockanalysis table 行数不足: " & strUrl
    End If

    Dim headerRow As Object: Set headerRow = objTb.Rows.Item(0)
    Dim periods As Object: Set periods = CreateObject("Scripting.Dictionary")
    Dim j As Long, headerText As String, periodEnd As String

    For j = 1 To headerRow.Cells.Length - 1
        headerText = KRCellText(headerRow, j)
        periodEnd = KRPeriodFromHeader(headerText)
        If Len(periodEnd) > 0 Then
            If MatchKRPeriod(periodEnd, strQuarter, lngYear, headerText, includePriorYear) Then
                periods.Item(CStr(j)) = periodEnd
            End If
        End If
    Next j

    If periods.Count = 0 Then Exit Sub

    Dim dictRows As Object: Set dictRows = CreateObject("Scripting.Dictionary")
    dictRows.CompareMode = vbTextCompare
    Dim r As Long, rowLabel As String
    For r = 2 To objTb.Rows.Length - 1
        rowLabel = KRCellText(objTb.Rows.Item(r), 0)
        If Len(rowLabel) > 0 Then
            If Not dictRows.Exists(rowLabel) Then dictRows.Add rowLabel, objTb.Rows.Item(r)
        End If
    Next r

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
    Dim scaleDivisor As Double, unitText As String, dictPer As Object

    For ci = LBound(conceptMap) To UBound(conceptMap)
        mapEntry = conceptMap(ci)
        strCat = MapEntryCategory(mapEntry)
        strLabel = MapEntryLabel(mapEntry)
        If Not mapSA.Exists(strLabel) Then GoTo NextConcept

        candidates = mapSA.Item(strLabel)
        totalCand = UBound(candidates) - LBound(candidates) + 1
        candIdx = 0
        For Each cand In candidates
            candIdx = candIdx + 1
            If dictRows.Exists(CStr(cand)) Then
                Set objRow = dictRows.Item(CStr(cand))
                scaleDivisor = KRMapEntryScale(mapEntry)
                unitText = KRMapEntryUnit(mapEntry)
                Dim matchPeriodEnd As String: matchPeriodEnd = ""

                For Each key In periods.Keys
                    j = CLng(key)
                    If j < objRow.Cells.Length Then
                        rawValue = KRCellText(objRow, j)
                        If KRTryScaledValue(rawValue, scaleDivisor, scaledValue) Then
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
                            If dictPer.Exists(strLabel) Then
                                dictPer.Item(strLabel) = scaledValue
                            Else
                                dictPer.Add strLabel, scaledValue
                            End If
                        End If
                    End If
                Next key

                AddOrUpdateKRMatch dictMatchInfo, strLabel, CStr(cand), unitText, candIdx, totalCand, matchPeriodEnd
                Exit For
            End If
        Next cand
NextConcept:
    Next ci
End Sub


Private Function StockAnalysisKRUrl(ByVal strTicker As String, _
                                    ByVal strKind As String, _
                                    ByVal quarterly As Boolean) As String
    Dim pathPart As String
    Select Case strKind
        Case "BalanceSheet": pathPart = "financials/balance-sheet/"
        Case "Income":       pathPart = "financials/"
        Case "CashFlow":     pathPart = "financials/cash-flow-statement/"
        Case Else
            Err.Raise vbObjectError + 745, "StockAnalysisKRUrl", _
                "stockanalysis KR 不支持报表类型: " & strKind
    End Select

    StockAnalysisKRUrl = "https://stockanalysis.com/quote/krx/" & NormalizeKRTicker(strTicker) & "/" & pathPart
    If quarterly Then StockAnalysisKRUrl = StockAnalysisKRUrl & "?p=quarterly"
End Function


Private Function StockAnalysisHttpGet(ByVal strUrl As String, Optional ByVal cacheKey As String = "") As String
    Dim result As THttpResult
    StockAnalysisHttpGet = RunCachedHttpGet(strUrl, cacheKey, "STOCKANALYSIS_KR", 24, result)
End Function


Private Function StockAnalysisFieldMapForKindKR(ByVal strKind As String) As Object
    Dim mapSA As Object: Set mapSA = CreateObject("Scripting.Dictionary")
    mapSA.CompareMode = vbTextCompare

    Select Case strKind
        Case "BalanceSheet"
            mapSA.Add "Cash & equivalents", Array("Cash & Equivalents")
            mapSA.Add "Accounts receivable, net", Array("Accounts Receivable", "Receivables")
            mapSA.Add "Inventory", Array("Inventory")
            mapSA.Add "Total current assets", Array("Total Current Assets")
            mapSA.Add "Property, plant & equipment, net", Array("Property, Plant & Equipment")
            mapSA.Add "Investments", Array("Long-Term Investments", "Short-Term Investments")
            mapSA.Add "Total assets", Array("Total Assets")
            mapSA.Add "Accounts payable", Array("Accounts Payable")
            mapSA.Add "Short-term debt", Array("Short-Term Debt", "Current Portion of Long-Term Debt")
            mapSA.Add "Total current liabilities", Array("Total Current Liabilities")
            mapSA.Add "Long-term debt", Array("Long-Term Debt")
            mapSA.Add "Total liabilities", Array("Total Liabilities")
            mapSA.Add "Minority interests", Array("Minority Interest")
            mapSA.Add "Total equity", Array("Shareholders' Equity", "Total Common Equity")
            mapSA.Add "Total stockholders' equity", Array("Shareholders' Equity", "Total Common Equity")
            mapSA.Add "Total liabilities & equity", Array("Total Liabilities & Equity")

        Case "Income"
            mapSA.Add "Revenue", Array("Revenue")
            mapSA.Add "Cost of goods & services sold", Array("Cost of Revenue")
            mapSA.Add "Gross profit", Array("Gross Profit")
            mapSA.Add "R&D expense", Array("Research & Development")
            mapSA.Add "SG&A expense", Array("Selling, General & Admin")
            mapSA.Add "Total operating expenses", Array("Operating Expenses")
            mapSA.Add "Operating income", Array("Operating Income")
            mapSA.Add "Interest expense", Array("Interest Expense")
            mapSA.Add "Pre-tax income", Array("Pretax Income", "EBT Excluding Unusual Items")
            mapSA.Add "Income tax expense", Array("Income Tax Expense")
            mapSA.Add "Net income", Array("Net Income")
            mapSA.Add "Basic EPS", Array("EPS (Basic)")
            mapSA.Add "Diluted EPS", Array("EPS (Diluted)")

        Case "CashFlow"
            mapSA.Add "Net income", Array("Net Income")
            mapSA.Add "Depreciation & amortization", Array("Depreciation & Amortization")
            mapSA.Add "Change in accounts receivable", Array("Change in Accounts Receivable")
            mapSA.Add "Change in inventory", Array("Change in Inventory")
            mapSA.Add "Change in accounts payable", Array("Change in Accounts Payable")
            mapSA.Add "Cash from operations", Array("Operating Cash Flow")
            mapSA.Add "Capex", Array("Capital Expenditures")
            mapSA.Add "Cash from investing", Array("Investing Cash Flow")
            mapSA.Add "Cash from financing", Array("Financing Cash Flow")
            mapSA.Add "Dividends paid", Array("Common Dividends Paid")
            mapSA.Add "FX effect on cash", Array("Foreign Exchange Rate Adjustments")
            mapSA.Add "Net cash flow", Array("Net Cash Flow")
            mapSA.Add "Free cash flow", Array("Free Cash Flow")

        Case Else
            Err.Raise vbObjectError + 747, "StockAnalysisFieldMapForKindKR", _
                "stockanalysis KR 不支持报表类型: " & strKind
    End Select

    Set StockAnalysisFieldMapForKindKR = mapSA
End Function


Private Function MatchKRPeriod(ByVal periodEnd As String, _
                               ByVal strQuarter As String, _
                               ByVal lngYear As Long, _
                               ByVal headerText As String, _
                               Optional ByVal includePriorYear As Boolean = False) As Boolean
    MatchKRPeriod = False
    If Len(periodEnd) < 10 Then Exit Function

    If lngYear > 0 Then
        Dim periodYear As Long: periodYear = CLng(Left$(periodEnd, 4))
        If includePriorYear Then
            If periodYear <> lngYear And periodYear <> lngYear - 1 Then Exit Function
        Else
            If periodYear <> lngYear Then Exit Function
        End If
    End If

    Select Case UCase$(Trim$(strQuarter))
        Case "全部", ""
            MatchKRPeriod = True
        Case "Q1"
            MatchKRPeriod = (Right$(periodEnd, 5) = "03-31")
        Case "Q2"
            MatchKRPeriod = (Right$(periodEnd, 5) = "06-30")
        Case "Q3"
            MatchKRPeriod = (Right$(periodEnd, 5) = "09-30")
        Case "Q4"
            MatchKRPeriod = (Right$(periodEnd, 5) = "12-31" And Left$(UCase$(Trim$(headerText)), 2) = "FY")
        Case Else
            MatchKRPeriod = True
    End Select
End Function


Private Function KRPeriodFromHeader(ByVal headerText As String) As String
    Dim s As String: s = UCase$(KRCleanText(headerText))
    If Len(s) = 0 Or s = "TTM" Or s = "CURRENT" Then Exit Function

    Dim parts As Variant: parts = Split(s, " ")
    If UBound(parts) < 1 Then Exit Function

    Dim y As Long, q As String, monthDay As String
    If parts(0) = "FY" And IsNumeric(parts(1)) Then
        y = CLng(parts(1))
        KRPeriodFromHeader = Format$(DateSerial(y, 12, 31), "yyyy-mm-dd")
    ElseIf Left$(parts(0), 1) = "Q" And Len(parts(0)) >= 2 And IsNumeric(parts(1)) Then
        q = Mid$(parts(0), 2, 1)
        y = CLng(parts(1))
        Select Case q
            Case "1": monthDay = "03-31"
            Case "2": monthDay = "06-30"
            Case "3": monthDay = "09-30"
            Case "4": monthDay = "12-31"
            Case Else: Exit Function
        End Select
        KRPeriodFromHeader = CStr(y) & "-" & monthDay
    End If
End Function


Private Function NormalizeKRTicker(ByVal rawCode As String) As String
    Dim s As String: s = UCase$(Trim$(rawCode))
    s = Replace(s, "KRX", "", 1, -1, vbTextCompare)
    s = Replace(s, "KR", "", 1, -1, vbTextCompare)
    s = Replace(s, "A", "", 1, -1, vbTextCompare)
    s = Replace(s, ".", "")
    s = Replace(s, " ", "")
    If Len(s) > 0 And IsNumeric(s) Then s = CStr(CLng(s))
    NormalizeKRTicker = Right$("000000" & s, 6)
End Function


Private Function KRCellText(ByVal objRow As Object, ByVal colIndex As Long) As String
    On Error Resume Next
    KRCellText = KRCleanText(CStr(objRow.Cells.Item(colIndex).innerText))
    Err.Clear
    On Error GoTo 0
End Function


Private Function KRCleanText(ByVal rawText As String) As String
    Dim s As String: s = CStr(rawText)
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    KRCleanText = Trim$(s)
End Function


Private Function KRTryScaledValue(ByVal rawValue As String, _
                                  ByVal scaleDivisor As Double, _
                                  ByRef scaledValue As Double) As Boolean
    Dim s As String: s = KRCleanText(rawValue)
    If Len(s) = 0 Or s = "-" Or s = "--" Then Exit Function

    s = Replace(s, ",", "")
    s = Replace(s, "−", "-")
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then _
        s = "-" & Mid$(s, 2, Len(s) - 2)

    Dim isPct As Boolean
    If Right$(s, 1) = "%" Then
        isPct = True
        s = Left$(s, Len(s) - 1)
    End If

    If Not IsNumeric(s) Then Exit Function
    If scaleDivisor = 0 Then scaleDivisor = 1
    scaledValue = CDbl(s)
    If isPct Then
        scaledValue = scaledValue / 100#
    Else
        scaledValue = scaledValue / scaleDivisor
    End If
    KRTryScaledValue = True
End Function


Private Function KRMapEntryScale(ByVal entry As Variant) As Double
    On Error Resume Next
    If UBound(entry) >= 4 Then
        KRMapEntryScale = CDbl(entry(4))
    Else
        KRMapEntryScale = 1000#
    End If
    If Err.Number <> 0 Or KRMapEntryScale = 0 Then
        Err.Clear
        KRMapEntryScale = 1000#
    End If
    On Error GoTo 0
End Function


Private Function KRMapEntryUnit(ByVal entry As Variant) As String
    On Error Resume Next
    If UBound(entry) >= 3 Then KRMapEntryUnit = CStr(entry(3))
    If Err.Number <> 0 Or Len(KRMapEntryUnit) = 0 Then
        Err.Clear
        KRMapEntryUnit = "KRW billions"
    End If
    On Error GoTo 0
End Function


Private Sub AddOrUpdateKRMatch(ByVal dictMatchInfo As Object, _
                               ByVal label As String, _
                               ByVal fieldName As String, _
                               ByVal unitText As String, _
                               ByVal candIdx As Long, _
                               ByVal totalCand As Long, _
                               ByVal periodEnd As String)
    Dim scoreText As String: scoreText = CStr(candIdx) & "/" & CStr(totalCand)
    If dictMatchInfo.Exists(label) Then
        Dim oldInfo As Variant: oldInfo = dictMatchInfo.Item(label)
        Dim oldIdx As Long: oldIdx = CLng(Split(CStr(oldInfo(2)), "/")(0))
        If candIdx > oldIdx Then Exit Sub
        If candIdx = oldIdx Then
            If periodEnd > CStr(oldInfo(3)) Then oldInfo(3) = periodEnd
            dictMatchInfo.Item(label) = oldInfo
            Exit Sub
        End If
        dictMatchInfo.Item(label) = Array(fieldName, unitText, scoreText, periodEnd)
    Else
        dictMatchInfo.Add label, Array(fieldName, unitText, scoreText, periodEnd)
    End If
End Sub


Private Sub AppendKRDiagnosticsForConceptMap(ByVal ticker As String, _
                                             ByVal strKind As String, _
                                             ByVal conceptMap As Variant, _
                                             ByVal dictMatchInfo As Object, _
                                             ByVal collRows As Collection)
    If collRows Is Nothing Then Exit Sub

    Dim i As Long, mapEntry As Variant, label As String, info As Variant
    For i = LBound(conceptMap) To UBound(conceptMap)
        mapEntry = conceptMap(i)
        label = MapEntryLabel(mapEntry)
        If dictMatchInfo.Exists(label) Then
            info = dictMatchInfo.Item(label)
            Dim fxText As String: fxText = FxRateTextForDiagnostic("KRW", CStr(info(3)), strKind)
            AddDiagnosticRow collRows, ticker, strKind, label, "OK_STOCKANALYSIS", _
                             "stockanalysis.com", "HTML", CStr(info(0)), _
                             CStr(info(1)), CStr(info(2)), "exact field match via htmlfile DOM", fxText
        Else
            AddDiagnosticRow collRows, ticker, strKind, label, "MISSING", _
                             "stockanalysis.com", "HTML", "—", "KRW billions", "—", _
                             "stockanalysis table did not contain mapped field"
        End If
    Next i
End Sub


Private Function HasAnyCoreLabelKR(ByVal dictData As Object, _
                                   ByVal ticker As String, _
                                   ByVal coreLabels As Variant) As Boolean
    On Error Resume Next
    If Not dictData.Exists(ticker) Then GoTo CleanExit
    Dim dictCompany As Object: Set dictCompany = dictData.Item(ticker)
    Dim p As Variant, i As Long, dictPer As Object
    For Each p In dictCompany.Keys
        Set dictPer = dictCompany.Item(p)
        For i = LBound(coreLabels) To UBound(coreLabels)
            If dictPer.Exists(CStr(coreLabels(i))) Then
                HasAnyCoreLabelKR = True
                GoTo CleanExit
            End If
        Next i
    Next p
CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Sub CommitTempKRData(ByVal strCode As String, _
                             ByVal tempData As Object, _
                             ByVal tempPeriodSet As Object, _
                             ByVal tempIndicatorSet As Object, _
                             ByVal tempCategoryMap As Object, _
                             ByRef dictData As Object, _
                             ByRef dictPeriodSet As Object, _
                             ByRef dictIndicatorSet As Object, _
                             ByRef dictCategoryMap As Object)
    If tempData.Exists(strCode) Then
        If dictData.Exists(strCode) Then dictData.Remove strCode
        dictData.Add strCode, tempData.Item(strCode)
    End If

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


Private Function OrderIndicatorsByKRConceptMap(ByVal dictIndicatorSet As Object, _
                                               ByVal conceptMap As Variant) As Variant
    Dim coll As New Collection
    Dim i As Long, label As String
    For i = LBound(conceptMap) To UBound(conceptMap)
        label = MapEntryLabel(conceptMap(i))
        If dictIndicatorSet.Exists(label) Then coll.Add label
    Next i

    If coll.Count = 0 Then
        OrderIndicatorsByKRConceptMap = Array()
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(1 To coll.Count)
    For i = 1 To coll.Count
        arr(i) = CStr(coll.Item(i))
    Next i
    OrderIndicatorsByKRConceptMap = arr
End Function


Private Sub ClearKRWideTableOutput(ByVal wsTarget As Worksheet)
    On Error Resume Next
    wsTarget.UsedRange.UnMerge
    Dim lastRow As Long, lastCol As Long
    lastRow = wsTarget.UsedRange.Rows.Count + wsTarget.UsedRange.Row - 1
    lastCol = wsTarget.UsedRange.Columns.Count + wsTarget.UsedRange.Column - 1
    If lastRow < 2 Then lastRow = 2
    If lastCol < 3 Then lastCol = 3
    wsTarget.Range(wsTarget.Cells(1, 3), wsTarget.Cells(lastRow, lastCol)).Clear
    If lastRow >= 2 Then wsTarget.Range(wsTarget.Cells(2, 1), wsTarget.Cells(lastRow, 2)).Clear
    wsTarget.Range("A1").Value = "大类"
    wsTarget.Range("B1").Value = "指标名称"
    With wsTarget.Range("A1:B1")
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Err.Clear
    On Error GoTo 0
End Sub


Private Function KRByteToStr(arrByte, ByVal strCharSet As String) As String
    With CreateObject("Adodb.Stream")
        .Type = 1
        .Open
        .Write arrByte
        .Position = 0
        .Type = 2
        .Charset = strCharSet
        KRByteToStr = .ReadText
        .Close
    End With
End Function

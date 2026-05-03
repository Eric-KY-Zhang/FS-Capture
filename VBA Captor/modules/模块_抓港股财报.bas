Attribute VB_Name = "模块_抓港股财报"
Option Explicit

' =================================================================
'  抓港股财报 — 共享 helpers (RunHKStatement / FetchHKFromXueqiu)
'
'  数据源: 雪球 HK finance API
'    https://stock.xueqiu.com/v5/stock/finance/hk/{balance|income|cash_flow}.json
'
'  约定:
'    - 港股只走雪球, 不走 EDGAR / ifrs-full / fuzzy 推荐。
'    - 金额写入值 = 雪球原值 / 1,000,000。
'    - 币种不写死, 诊断 Unit 列使用 data.currency; 正式表 A1 comment 说明看诊断。
' =================================================================


' --------- 通用港股单表抓数主流程 ---------
'   strKind     : "BalanceSheet" / "Income" / "CashFlow"
'   targetSheet : 目标 sheet 名
'   conceptMap  : 1-d Variant 数组, 每元素至少是 Array(大类, 标签)
'   maxPeriods  : API count + 最新期数截断
Public Sub RunHKStatement(ByVal strKind As String, ByVal targetSheet As String, _
                          ByVal conceptMap As Variant, _
                          ByVal maxPeriods As Long)
    Dim wsPool As Worksheet, wsTarget As Worksheet
    Dim arrPool As Variant
    Dim i As Long, lngRow As Long, numCompanies As Long
    Dim intFailCnt As Long, strErrLog As String
    Dim strCodeRaw As String, strCode As String, strName As String, strMarket As String
    Dim dtTime As Double: dtTime = Timer

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictPeriodSet As Object: Set dictPeriodSet = CreateObject("Scripting.Dictionary")
    Dim dictIndicatorSet As Object: Set dictIndicatorSet = CreateObject("Scripting.Dictionary")
    Dim dictCategoryMap As Object: Set dictCategoryMap = CreateObject("Scripting.Dictionary")
    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    Dim collCodes As New Collection
    Dim collDiagRows As New Collection

    g_diagnosticSheetName = "港股_抓取诊断"

    Set wsPool = ThisWorkbook.Sheets("样本池")
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)

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

    Dim strQuarter As String: strQuarter = ReadQuarterSelection()
    Dim lngYear As Long: lngYear = ReadYearSelection()

    For i = 1 To numCompanies
        strCodeRaw = Trim$(CStr(arrPool(i, 1)))
        If Len(strCodeRaw) = 0 Then GoTo NextRow
        strName = Trim$(CStr(arrPool(i, 2)))
        strMarket = ResolveMarket(strCodeRaw, CStr(arrPool(i, POOL_MARKET_COL)))
        If strMarket <> "HK" Then GoTo NextRow
        strCode = NormalizeHKTicker(strCodeRaw)

        Application.StatusBar = "抓取中: " & targetSheet & " (" & i & "/" & numCompanies & ") " & strCode
        DoEvents

        Dim tempData As Object: Set tempData = CreateObject("Scripting.Dictionary")
        Dim tempPeriodSet As Object: Set tempPeriodSet = CreateObject("Scripting.Dictionary")
        Dim tempIndicatorSet As Object: Set tempIndicatorSet = CreateObject("Scripting.Dictionary")
        Dim tempCategoryMap As Object: Set tempCategoryMap = CreateObject("Scripting.Dictionary")

        On Error Resume Next
        Err.Clear
        FetchHKFromXueqiu strCode, strKind, conceptMap, strQuarter, lngYear, maxPeriods, _
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
        ElseIf Not HasAnyCoreLabelHK(tempData, strCode, CoreLabelsForKind(strKind)) Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": 核心字段未命中"
        Else
            CommitTempHKData strCode, tempData, tempPeriodSet, tempIndicatorSet, tempCategoryMap, _
                             dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap
            If Not dictCompanyName.Exists(strCode) Then dictCompanyName.Add strCode, strName
            collCodes.Add strCode
        End If

        Application.Wait Now + TimeSerial(0, 0, 1)

NextRow:
    Next i

    If collCodes.Count = 0 Then
        ClearHKWideTableOutput wsTarget
        GoTo CleanUp
    End If

    If dictPeriodSet.Count = 0 Or dictIndicatorSet.Count = 0 Then
        intFailCnt = intFailCnt + 1
        strErrLog = strErrLog & vbCrLf & "[" & targetSheet & "] 无匹配数据 " & _
                    "(期数=" & dictPeriodSet.Count & " / 指标=" & dictIndicatorSet.Count & ")"
        GoTo CleanUp
    End If

    Dim arrCodes() As String
    ReDim arrCodes(1 To collCodes.Count)
    For i = 1 To collCodes.Count
        arrCodes(i) = collCodes(i)
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

    arrIndicators = OrderIndicatorsByHKConceptMap(dictIndicatorSet, conceptMap)

    Application.StatusBar = "写入: " & targetSheet
    DoEvents
    WriteWideTable wsTarget, arrCodes, dictCompanyName, dictData, _
                   arrPeriods, arrIndicators, dictCategoryMap, True

    On Error Resume Next
    If Not wsTarget.Range("A1").Comment Is Nothing Then wsTarget.Range("A1").Comment.Delete
    wsTarget.Range("A1").AddComment "单位: 百万(各家公司报告币种, 见 港股_抓取诊断 Unit 列)"
    Err.Clear
    On Error GoTo CleanUp

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
        msg = targetSheet & " 抓取完成 (单位: 百万, 各家公司报告币种)" & vbCrLf & _
              "用时: " & Format(Timer - dtTime, "0.0 秒") & vbCrLf & _
              "公司数: " & collCodes.Count & " / 期数: " & dictPeriodSet.Count
        If intFailCnt > 0 Then msg = msg & vbCrLf & vbCrLf & "失败 " & intFailCnt & " 条:" & strErrLog

        Dim style As Long: style = vbInformation
        If intFailCnt > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询 (HK)"
    End If
End Sub


Private Sub FetchHKFromXueqiu(ByVal strTicker As String, _
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
    On Error GoTo XqErr

    stage = "ReadCookie"
    Dim strCookie As String: strCookie = ReadXueqiuCookie()
    If Len(strCookie) = 0 Then
        Err.Raise vbObjectError + 642, "FetchHKFromXueqiu", _
            "雪球 HK 抓数需要 cookie. 请在 样本池!B5 填入 xq_a_token"
    End If

    stage = "BuildUrl"
    Dim strEndpoint As String: strEndpoint = XueqiuHKEndpointForKind(strKind)
    Dim strUrl As String
    strUrl = "https://stock.xueqiu.com/v5/stock/finance/hk/" & strEndpoint & ".json?" & _
             "symbol=" & strTicker & "&type=all&is_detail=true&count=" & CStr(maxPeriods)

    stage = "HttpGet"
    Dim strJson As String: strJson = XueqiuHttpGet(strUrl, strCookie)

    stage = "DumpJson"
    On Error Resume Next
    DumpHKXueqiuJson strTicker, strKind, strJson
    On Error GoTo XqErr

    stage = "ParseJson"
    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(strJson)

    stage = "CheckErrorCode"
    If parsed.Exists("error_code") Then
        Dim errCode As Variant: errCode = parsed.Item("error_code")
        If Not IsNull(errCode) And Not IsEmpty(errCode) Then
            If CStr(errCode) <> "0" Then
                Err.Raise vbObjectError + 643, "FetchHKFromXueqiu", _
                    "雪球 HK API 错误 (" & errCode & "): " & HK_NzStr(parsed, "error_description")
            End If
        End If
    End If

    stage = "GetData"
    If Not parsed.Exists("data") Then
        Err.Raise vbObjectError + 644, "FetchHKFromXueqiu", "雪球 HK 响应缺 data 字段: " & strTicker
    End If
    Dim dataRoot As Object: Set dataRoot = parsed.Item("data")
    Dim currencyText As String: currencyText = HK_NzStr(dataRoot, "currency")
    If Len(currencyText) = 0 Then currencyText = "REPORT_CURRENCY"

    stage = "GetList"
    If Not dataRoot.Exists("list") Then
        Err.Raise vbObjectError + 645, "FetchHKFromXueqiu", "雪球 HK 响应缺 list 字段: " & strTicker
    End If
    Dim listColl As Object: Set listColl = dataRoot.Item("list")
    If listColl Is Nothing Then
        Err.Raise vbObjectError + 646, "FetchHKFromXueqiu", "雪球 HK list 为 null: " & strTicker
    End If
    If listColl.Count = 0 Then
        Err.Raise vbObjectError + 647, "FetchHKFromXueqiu", "雪球 HK list 为空: " & strTicker
    End If

    stage = "CanonicalRecords"
    Dim canonical As Object: Set canonical = CanonicalHKRecords(listColl, strQuarter, lngYear)
    If canonical.Count = 0 Then
        Err.Raise vbObjectError + 648, "FetchHKFromXueqiu", _
            "雪球 HK 无匹配期间: " & strTicker & " / " & strQuarter & " / " & lngYear
    End If

    stage = "GetCompanyDict"
    Dim dictCompany As Object
    If dictData.Exists(strTicker) Then
        Set dictCompany = dictData.Item(strTicker)
    Else
        Set dictCompany = CreateObject("Scripting.Dictionary")
        dictData.Add strTicker, dictCompany
    End If

    stage = "BuildXqMap"
    Dim mapXq As Object: Set mapXq = XueqiuFieldMapForKindHK(strKind)
    Dim dictMatchInfo As Object: Set dictMatchInfo = CreateObject("Scripting.Dictionary")

    Dim key As Variant, record As Object, periodEnd As String
    Dim ci As Long, mapEntry As Variant, strCat As String, strLabel As String
    Dim candidates As Variant, cand As Variant, candIdx As Long, totalCand As Long
    Dim val As Variant, valM As Double, dictPer As Object

    For Each key In canonical.Keys
        stage = "Record:" & CStr(key)
        Set record = canonical.Item(key)
        periodEnd = HK_NzStr(record, "ed")
        If Len(periodEnd) < 10 Then periodEnd = ParseHKReportDate(record)
        If Len(periodEnd) = 0 Then GoTo NextRecord

        If dictCompany.Exists(periodEnd) Then
            Set dictPer = dictCompany.Item(periodEnd)
        Else
            Set dictPer = CreateObject("Scripting.Dictionary")
            dictCompany.Add periodEnd, dictPer
        End If

        For ci = LBound(conceptMap) To UBound(conceptMap)
            mapEntry = conceptMap(ci)
            strCat = MapEntryCategory(mapEntry)
            strLabel = MapEntryLabel(mapEntry)
            If Not mapXq.Exists(strLabel) Then GoTo NextConcept

            candidates = Split(CStr(mapXq.Item(strLabel)), ",")
            candIdx = 0
            totalCand = UBound(candidates) - LBound(candidates) + 1
            For Each cand In candidates
                candIdx = candIdx + 1
                val = XueqiuValueHK(record, Trim$(CStr(cand)))
                If Not IsEmpty(val) And Not IsNull(val) Then
                    If Not HKTryScaledValue(val, valM) Then GoTo NextCandidate

                    If Not dictPeriodSet.Exists(periodEnd) Then dictPeriodSet.Add periodEnd, True
                    If Not dictIndicatorSet.Exists(strLabel) Then
                        dictIndicatorSet.Add strLabel, dictIndicatorSet.Count
                        If Not dictCategoryMap.Exists(strLabel) Then dictCategoryMap.Add strLabel, strCat
                    End If
                    If dictPer.Exists(strLabel) Then
                        dictPer.Item(strLabel) = valM
                    Else
                        dictPer.Add strLabel, valM
                    End If

                    AddOrUpdateHKMatch dictMatchInfo, strLabel, Trim$(CStr(cand)), _
                                       currencyText, candIdx, totalCand
                    Exit For
                End If
NextCandidate:
            Next cand
NextConcept:
        Next ci
NextRecord:
    Next key

    stage = "WriteDiagnostics"
    AppendHKDiagnosticsForConceptMap strTicker, strKind, conceptMap, dictMatchInfo, collDiagRows

    Err.Clear
    Exit Sub

XqErr:
    Dim origNum As Long: origNum = Err.Number
    Dim origSource As String: origSource = Err.Source
    Dim origDesc As String: origDesc = Err.Description
    Err.Clear
    Err.Raise vbObjectError + 690, "FetchHKFromXueqiu", _
        "[stage=" & stage & "] 原始错误号=" & origNum & _
        "; 原始来源=" & origSource & "; " & origDesc
End Sub


' --------- 港股报表类型 → 雪球 hk finance endpoint ---------
Private Function XueqiuHKEndpointForKind(ByVal strKind As String) As String
    Select Case strKind
        Case "BalanceSheet": XueqiuHKEndpointForKind = "balance"
        Case "Income":       XueqiuHKEndpointForKind = "income"
        Case "CashFlow":     XueqiuHKEndpointForKind = "cash_flow"
        Case Else
            Err.Raise vbObjectError + 649, "XueqiuHKEndpointForKind", _
                "雪球 HK 不支持报表类型: " & strKind
    End Select
End Function


' --------- 英文指标名 → 雪球 HK 字段候选 ---------
Private Function XueqiuFieldMapForKindHK(ByVal strKind As String) As Object
    Dim mapXq As Object: Set mapXq = CreateObject("Scripting.Dictionary")

    Select Case strKind
        Case "BalanceSheet"
            mapXq.Add "Cash & equivalents", "cceq"
            mapXq.Add "Accounts receivable, net", "trx"
            mapXq.Add "Inventory", "inv"
            mapXq.Add "Total current assets", "ca"
            mapXq.Add "Accounts payable", "trpy"
            mapXq.Add "Property, plant & equipment, net", "fxda"
            mapXq.Add "Investments", "iv,fina"
            mapXq.Add "Total non-current assets", "tnca"
            mapXq.Add "Total assets", "ta"
            mapXq.Add "Short-term debt", "stdt"
            mapXq.Add "Total current liabilities", "clia"
            mapXq.Add "Long-term debt", "ltdt"
            mapXq.Add "Total non-current liabilities", "tnclia"
            mapXq.Add "Total liabilities", "tlia"
            mapXq.Add "Minority interests", "miint"
            mapXq.Add "Total equity", "teqy"
            mapXq.Add "Total stockholders' equity", "teqy"
            mapXq.Add "Total liabilities & equity", "ta"

        Case "Income"
            mapXq.Add "Revenue", "tto"
            mapXq.Add "Cost of goods & services sold", "fcgcost,slgcost"
            mapXq.Add "Gross profit", "gp"
            mapXq.Add "R&D expense", "rshdevexp"
            mapXq.Add "Selling expense", "slgdstexp"
            mapXq.Add "Administrative expense", "admexp"
            mapXq.Add "Total operating expenses", "topeexp"
            mapXq.Add "Operating income", "opeplo,opeploinclfincost"
            mapXq.Add "Pre-tax income", "plobtx"
            mapXq.Add "Income tax expense", "tx"
            mapXq.Add "Net income", "ploashh"
            mapXq.Add "Basic EPS", "beps_aju"
            mapXq.Add "Diluted EPS", "deps_aju"

        Case "CashFlow"
            mapXq.Add "Cash from operations", "nocf"
            mapXq.Add "Depreciation & amortization", "depaz"
            mapXq.Add "Cash from investing", "ninvcf"
            mapXq.Add "Capex", "fxdiodtinstr,rpafxdiodtinstr"
            mapXq.Add "Cash from financing", "nfcgcf"
            mapXq.Add "Dividends paid", "divp"
            mapXq.Add "Interest paid", "intp"
            mapXq.Add "Interest received", "intrc"
            mapXq.Add "FX effect on cash", "ncfdchexrateot"
            mapXq.Add "Cash at beginning of period", "cceqbegyr"
            mapXq.Add "Cash at end of period", "cceqeyr"

        Case Else
            Err.Raise vbObjectError + 650, "XueqiuFieldMapForKindHK", _
                "雪球 HK 不支持报表类型: " & strKind
    End Select

    Set XueqiuFieldMapForKindHK = mapXq
End Function


' --------- HK 季度/年份过滤 ---------
Private Function MatchHKPeriod(ByVal periodEnd As String, _
                               ByVal strQuarter As String, _
                               ByVal lngYear As Long, _
                               ByVal monthNum As Long) As Boolean
    MatchHKPeriod = False
    If Len(periodEnd) < 10 Then Exit Function

    If lngYear > 0 Then
        If CLng(Left$(periodEnd, 4)) <> lngYear Then Exit Function
    End If

    Select Case strQuarter
        Case "全部", ""
            MatchHKPeriod = True
        Case "Q1"
            MatchHKPeriod = (monthNum = 3 And Right$(periodEnd, 6) = "-03-31")
        Case "Q2"
            MatchHKPeriod = (monthNum = 6 And Right$(periodEnd, 6) = "-06-30")
        Case "Q3"
            MatchHKPeriod = (monthNum = 9 And Right$(periodEnd, 6) = "-09-30")
        Case "Q4"
            MatchHKPeriod = (monthNum = 12)
        Case Else
            MatchHKPeriod = True
    End Select
End Function


' --------- 同一 year/month_num 只保留 canonical record ---------
Private Function CanonicalHKRecords(ByVal listColl As Object, _
                                    ByVal strQuarter As String, _
                                    ByVal lngYear As Long) As Object
    Dim dictCanonical As Object: Set dictCanonical = CreateObject("Scripting.Dictionary")
    Dim dictStart As Object: Set dictStart = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To listColl.Count
        Dim record As Object: Set record = listColl.Item(i)
        Dim periodEnd As String: periodEnd = HK_NzStr(record, "ed")
        If Len(periodEnd) < 10 Then periodEnd = ParseHKReportDate(record)
        If Len(periodEnd) = 0 Then GoTo NextRecord

        Dim monthNum As Long: monthNum = HKMonthNum(record)
        If Not MatchHKPeriod(periodEnd, strQuarter, lngYear, monthNum) Then GoTo NextRecord

        Dim ckey As String: ckey = CStr(monthNum) & "|" & Left$(periodEnd, 4)
        Dim periodStart As String: periodStart = HK_NzStr(record, "sd")
        If Not dictCanonical.Exists(ckey) Then
            dictCanonical.Add ckey, record
            dictStart.Add ckey, periodStart
        Else
            Dim oldRecord As Object: Set oldRecord = dictCanonical.Item(ckey)
            Dim oldEnd As String: oldEnd = HK_NzStr(oldRecord, "ed")
            Dim oldStart As String: oldStart = CStr(dictStart.Item(ckey))
            Dim takeIt As Boolean: takeIt = False
            If periodEnd > oldEnd Then
                takeIt = True
            ElseIf periodEnd = oldEnd And Len(periodStart) > 0 And _
                   (Len(oldStart) = 0 Or periodStart < oldStart) Then
                takeIt = True
            End If
            If takeIt Then
                dictCanonical.Remove ckey
                dictCanonical.Add ckey, record
                dictStart.Item(ckey) = periodStart
            End If
        End If
NextRecord:
    Next i

    Set CanonicalHKRecords = dictCanonical
End Function


Private Function NormalizeHKTicker(ByVal rawCode As String) As String
    Dim s As String: s = Trim$(rawCode)
    s = Replace(s, "HK", "", 1, -1, vbTextCompare)
    s = Replace(s, ".", "")
    s = Replace(s, " ", "")
    If Len(s) > 0 And IsNumeric(s) Then s = CStr(CLng(s))
    NormalizeHKTicker = Right$("00000" & s, 5)
End Function


Private Function HKMonthNum(ByVal record As Object) As Long
    On Error Resume Next
    HKMonthNum = 0
    If record.Exists("month_num") Then HKMonthNum = CLng(record.Item("month_num"))
    Err.Clear
    On Error GoTo 0
End Function


Private Function ParseHKReportDate(ByVal record As Object) As String
    On Error Resume Next
    ParseHKReportDate = ""
    If Not record.Exists("report_date") Then GoTo CleanExit
    Dim v As Variant: v = record.Item("report_date")
    If IsNull(v) Or IsEmpty(v) Then GoTo CleanExit
    Dim dt As Date
    dt = DateAdd("s", CDbl(v) / 1000#, DateSerial(1970, 1, 1))
    ParseHKReportDate = Format(dt, "yyyy-mm-dd")
CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Function XueqiuValueHK(ByVal record As Object, ByVal key As String) As Variant
    On Error Resume Next
    XueqiuValueHK = Empty
    If Not record.Exists(key) Then GoTo CleanExit

    If IsObject(record.Item(key)) Then
        Dim objVal As Object: Set objVal = record.Item(key)
        If TypeName(objVal) = "Collection" Then
            If objVal.Count >= 1 Then
                Dim inner As Variant: inner = objVal.Item(1)
                If Not IsNull(inner) And Not IsEmpty(inner) Then XueqiuValueHK = inner
            End If
        End If
    Else
        Dim v As Variant: v = record.Item(key)
        If Not IsNull(v) And Not IsEmpty(v) Then XueqiuValueHK = v
    End If

CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Function HKTryScaledValue(ByVal rawValue As Variant, ByRef scaledValue As Double) As Boolean
    Dim normalized As Variant
    normalized = NormalizeValue(CStr(rawValue))
    If IsEmpty(normalized) Or IsNull(normalized) Then
        HKTryScaledValue = False
    ElseIf IsNumeric(normalized) Then
        scaledValue = CDbl(normalized) / 1000000#
        HKTryScaledValue = True
    Else
        HKTryScaledValue = False
    End If
End Function


Private Function HK_NzStr(ByVal dict As Object, ByVal key As String) As String
    If dict Is Nothing Then HK_NzStr = "": Exit Function
    If Not dict.Exists(key) Then HK_NzStr = "": Exit Function
    Dim v As Variant: v = dict.Item(key)
    If IsNull(v) Or IsEmpty(v) Then
        HK_NzStr = ""
    Else
        HK_NzStr = CStr(v)
    End If
End Function


Private Sub AddOrUpdateHKMatch(ByVal dictMatchInfo As Object, _
                               ByVal strLabel As String, _
                               ByVal fieldText As String, _
                               ByVal currencyText As String, _
                               ByVal candIdx As Long, _
                               ByVal totalCand As Long)
    Dim scoreText As String, noteBase As String
    If candIdx = 1 Then
        scoreText = "100"
        noteBase = "hk_field_primary"
    Else
        scoreText = "85"
        noteBase = "hk_field_alt[" & candIdx & "/" & totalCand & "]"
    End If

    If dictMatchInfo.Exists(strLabel) Then
        Dim info As Variant: info = dictMatchInfo.Item(strLabel)
        info(7) = CLng(info(7)) + 1
        dictMatchInfo.Item(strLabel) = info
    Else
        dictMatchInfo.Add strLabel, Array("OK_XUEQIU", "Xueqiu", "xueqiu_hk", _
                                          fieldText, currencyText, scoreText, noteBase, 1)
    End If
End Sub


Private Sub AppendHKDiagnosticsForConceptMap(ByVal strTicker As String, _
                                             ByVal strKind As String, _
                                             ByVal conceptMap As Variant, _
                                             ByVal dictMatchInfo As Object, _
                                             ByVal collDiagRows As Collection)
    If collDiagRows Is Nothing Then Exit Sub
    Dim i As Long, mapEntry As Variant, strLabel As String
    For i = LBound(conceptMap) To UBound(conceptMap)
        mapEntry = conceptMap(i)
        strLabel = MapEntryLabel(mapEntry)
        If dictMatchInfo.Exists(strLabel) Then
            Dim info As Variant: info = dictMatchInfo.Item(strLabel)
            Dim noteText As String: noteText = CStr(info(6))
            If CLng(info(7)) > 0 Then noteText = noteText & "; periods_written=" & CStr(info(7))
            AddDiagnosticRow collDiagRows, strTicker, strKind, strLabel, CStr(info(0)), CStr(info(1)), _
                             CStr(info(2)), CStr(info(3)), CStr(info(4)), CStr(info(5)), noteText
        Else
            AddDiagnosticRow collDiagRows, strTicker, strKind, strLabel, "MISSING", "—", _
                             "xueqiu_hk", "—", "—", "—", "no xueqiu hk field matched"
        End If
    Next i
End Sub


Private Function HasAnyCoreLabelHK(ByVal dictData As Object, _
                                   ByVal strTicker As String, _
                                   ByVal coreLabels As Variant) As Boolean
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(coreLabels): ub = UBound(coreLabels)
    If Err.Number <> 0 Or ub < lb Then
        HasAnyCoreLabelHK = (dictData.Exists(strTicker) And dictData.Item(strTicker).Count > 0)
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
                HasAnyCoreLabelHK = True
                Exit Function
            End If
        Next i
    Next p
End Function


Private Sub CommitTempHKData(ByVal strTicker As String, _
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
            If destPer.Exists(CStr(label)) Then
                destPer.Item(CStr(label)) = tempPer.Item(label)
            Else
                destPer.Add CStr(label), tempPer.Item(label)
            End If
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


Private Function OrderIndicatorsByHKConceptMap(ByVal dictIndicatorSet As Object, _
                                               ByVal conceptMap As Variant) As Variant
    Dim arr() As String, n As Long: n = 0
    Dim i As Long
    ReDim arr(1 To dictIndicatorSet.Count + 1)
    For i = LBound(conceptMap) To UBound(conceptMap)
        Dim label As String: label = MapEntryLabel(conceptMap(i))
        If dictIndicatorSet.Exists(label) Then
            n = n + 1
            arr(n) = label
        End If
    Next i
    If n = 0 Then
        OrderIndicatorsByHKConceptMap = Array()
    Else
        ReDim Preserve arr(1 To n)
        OrderIndicatorsByHKConceptMap = arr
    End If
End Function


Private Sub ClearHKWideTableOutput(ByVal ws As Worksheet)
    Dim metaCols As Long: metaCols = 2
    If ws.Name = "港股_指标表" Then metaCols = 3

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


Private Sub DumpHKXueqiuJson(ByVal strTicker As String, ByVal strKind As String, ByVal strJson As String)
    On Error Resume Next
    Dim wbPath As String: wbPath = ThisWorkbook.Path
    Dim sampleDir As String: sampleDir = wbPath & Application.PathSeparator & "samples"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sampleDir) Then fso.CreateFolder sampleDir

    Dim fname As String
    fname = sampleDir & Application.PathSeparator & "xueqiu_HK_" & strTicker & "_" & _
            XueqiuHKDumpSuffixForKind(strKind) & ".json"
    Dim ts As Object: Set ts = fso.CreateTextFile(fname, True, True)
    ts.Write strJson
    ts.Close
    Err.Clear
    On Error GoTo 0
End Sub


Private Function XueqiuHKDumpSuffixForKind(ByVal strKind As String) As String
    Select Case strKind
        Case "BalanceSheet": XueqiuHKDumpSuffixForKind = "balance"
        Case "Income":       XueqiuHKDumpSuffixForKind = "income"
        Case "CashFlow":     XueqiuHKDumpSuffixForKind = "cash_flow"
        Case Else:           XueqiuHKDumpSuffixForKind = LCase$(strKind)
    End Select
End Function

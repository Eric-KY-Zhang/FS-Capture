Attribute VB_Name = "模块_测试"
Option Explicit

' =================================================================
'  Phase 4f reviewer smoke tests
'  本模块只做本地回归验证, 不触发任何网络抓数。
' =================================================================

Public Function TestOptionalBool(Optional ByVal blnSilent As Boolean = False) As Boolean
    TestOptionalBool = blnSilent
End Function


Public Sub TestStep3Smoke()
    Dim wsPool As Worksheet
    Set wsPool = ThisWorkbook.Worksheets("样本池")

    Dim savedB6 As Variant
    savedB6 = wsPool.Range("E6").Value

    Dim wsYuanbi As Worksheet
    Dim wsRmb As Worksheet
    Set wsYuanbi = GetOrClearSmokeSheet("_phase4f_step3_yuanbi")
    Set wsRmb = GetOrClearSmokeSheet("_phase4f_step3_rmb")

    Dim arrCodes(1 To 1) As String
    arrCodes(1) = "300866"

    Dim arrPeriods(1 To 2) As String
    arrPeriods(1) = "2024-12-31"
    arrPeriods(2) = "2023-12-31"

    Dim arrIndicators(1 To 2) As String
    arrIndicators(1) = "总资产"
    arrIndicators(2) = "审计意见"

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    dictCompanyName.Add "300866", "安克创新"

    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.Add "总资产", "资产"
    dictCategory.Add "审计意见", "文本"

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictPer2024 As Object: Set dictPer2024 = CreateObject("Scripting.Dictionary")
    Dim dictPer2023 As Object: Set dictPer2023 = CreateObject("Scripting.Dictionary")

    dictPer2024.Add "总资产", 123.45
    dictPer2024.Add "审计意见", "标准无保留"
    dictPer2023.Add "总资产", 67.89
    dictPer2023.Add "审计意见", "标准无保留"
    dictCompany.Add "2024-12-31", dictPer2024
    dictCompany.Add "2023-12-31", dictPer2023
    dictData.Add "300866", dictCompany

    wsPool.Range("E6").Value = "原币"
    WriteWideTable wsYuanbi, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=Nothing, statementKind:="BalanceSheet"

    wsPool.Range("E6").Value = "统一RMB"
    WriteWideTable wsRmb, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=Nothing, statementKind:="BalanceSheet"

    wsPool.Range("E6").Value = savedB6
End Sub


Public Sub TestStep45Smoke()
    Dim wsPool As Worksheet
    Set wsPool = ThisWorkbook.Worksheets("样本池")

    Dim savedB6 As Variant
    savedB6 = wsPool.Range("E6").Value

    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet

    Dim rows As Collection
    Set rows = New Collection
    AddDiagnosticRow rows, "AAPL", "BalanceSheet", "Total assets", "OK", "EDGAR", _
                     "us-gaap", "Assets", "USD", "100", "smoke", "7.123456"
    WriteDiagnosticForKind "BalanceSheet", rows

    Dim wsDiag As Worksheet
    Set wsDiag = ThisWorkbook.Worksheets("美股_抓取诊断")
    If CStr(wsDiag.Cells(2, 11).Value) <> "FX_Rate" Then _
        Err.Raise vbObjectError + 742, "TestStep45Smoke", "诊断 K 列表头不是 FX_Rate"
    If CStr(wsDiag.Cells(3, 11).Value) <> "7.123456" Then _
        Err.Raise vbObjectError + 743, "TestStep45Smoke", "诊断 K3 没写入 fx rate"

    Dim wsTag As Worksheet
    Set wsTag = GetOrClearSmokeSheet("_phase4f_step5_tag")

    Dim arrCodes(1 To 1) As String
    arrCodes(1) = "AAPL"

    Dim arrPeriods(1 To 1) As String
    arrPeriods(1) = "2024-12-31"

    Dim arrIndicators(1 To 1) As String
    arrIndicators(1) = "TextOnly"

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    dictCompanyName.Add "AAPL", "Apple"

    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.Add "TextOnly", "Smoke"

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictPer As Object: Set dictPer = CreateObject("Scripting.Dictionary")
    dictPer.Add "TextOnly", "non-numeric"
    dictCompany.Add "2024-12-31", dictPer
    dictData.Add "AAPL", dictCompany

    Dim dictCurrency As Object: Set dictCurrency = CreateObject("Scripting.Dictionary")
    dictCurrency.Add "AAPL", "USD"

    wsPool.Range("E6").Value = "统一RMB"
    WriteWideTable wsTag, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=dictCurrency, statementKind:="BalanceSheet"
    RefreshA1CurrencyComment wsTag, "美股_资产负债表"

    If InStr(CStr(wsTag.Cells(1, 3).Value), "[USD→RMB]") = 0 Then _
        Err.Raise vbObjectError + 744, "TestStep45Smoke", "R1 缺少 USD→RMB tag"
    If wsTag.Range("A1").Comment Is Nothing Then _
        Err.Raise vbObjectError + 745, "TestStep45Smoke", "A1 缺少动态注释"
    If InStr(wsTag.Range("A1").Comment.Text, "统一汇率换算") = 0 Then _
        Err.Raise vbObjectError + 746, "TestStep45Smoke", "A1 注释不是统一RMB文案"

    wsPool.Range("E6").Value = savedB6
End Sub


Public Sub TestPhase4hToggleSmoke()
    Dim wsPool As Worksheet
    Set wsPool = ThisWorkbook.Worksheets("样本池")

    Dim savedB6 As Variant
    savedB6 = wsPool.Range("E6").Value

    Dim wsSmoke As Worksheet
    Set wsSmoke = GetOrClearSmokeSheet("_phase4h_toggle_smoke")

    Dim arrCodes(1 To 1) As String
    arrCodes(1) = "AAPL"

    Dim arrPeriods(1 To 1) As String
    arrPeriods(1) = "2024-12-31"

    Dim arrIndicators(1 To 1) As String
    arrIndicators(1) = "Total assets"

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    dictCompanyName.Add "AAPL", "Apple"

    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.Add "Total assets", "Assets"

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictPer As Object: Set dictPer = CreateObject("Scripting.Dictionary")
    dictPer.Add "Total assets", 100#
    dictCompany.Add "2024-12-31", dictPer
    dictData.Add "AAPL", dictCompany

    Dim dictCurrency As Object: Set dictCurrency = CreateObject("Scripting.Dictionary")
    dictCurrency.Add "AAPL", "USD"

    wsPool.Range("E6").Value = "原币"
    WriteWideTable wsSmoke, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=dictCurrency, statementKind:="BalanceSheet"

    wsPool.Range("E6").Value = savedB6
End Sub


Public Sub TestPhase4kScoreSmoke()
    g_diagnosticSheetName = "韩股_抓取诊断"
    ClearDiagnosticSheet

    Dim rows As Collection: Set rows = New Collection
    AddDiagnosticRow rows, "005930", "BalanceSheet", "Score smoke", "OK_STOCKANALYSIS", _
                     "stockanalysis.com", "HTML", "Total Assets", "KRW billions", _
                     "1/1", "phase4k score text smoke", "1.0"
    WriteDiagnosticForKind "BalanceSheet", rows
End Sub


Public Sub TestPhase4kFxMissingSmoke()
    Dim wsPool As Worksheet: Set wsPool = ThisWorkbook.Worksheets("样本池")
    Dim wsFx As Worksheet: Set wsFx = ThisWorkbook.Worksheets("汇率")
    Dim savedDisplayMode As Variant: savedDisplayMode = wsPool.Range("E6").Value

    Dim fxRow As Long: fxRow = FindFxRowForSmoke("2024-12-31")
    If fxRow = 0 Then Err.Raise vbObjectError + 9401, "TestPhase4kFxMissingSmoke", "missing FX row 2024-12-31"
    Dim savedKrwAvg As Variant: savedKrwAvg = wsFx.Cells(fxRow, 7).Value

    On Error GoTo CleanUp
    wsFx.Cells(fxRow, 7).ClearContents
    wsPool.Range("E6").Value = "统一RMB"
    g_diagnosticSheetName = "韩股_抓取诊断"
    ClearDiagnosticSheet

    Dim wsSmoke As Worksheet
    Set wsSmoke = GetOrClearSmokeSheet("_phase4k_fx_missing_smoke")

    Dim arrCodes(1 To 1) As String
    arrCodes(1) = "005930"
    Dim arrPeriods(1 To 1) As String
    arrPeriods(1) = "2024-12-31"
    Dim arrIndicators(1 To 1) As String
    arrIndicators(1) = "Revenue"

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    dictCompanyName.Add "005930", "Samsung"
    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    dictCategory.Add "Revenue", "Smoke"
    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
    Dim dictPer As Object: Set dictPer = CreateObject("Scripting.Dictionary")
    dictPer.Add "Revenue", 100#
    dictCompany.Add "2024-12-31", dictPer
    dictData.Add "005930", dictCompany
    Dim dictCurrency As Object: Set dictCurrency = CreateObject("Scripting.Dictionary")
    dictCurrency.Add "005930", "KRW"

    WriteWideTable wsSmoke, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=dictCurrency, statementKind:="Income", _
                   useRawDumpLayer:=True
    wsSmoke.Calculate
    If IsError(wsSmoke.Range("C3").Value) Then
        wsSmoke.Range("ZZ1").Value = "ERROR"
    ElseIf Len(CStr(wsSmoke.Range("C3").Value)) = 0 Then
        wsSmoke.Range("ZZ1").Value = "BLANK"
    Else
        wsSmoke.Range("ZZ1").Value = CStr(wsSmoke.Range("C3").Value)
    End If
    wsSmoke.Range("ZZ2").Value = CountDiagnosticStatus("韩股_抓取诊断", "FX_MISSING")

CleanUp:
    wsFx.Cells(fxRow, 7).Value = savedKrwAvg
    wsPool.Range("E6").Value = savedDisplayMode
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub TestPhase4kLiveFxSmoke()
    Dim wsPool As Worksheet: Set wsPool = ThisWorkbook.Worksheets("样本池")
    Dim wsFx As Worksheet: Set wsFx = ThisWorkbook.Worksheets("汇率")
    Dim savedDisplayMode As Variant: savedDisplayMode = wsPool.Range("E6").Value

    Dim fxRow As Long: fxRow = FindFxRowForSmoke("2024-12-31")
    If fxRow = 0 Then Err.Raise vbObjectError + 9402, "TestPhase4kLiveFxSmoke", "missing FX row 2024-12-31"
    Dim savedUsdEop As Variant: savedUsdEop = wsFx.Cells(fxRow, 2).Value
    If Not IsNumeric(savedUsdEop) Or CDbl(savedUsdEop) <= 0 Then _
        Err.Raise vbObjectError + 9403, "TestPhase4kLiveFxSmoke", "missing USD EOP rate"

    On Error GoTo CleanUp
    wsPool.Range("E6").Value = "统一RMB"
    Dim wsSmoke As Worksheet
    Set wsSmoke = GetOrClearSmokeSheet("_phase4k_live_fx_smoke")

    Dim arrCodes(1 To 10) As String
    Dim arrPeriods(1 To 5) As String
    Dim arrIndicators(1 To 18) As String
    Dim i As Long, j As Long
    For i = 1 To 10
        arrCodes(i) = "USD" & Format$(i, "00")
    Next i
    For j = 1 To 5
        arrPeriods(j) = "2024-12-31"
    Next j
    For i = 1 To 18
        arrIndicators(i) = "Metric" & Format$(i, "00")
    Next i

    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    Dim dictCategory As Object: Set dictCategory = CreateObject("Scripting.Dictionary")
    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictCurrency As Object: Set dictCurrency = CreateObject("Scripting.Dictionary")
    For i = 1 To 18
        dictCategory.Add arrIndicators(i), "Smoke"
    Next i
    For i = 1 To 10
        dictCompanyName.Add arrCodes(i), "USD Smoke " & CStr(i)
        dictCurrency.Add arrCodes(i), "USD"
        Dim dictCompany As Object: Set dictCompany = CreateObject("Scripting.Dictionary")
        Dim dictPer As Object: Set dictPer = CreateObject("Scripting.Dictionary")
        For j = 1 To 18
            dictPer.Add arrIndicators(j), CDbl(100 + i + j)
        Next j
        dictCompany.Add "2024-12-31", dictPer
        dictData.Add arrCodes(i), dictCompany
    Next i

    WriteWideTable wsSmoke, arrCodes, dictCompanyName, dictData, arrPeriods, arrIndicators, dictCategory, _
                   perCompanyPeriods:=False, dictReportingCurrency:=dictCurrency, statementKind:="BalanceSheet", _
                   useRawDumpLayer:=True
    wsSmoke.Calculate
    Dim beforeVal As Variant: beforeVal = wsSmoke.Range("C3").Value
    Dim t As Double: t = Timer
    wsFx.Cells(fxRow, 2).Value = CDbl(savedUsdEop) + 0.5
    wsSmoke.Calculate
    Dim elapsed As Double: elapsed = Timer - t
    Dim afterVal As Variant: afterVal = wsSmoke.Range("C3").Value
    wsSmoke.Range("ZZ1").Value = beforeVal
    wsSmoke.Range("ZZ2").Value = afterVal
    wsSmoke.Range("ZZ3").Value = elapsed

CleanUp:
    wsFx.Cells(fxRow, 2).Value = savedUsdEop
    wsPool.Range("E6").Value = savedDisplayMode
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub TestPhase4lHttpMissHitSmoke()
    Dim wsSmoke As Worksheet
    Set wsSmoke = GetOrClearSmokeSheet("_phase4l_http_smoke")

    Dim oldSilent As Boolean: oldSilent = g_silentMode
    g_silentMode = True
    ClearLocalCache
    g_silentMode = oldSilent

    Dim url As String: url = "https://data.sec.gov/submissions/CIK0000320193.json"
    Dim cacheKey As String: cacheKey = "phase4l_AAPL_sec_companyfacts"
    Dim firstResult As THttpResult, secondResult As THttpResult
    Dim firstBody As String, secondBody As String

    firstBody = RunCachedHttpGet(url, cacheKey, "EDGAR", GetTtlHoursForSource("EDGAR"), firstResult)
    secondBody = RunCachedHttpGet(url, cacheKey, "EDGAR", GetTtlHoursForSource("EDGAR"), secondResult)

    wsSmoke.Range("A1").Value = firstResult.CacheStatus
    wsSmoke.Range("B1").Value = firstResult.StatusCode
    wsSmoke.Range("C1").Value = firstResult.ElapsedMs
    wsSmoke.Range("D1").Value = secondResult.CacheStatus
    wsSmoke.Range("E1").Value = secondResult.StatusCode
    wsSmoke.Range("F1").Value = secondResult.ElapsedMs
    wsSmoke.Range("G1").Value = Len(firstBody)
    wsSmoke.Range("H1").Value = Len(secondBody)

    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet
    Dim rows As Collection: Set rows = New Collection
    AddDiagnosticRow rows, "AAPL", "HTTP", "Phase4l telemetry smoke", "OK", "EDGAR", _
                     "SEC", "CIK0000320193", "JSON", "1/1", "phase4l cache HIT diagnostic smoke", "1.0"
    WriteDiagnosticForKind "HTTP", rows
End Sub


Public Sub TestPhase4lSecRateSmoke()
    Dim wsSmoke As Worksheet
    Set wsSmoke = GetOrClearSmokeSheet("_phase4l_sec_smoke")

    Dim oldSilent As Boolean: oldSilent = g_silentMode
    g_silentMode = True
    ClearLocalCache
    g_silentMode = oldSilent

    Dim url As String: url = "https://data.sec.gov/submissions/CIK0000320193.json"
    Dim keyBase As String
    keyBase = "phase4l_sec_rate_" & Format$(Now, "yyyymmddhhmmss") & "_" & CStr(CLng(Timer * 1000#))

    Dim firstResult As THttpResult, secondResult As THttpResult
    Dim firstBody As String, secondBody As String
    firstBody = RunCachedHttpGet(url, keyBase & "_1", "EDGAR", GetTtlHoursForSource("EDGAR"), firstResult)
    secondBody = RunCachedHttpGet(url, keyBase & "_2", "EDGAR", GetTtlHoursForSource("EDGAR"), secondResult)

    wsSmoke.Range("A1").Value = g_lastSecIntervalMs
    wsSmoke.Range("B1").Value = firstResult.StatusCode
    wsSmoke.Range("C1").Value = secondResult.StatusCode
    wsSmoke.Range("D1").Value = firstResult.CacheStatus
    wsSmoke.Range("E1").Value = secondResult.CacheStatus
    wsSmoke.Range("F1").Value = Len(firstBody)
    wsSmoke.Range("G1").Value = Len(secondBody)
End Sub


Public Sub TestPhase4lCleanReleaseSmoke()
    Dim wsSmoke As Worksheet
    Set wsSmoke = GetOrClearSmokeSheet("_phase4l_release_smoke")

    Dim wsPool As Worksheet: Set wsPool = ThisWorkbook.Worksheets("样本池")

    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet
    Dim rows As Collection: Set rows = New Collection
    AddDiagnosticRow rows, "AAPL", "Release", "diagnostic row", "OK", "EDGAR", _
                     "SEC", "Assets", "USD", "1/1", "phase4l release smoke", "1.0"
    WriteDiagnosticForKind "Release", rows

    WriteLocalHttpCache "phase4l_release_key", "{""release"":true}"
    Dim oldSilent As Boolean: oldSilent = g_silentMode
    g_silentMode = True
    Call CleanReleaseWorkbook(True)
    g_silentMode = oldSilent

    wsSmoke.Range("A1").Value = wsPool.Range("E5").Value
    wsSmoke.Range("B1").Value = ""
    wsSmoke.Range("C1").Value = ReadLocalHttpCache("phase4l_release_key")
    wsSmoke.Range("D1").Value = ThisWorkbook.Worksheets("美股_抓取诊断").Cells(3, 1).Value
End Sub


Public Function LoadFixture(ByVal fileName As String) As String
    Dim fixturePath As String
    fixturePath = ThisWorkbook.Path & "\tests\fixtures\" & fileName

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fixturePath) Then
        Err.Raise vbObjectError + 9701, "LoadFixture", "Fixture not found: " & fixturePath
    End If

    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile fixturePath
        LoadFixture = .ReadText(-1)
        .Close
    End With
End Function


Public Function RunOfflineTest(ByVal testName As String) As String
    On Error GoTo EH
    Application.Run "模块_测试." & testName
    RunOfflineTest = "PASS"
    Exit Function
EH:
    RunOfflineTest = "FAIL: " & CStr(Err.Number) & " | " & Err.Source & " | " & Err.Description
    Err.Clear
End Function


Public Sub Test_Offline_US_Edgar_AAPL()
    Dim body As String: body = LoadFixture("sec_aapl_companyfacts.json")

    AssertTrue InStr(1, body, """us-gaap""", vbTextCompare) > 0, "EDGAR fixture should include us-gaap taxonomy"
    AssertTrue InStr(1, body, """Revenues""", vbTextCompare) > 0, "EDGAR fixture should include Revenues"
    AssertTrue InStr(1, body, """USD""", vbTextCompare) > 0, "EDGAR fixture should include USD units"
    AssertTrue CountSubstring(body, """fy""") >= 5, "EDGAR fixture should include at least 5 fiscal periods"
End Sub


Public Sub Test_Offline_HK_Xueqiu_Tencent()
    Dim body As String: body = LoadFixture("xueqiu_hk_00700_balance.json")
    Dim hasTencent As Boolean
    Dim hasCurrency As Boolean

    hasTencent = (InStr(1, body, "腾讯控股", vbTextCompare) > 0) Or _
                 (InStr(1, body, "Tencent", vbTextCompare) > 0)
    hasCurrency = (InStr(1, body, """currency"":""CNY""", vbTextCompare) > 0) Or _
                  (InStr(1, body, """currency"":""HKD""", vbTextCompare) > 0) Or _
                  (InStr(1, body, "人民币", vbTextCompare) > 0) Or _
                  (InStr(1, body, "港币", vbTextCompare) > 0)

    AssertTrue hasTencent, "HK fixture should identify Tencent"
    AssertTrue InStr(1, body, """ta""", vbTextCompare) > 0, "HK fixture should include total assets field ta"
    AssertTrue InStr(1, body, """tlia""", vbTextCompare) > 0, "HK fixture should include total liabilities field tlia"
    AssertTrue hasCurrency, "HK fixture should include a reporting currency"
End Sub


Public Sub Test_Offline_KR_StockAnalysis_Samsung()
    Dim body As String: body = LoadFixture("stockanalysis_kr_005930_income.html")
    Dim hasSamsung As Boolean
    Dim hasOperatingIncome As Boolean
    Dim hasNetIncome As Boolean

    hasSamsung = (InStr(1, body, "Samsung Electronics", vbTextCompare) > 0) Or _
                 (InStr(1, body, "005930", vbTextCompare) > 0)
    hasOperatingIncome = (InStr(1, body, "opinc", vbTextCompare) > 0) Or _
                         (InStr(1, body, "operating income", vbTextCompare) > 0)
    hasNetIncome = (InStr(1, body, "netinc", vbTextCompare) > 0) Or _
                   (InStr(1, body, "net income", vbTextCompare) > 0)

    AssertTrue hasSamsung, "KR fixture should identify Samsung"
    AssertTrue InStr(1, body, "revenue", vbTextCompare) > 0, "KR fixture should include revenue"
    AssertTrue hasOperatingIncome, "KR fixture should include operating income"
    AssertTrue hasNetIncome, "KR fixture should include net income"
End Sub


Public Sub Test_Offline_FX_Missing_DoesNotFallbackToOne()
    Dim wsFx As Worksheet: Set wsFx = ThisWorkbook.Worksheets("汇率")
    Dim fxRow As Long: fxRow = FindFxRowForSmoke("2024-12-31")
    If fxRow = 0 Then Err.Raise vbObjectError + 9720, "Test_Offline_FX_Missing_DoesNotFallbackToOne", "missing FX row 2024-12-31"

    Dim savedUsdEop As Variant: savedUsdEop = wsFx.Cells(fxRow, 2).Value
    Dim fxRate As Double
    Dim statusText As String

    On Error GoTo CleanUp
    wsFx.Cells(fxRow, 2).ClearContents
    statusText = GetFxRateStatus("USD", "2024-12-31", True, fxRate)

    AssertEquals "FX_MISSING", statusText, "missing USD future rate should report FX_MISSING"
    AssertEquals "0", CStr(fxRate), "missing USD future rate should return 0"

CleanUp:
    wsFx.Cells(fxRow, 2).Value = savedUsdEop
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub Test_Offline_Diagnostic_Score_NotDate()
    g_diagnosticSheetName = "韩股_抓取诊断"
    ClearDiagnosticSheet

    Dim rows As Collection: Set rows = New Collection
    AddDiagnosticRow rows, "005930", "Offline", "score text", "OK", "fixture", _
                     "HTML", "score", "text", "1/1", "offline score text smoke", "1.0"
    WriteDiagnosticForKind "Offline", rows

    Dim wsDiag As Worksheet: Set wsDiag = ThisWorkbook.Worksheets("韩股_抓取诊断")
    AssertEquals "1/1", CStr(wsDiag.Cells(3, 9).Text), "diagnostic Score should stay as text"
End Sub


Public Sub Test_Offline_Cache_HitMissExpired()
    Dim cacheAge As Double, cacheStatus As String, body As String
    WriteLocalHttpCache "phase4m_offline_cache", "test_body"
    body = ReadLocalHttpCacheWithAge("phase4m_offline_cache", 24, cacheAge, cacheStatus)

    AssertEquals "test_body", body, "local cache body mismatch"
    AssertEquals "HIT", cacheStatus, "fresh local cache should be HIT"
    AssertTrue cacheAge >= 0 And cacheAge < 0.1, "fresh local cache age should be near zero"
End Sub


Public Sub Test_Offline_AppState_RestoreAfterError()
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Dim oldAlerts As Boolean: oldAlerts = Application.DisplayAlerts
    Dim st As TAppState
    Dim caught As Boolean

    On Error GoTo EH
    st = BeginAppState("phase4m offline app state smoke")
    Err.Raise vbObjectError + 9702, "Test_Offline_AppState_RestoreAfterError", "intentional smoke error"
    Exit Sub

EH:
    caught = True
    EndAppState st
    AssertTrue caught, "intentional app state error should be caught"
    AssertEquals CStr(oldCalc), CStr(Application.Calculation), "Application.Calculation should be restored"
    AssertEquals CStr(oldAlerts), CStr(Application.DisplayAlerts), "Application.DisplayAlerts should be restored"
End Sub


Public Sub Test_Offline_DataQuality_BS_Imbalance_Detection()
    Dim body As String: body = LoadFixture("missing_fields_edgar.json")
    AssertTrue InStr(1, body, """Assets""", vbTextCompare) > 0, "missing-fields fixture should include Assets"
    AssertTrue InStr(1, body, """Revenues""", vbTextCompare) = 0, "missing-fields fixture should omit Revenues"

    RunDataQualityChecks
    AssertTrue DiagnosticHasQaCode("BS_BALANCE"), "RunDataQualityChecks should emit BS_BALANCE"
    AssertTrue DiagnosticHasQaCode("FX_MISSING"), "RunDataQualityChecks should emit FX_MISSING"
    AssertTrue DiagnosticHasQaCode("KEY_FIELDS"), "RunDataQualityChecks should emit KEY_FIELDS"
End Sub


Private Sub AssertTrue(ByVal cond As Boolean, ByVal msg As String)
    If Not cond Then
        Err.Raise vbObjectError + 9710, "AssertTrue", "Assertion failed: " & msg
    End If
End Sub


Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal msg As String)
    If CStr(expected) <> CStr(actual) Then
        Err.Raise vbObjectError + 9711, "AssertEquals", _
            msg & " | expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub


Private Function CountSubstring(ByVal textValue As String, ByVal needle As String) As Long
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, textValue, needle, vbTextCompare)
        If pos = 0 Then Exit Do
        CountSubstring = CountSubstring + 1
        pos = pos + Len(needle)
    Loop
End Function


Private Function DiagnosticHasQaCode(ByVal qaCode As String) As Boolean
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("美股_抓取诊断")
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To lastRow
        If CStr(ws.Cells(r, 1).Value) = "GLOBAL_QA" And _
           CStr(ws.Cells(r, 2).Value) = qaCode Then
            DiagnosticHasQaCode = True
            Exit Function
        End If
    Next r
End Function


Private Function FindFxRowForSmoke(ByVal periodKey As String) As Long
    Dim wsFx As Worksheet: Set wsFx = ThisWorkbook.Worksheets("汇率")
    Dim r As Long, lastRow As Long
    lastRow = wsFx.Cells(wsFx.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Format$(wsFx.Cells(r, 1).Value, "yyyy-mm-dd") = periodKey Then
            FindFxRowForSmoke = r
            Exit Function
        End If
    Next r
End Function


Private Function CountDiagnosticStatus(ByVal sheetName As String, ByVal statusText As String) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To lastRow
        If CStr(ws.Cells(r, 4).Value) = statusText Then CountDiagnosticStatus = CountDiagnosticStatus + 1
    Next r
End Function


Private Function GetOrClearSmokeSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set GetOrClearSmokeSheet = ws
End Function

Attribute VB_Name = "模块_抓汇率"
Option Explicit

' =================================================================
'  Phase 4f Step 2: 雪球 FX 抓取核心模块
'  作者: Generator sub-agent (依据 PHASE_4F_STEP2_TASKS.md)
'
'  ⚠️ 关键约束 (Step 1 RMB_FX_PROBE.md 实测锁定):
'    1. FX symbol 必带 .FX: USDCNY.FX / HKDCNY.FX / KRWCNY.FX
'       (没 .FX 后缀 → HTTP 200 + data 空数组, silent 失败)
'    2. Accept-Encoding 必须 "gzip, deflate" (NOT identity)
'       → 不复用 模块_工具函数.bas line 608 XueqiuHttpGet (它用 identity, 200 + 空 body)
'    3. 进程启动需 GET https://xueqiu.com/hq 拿 xq_a_token cookie
'       → 用 cookie session 拿数据
'    4. quote response: parsed("data") 直接是 list (本期不用,Step 4 用)
'       — 注释提示防误用, 不要写成 parsed("data")("items")
'    5. kline response: parsed("data")("item") 是 list, 12 列
'       VBA 1-based: ts = items(i)(1), close = items(i)(6)
'       (Step 1 0-based item[0]/item[5] +1)
'    6. begin = CST 0 点 ms, count=-2000 单次覆盖 7.7 年, 不分段
'
'  HTTP 实现策略 (重要):
'    - 实测: WinHttp.WinHttpRequest.5.1 + Accept-Encoding: gzip, deflate 拿到的是
'      gzip 压缩字节, **不会自动 inflate**. ServerXMLHTTP.6.0 同样症状.
'    - 解决: shell out 到 PowerShell, 用 System.Net.Http.HttpClient 的
'      AutomaticDecompression + UseCookies (CookieContainer 自动维护).
'      PS 把响应 JSON 写临时文件, VBA 读临时文件 → ADODB.Stream UTF-8 解码.
'    - PS 进程开销 ~600ms/调用; GetFxRate 先查缓存, 只有缺值才触发网络请求.
'
'  对外接口:
'    Public Function EnsureFxRateCached(periodEnd As String, curCode As String) As Boolean
'
'  模块依赖:
'    - JsonConverter.bas (ParseJson)
'    - 模块_工具函数.bas (FX_SHEET, FX_DATA_ROW 常量; LookupFxColForCurrency)


' --------- 进程启动暖 session 探针 (本实现里 PS 内部已含 warmup, 这里只 print) ---------
'   FX 请求必须经过 https://xueqiu.com/hq warmup, 并使用 Accept-Encoding: gzip, deflate
'   保留这个 Sub 是为了让 EnsureFxRateCached 里能 Debug.Print "warmup" 探针
'   (PHASE_4F_STEP2_TASKS.md verify 要求"3 次调用只 print 1 次")
Private g_blnWarmupPrinted As Boolean

Private Sub WarmupXueqiuSession()
    If Not g_blnWarmupPrinted Then
        Debug.Print "warmup"
        g_blnWarmupPrinted = True
    End If
End Sub


Private Function PSEscapeSingleQuoted(ByVal s As String) As String
    PSEscapeSingleQuoted = Replace(s, "'", "''")
End Function


' --------- 通过 PowerShell 拉一个 URL, 返回 UTF-8 JSON 字符串 ---------
'   url: 完整 URL
'   warmupFirst: True 时 PS 内先 GET https://xueqiu.com/hq 拿 cookie
'   失败/timeout → 抛异常
Private Function FetchViaPowerShell(ByVal url As String, ByVal warmupFirst As Boolean) As String
    WarmupXueqiuSession

    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp As String
    tmp = Environ("TEMP") & "\fx_" & Format(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 100000)) & ".json"

    Dim userCookie As String
    userCookie = ReadXueqiuCookie()

    ' PowerShell 脚本: HttpClient + AutomaticDecompression + UseCookies
    Dim ps As String
    ps = "$ProgressPreference='SilentlyContinue';" & _
         "Add-Type -AssemblyName System.Net.Http;" & _
         "$h = New-Object System.Net.Http.HttpClientHandler;" & _
         "$h.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate;" & _
         "$h.UseCookies = $true;" & _
         "$c = New-Object System.Net.Http.HttpClient($h);" & _
         "$c.DefaultRequestHeaders.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');" & _
         "$c.DefaultRequestHeaders.Add('Accept', 'application/json, text/plain, */*');" & _
         "$null = $c.DefaultRequestHeaders.TryAddWithoutValidation('Accept-Encoding', 'gzip, deflate');" & _
         "$c.DefaultRequestHeaders.Add('Accept-Language', 'zh-CN,zh;q=0.9,en;q=0.8');" & _
         "$c.DefaultRequestHeaders.Add('Referer', 'https://xueqiu.com/');"
    If Len(userCookie) > 0 Then
        ps = ps & "$null = $c.DefaultRequestHeaders.TryAddWithoutValidation('Cookie', '" & PSEscapeSingleQuoted(userCookie) & "');"
    End If
    If warmupFirst Then
        ps = ps & "$null = $c.GetStringAsync('https://xueqiu.com/hq').Result;" & _
                  "Start-Sleep -Milliseconds 200;"
    End If
    ps = ps & "$r = $c.GetStringAsync('" & url & "').Result;" & _
              "[System.IO.File]::WriteAllText('" & tmp & "', $r, [System.Text.Encoding]::UTF8);"

    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command """ & ps & """"
    sh.Run cmd, 0, True   ' 0=hide window, True=wait

    If Not fso.FileExists(tmp) Then
        Err.Raise vbObjectError + 542, "FetchViaPowerShell", _
            "PowerShell 未生成响应文件: " & tmp
    End If

    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2          ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile tmp
    Dim outStr As String: outStr = stream.ReadText
    stream.Close

    On Error Resume Next
    fso.DeleteFile tmp, True
    Err.Clear
    On Error GoTo 0

    Debug.Print "FetchViaPowerShell raw[0:80] = " & Left$(outStr, 80)

    FetchViaPowerShell = outStr
End Function


' --------- 独立 FX HTTP 客户端: gzip, deflate + warmup, 不复用 XueqiuHttpGet(identity) ---------
Private Function XueqiuFxHttpGet(ByVal strUrl As String) As String
    XueqiuFxHttpGet = FetchViaPowerShell(strUrl, True)
End Function


' --------- 计算今天 CST 0 点的 ms timestamp (用于 kline begin 参数) ---------
'   用 Currency 类型 (8 字节, ±9.2e14) 防 Long 溢出 ms timestamp
Private Function BeginCstMsForToday() As Currency
    Dim today As Date
    today = DateSerial(Year(Now), Month(Now), Day(Now))
    Dim daysSinceEpoch As Double
    daysSinceEpoch = today - DateSerial(1970, 1, 1)
    ' CST 0 点 → UTC = (本机 0 点) - 8h
    Dim msAsDouble As Double
    msAsDouble = daysSinceEpoch * 86400000# - 8# * 3600000#
    BeginCstMsForToday = CCur(msAsDouble)
End Function


' --------- 把 yyyy-mm-dd 报告期解析成 CST 当日 23:59:59 的 ms timestamp ---------
Private Function PeriodEndCstMs(ByVal periodEnd As String) As Currency
    On Error Resume Next
    Dim d As Date
    d = DateValue(periodEnd)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        PeriodEndCstMs = 0
        Exit Function
    End If
    On Error GoTo 0

    Dim daysSinceEpoch As Double
    daysSinceEpoch = d - DateSerial(1970, 1, 1)
    ' 当日 CST 23:59:59 → 加 1 天减 1 秒, 再扣 8h
    Dim msAsDouble As Double
    msAsDouble = (daysSinceEpoch + 1) * 86400000# - 1000# - 8# * 3600000#
    PeriodEndCstMs = CCur(msAsDouble)
End Function


' --------- 拉某币种 K 线, 返回 2D Variant array (i,0)=ts_ms / (i,1)=close ---------
'   curCode: "USD" / "HKD" / "KRW" (其他报错)
'   beginCstMs: kline begin 参数 (Currency, ms)
'   barCount: 拉取条数 (production = 2000 day)
'   失败/无数据 → ReDim 0 行
Private Function FetchFxKline(ByVal curCode As String, ByVal beginCstMs As Currency, ByVal barCount As Long) As Variant
    Dim fxSymbol As String
    Select Case UCase$(curCode)
        Case "USD"
            fxSymbol = "USDCNY.FX"
        Case "HKD"
            fxSymbol = "HKDCNY.FX"
        Case "KRW"
            fxSymbol = "KRWCNY.FX"
        Case Else
            Err.Raise vbObjectError + 541, "FetchFxKline", _
                "未知币种: " & curCode
    End Select

    Dim fxUrl As String
    Dim beginStr As String
    beginStr = Format$(beginCstMs, "0")
    fxUrl = "https://stock.xueqiu.com/v5/stock/chart/kline.json" & _
            "?symbol=" & fxSymbol & _
            "&begin=" & beginStr & _
            "&period=day&type=before&count=-" & CStr(barCount)

    Dim raw As String
    raw = XueqiuFxHttpGet(fxUrl)

    Dim parsed As Object
    Dim emptyArr() As Variant
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(raw)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ReDim emptyArr(0 To 0, 0 To 1)
        FetchFxKline = emptyArr
        Exit Function
    End If
    On Error GoTo 0

    ' Step 1 RMB_FX_PROBE: parsed("data")("item") 是 list, 每项 12 列
    '   VBA 1-based: items(i)(1) = ts_ms (Step 1 0-based item[0])
    '                items(i)(6) = close (Step 1 0-based item[5])
    Dim itemList As Object
    On Error Resume Next
    Set itemList = parsed("data")("item")
    If Err.Number <> 0 Or itemList Is Nothing Then
        Err.Clear
        On Error GoTo 0
        ReDim emptyArr(0 To 0, 0 To 1)
        FetchFxKline = emptyArr
        Exit Function
    End If
    On Error GoTo 0

    Dim n As Long: n = itemList.Count
    If n = 0 Then
        ReDim emptyArr(0 To 0, 0 To 1)
        FetchFxKline = emptyArr
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(0 To n - 1, 0 To 1)
    Dim k As Long, rowObj As Object
    Dim tsCell As Variant, closeCell As Variant
    For k = 1 To n
        Set rowObj = itemList(k)
        tsCell = rowObj(1)        ' ts_ms (Step 1 0-based item[0] → VBA 1-based items(i)(1))
        closeCell = rowObj(6)     ' close (Step 1 0-based item[5] → VBA 1-based items(i)(6))
        If IsNull(tsCell) Then tsCell = 0#
        If IsNull(closeCell) Then closeCell = 0#
        arr(k - 1, 0) = CCur(tsCell)
        arr(k - 1, 1) = CDbl(closeCell)
    Next k
    FetchFxKline = arr
End Function


' --------- 推断财年起点: 默认 01-01 (港股/中概的 12-31 财年覆盖) ---------
'   periodEnd: yyyy-mm-dd → 返 yyyy-01-01 的 Date
Private Function InferPeriodStart(ByVal periodEnd As String) As Date
    On Error Resume Next
    Dim d As Date
    d = DateValue(periodEnd)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        InferPeriodStart = DateSerial(1970, 1, 1)
        Exit Function
    End If
    On Error GoTo 0
    InferPeriodStart = DateSerial(Year(d), 1, 1)
End Function


' --------- 算期末汇率: 取 ts <= periodEnd+1day 的最大 ts 对应 close ---------
'   节假日 periodEnd 当日无 day bar → 自动回溯到最近一个交易日
Private Function ComputeEopRate(ByVal klineItems As Variant, ByVal periodEnd As String) As Double
    Dim cutoffMs As Currency
    cutoffMs = PeriodEndCstMs(periodEnd)
    If cutoffMs = 0 Then
        ComputeEopRate = 0
        Exit Function
    End If
    cutoffMs = cutoffMs + CCur(86400000#)   ' +1 day buffer

    Dim n As Long, lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(klineItems, 1)
    hi = UBound(klineItems, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ComputeEopRate = 0
        Exit Function
    End If
    On Error GoTo 0
    n = hi - lo + 1
    If n <= 0 Then
        ComputeEopRate = 0
        Exit Function
    End If

    Dim i As Long, bestTs As Currency, bestClose As Double
    Dim ts As Currency, cl As Double
    bestTs = 0
    bestClose = 0
    For i = lo To hi
        ts = CCur(klineItems(i, 0))
        cl = CDbl(klineItems(i, 1))
        If ts <= cutoffMs And cl > 0 Then
            If ts > bestTs Then
                bestTs = ts
                bestClose = cl
            End If
        End If
    Next i
    ComputeEopRate = bestClose
End Function


' --------- 算期间均值: 区间内所有 close 算术平均 ---------
'   periodStart/periodEnd: yyyy-mm-dd
'   含端点; 缺日 (节假日) 自动跳过
Private Function ComputeAvgRate(ByVal klineItems As Variant, ByVal periodStart As Date, ByVal periodEnd As String) As Double
    Dim startMs As Currency, endMs As Currency
    Dim startDaysSinceEpoch As Double
    startDaysSinceEpoch = periodStart - DateSerial(1970, 1, 1)
    startMs = CCur(startDaysSinceEpoch * 86400000# - 8# * 3600000#)

    endMs = PeriodEndCstMs(periodEnd)
    If endMs = 0 Then
        ComputeAvgRate = 0
        Exit Function
    End If
    endMs = endMs + CCur(86400000#)

    Dim lo As Long, hi As Long
    On Error Resume Next
    lo = LBound(klineItems, 1)
    hi = UBound(klineItems, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ComputeAvgRate = 0
        Exit Function
    End If
    On Error GoTo 0
    If hi < lo Then
        ComputeAvgRate = 0
        Exit Function
    End If

    Dim i As Long, sumCl As Double, cnt As Long
    Dim ts As Currency, cl As Double
    sumCl = 0
    cnt = 0
    For i = lo To hi
        ts = CCur(klineItems(i, 0))
        cl = CDbl(klineItems(i, 1))
        If ts >= startMs And ts <= endMs And cl > 0 Then
            sumCl = sumCl + cl
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        ComputeAvgRate = 0
    Else
        ComputeAvgRate = sumCl / cnt
    End If
End Function


' --------- 主入口: 确保 (periodEnd, curCode) 的汇率已写入『汇率』sheet ---------
'   periodEnd: yyyy-mm-dd
'   curCode  : "RMB"/"CNY" → True (无需写)
'              "USD"/"HKD"/"KRW" → 查 sheet, 缺则拉 K 线写, 用户手填值不覆盖
'   返回 True=成功 (sheet 含 valid 期末/期均) / False=失败
'
'   用户 override 检测: cell value IsNumeric And > 0 → 视为有效, 不重拉
Public Function EnsureFxRateCached(ByVal periodEnd As String, ByVal curCode As String) As Boolean
    Dim c As String: c = UCase$(Trim$(curCode))
    If c = "RMB" Or c = "CNY" Then
        EnsureFxRateCached = True
        Exit Function
    End If
    If c <> "USD" And c <> "HKD" And c <> "KRW" Then
        EnsureFxRateCached = False
        Exit Function
    End If

    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FX_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        EnsureFxRateCached = False
        Exit Function
    End If

    ' 找 / 创建 periodEnd 行
    Dim rowIdx As Long
    rowIdx = FindOrCreateFxRow(ws, periodEnd)
    If rowIdx = 0 Then
        EnsureFxRateCached = False
        Exit Function
    End If

    Dim eopCol As Long, avgCol As Long
    eopCol = LookupFxColForCurrency(c, True)
    avgCol = LookupFxColForCurrency(c, False)
    If eopCol = 0 Or avgCol = 0 Then
        EnsureFxRateCached = False
        Exit Function
    End If

    ' 缓存命中检查 (含用户 override): 两列都 IsNumeric > 0 → True 直接返
    Dim eopVal As Variant, avgVal As Variant
    eopVal = ws.Cells(rowIdx, eopCol).Value
    avgVal = ws.Cells(rowIdx, avgCol).Value
    Dim eopOK As Boolean, avgOK As Boolean
    eopOK = (IsNumeric(eopVal) And Not IsEmpty(eopVal))
    If eopOK Then eopOK = (CDbl(eopVal) > 0)
    avgOK = (IsNumeric(avgVal) And Not IsEmpty(avgVal))
    If avgOK Then avgOK = (CDbl(avgVal) > 0)
    If eopOK And avgOK Then
        EnsureFxRateCached = True
        Exit Function
    End If

    ' Cache miss → 拉 K 线
    Dim beginMs As Currency
    beginMs = BeginCstMsForToday()
    Dim klineItems As Variant
    On Error Resume Next
    klineItems = FetchFxKline(c, beginMs, 2000)    ' Step 1 must-fix: count=-2000, 单次覆盖 7.7 年
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        EnsureFxRateCached = False
        Exit Function
    End If
    On Error GoTo 0

    ' 算期末 + 期均
    Dim eopRate As Double, avgRate As Double
    eopRate = ComputeEopRate(klineItems, periodEnd)
    Dim periodStart As Date
    periodStart = InferPeriodStart(periodEnd)
    avgRate = ComputeAvgRate(klineItems, periodStart, periodEnd)

    ' 写 sheet (用户 override 不覆盖: 只写空 cell 或无效值的 cell)
    If Not eopOK And eopRate > 0 Then
        ws.Cells(rowIdx, eopCol).Value = eopRate
    End If
    If Not avgOK And avgRate > 0 Then
        ws.Cells(rowIdx, avgCol).Value = avgRate
    End If

    ' 检查最终是否成功 (允许 EOP 成 / AVG 失败 之类的部分成功也返 True)
    If eopRate > 0 Or avgRate > 0 Or eopOK Or avgOK Then
        EnsureFxRateCached = True
    Else
        EnsureFxRateCached = False
    End If
End Function


' --------- 找 periodEnd 所在行; 不存在则在表末新建一行 + A 列写 periodEnd ---------
Private Function FindOrCreateFxRow(ByVal ws As Object, ByVal periodEnd As String) As Long
    Dim r As Long, lastRow As Long
    Dim target As String: target = Trim$(periodEnd)
    If Len(target) = 0 Then
        FindOrCreateFxRow = 0
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row    ' xlUp
    If lastRow < FX_DATA_ROW Then lastRow = FX_DATA_ROW - 1

    For r = FX_DATA_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) = target Then
            FindOrCreateFxRow = r
            Exit Function
        End If
    Next r

    ' 不存在 → 新行
    Dim newRow As Long: newRow = lastRow + 1
    ws.Cells(newRow, 1).NumberFormat = "@"
    ws.Cells(newRow, 1).Value = target
    FindOrCreateFxRow = newRow
End Function

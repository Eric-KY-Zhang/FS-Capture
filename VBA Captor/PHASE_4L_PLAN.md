# Phase 4l: 优化 Sprint 2 — HTTP/cache 诊断遥测 + 重试退避 + 发布清理

> **版本**: v2(2026-05-05,Phase 4l 优化 Sprint 2 闭环)
> **状态**: ✅ Phase 4l 全期闭环
> **作者**: Claude(planner) + Codex(generator)
> **背景**: Phase 4k 已修完 4 项数据准确性 + UX 大 bug(FX missing / live FX UDF / AppStateGuard / KR Score)。Phase 4l 继续优化 backlog 第二批 — 把"看不见的网络层 + 缓存层"变得可观测,加 retry / 限流防偶发失败,发布清理宏让 xlsm 可安全分享。

## 项目语境(给 Generator 的 anchor 段)

本期 3 项工作**全部在现有数据流内**,不引入新数据源 / 新抓数频次。Step 1 (P1-01) 给 HTTP / cache 函数返回结构化结果 + 诊断 sheet 加新列,Step 2 (P1-03) 给 HTTP 请求加重试退避(**减少**偶发失败而非新增请求),Step 3 (P0-01) 加 1 个独立的发布清理宏(用户分享 xlsm 前主动调,不影响日常使用)。

## 3 项任务清单

| # | Task | 严重性 | 工作量 |
|---|---|---|---|
| 1 | **P1-01** HTTP/cache 诊断遥测 | UX 大改进(用户能看清每次抓数走的是网络还是缓存)| 3h |
| 2 | **P1-03** HTTP 重试退避 + SEC 限流 | 稳定性(应对 429/5xx 偶发失败)| 2h |
| 3 | **P0-01** CleanReleaseWorkbook 宏 | 隐私安全(分享 xlsm 前清 cookie/缓存/元数据)| 1h |

总耗时 ~6h Codex,~1h Reviewer。

## Step 总览

| Step | 内容 | 估时 | 阻塞依赖 |
|---|---|---|---|
| 1 | P1-01:`THttpResult` struct + cache wrapper 返回结构化状态 + 诊断 sheet 扩展列 | 3h | 无 |
| 2 | P1-03:`HttpGetWithRetry` helper(指数退避 + 状态码白/黑名单 + SEC ≤10/s 限流)| 2h | 与 Step 1 共享 THttpResult 结构 |
| 3 | P0-01:`CleanReleaseWorkbook` 宏(清 B5 cookie + 诊断历史 + 元数据)| 1h | 无(独立)|
| 4 | 端到端回归 4 张 + 新增 `inspect_phase4l_state.py` + STATUS §CC 收口 | 0.5h | 1-3 |

**Codex 工作流建议**:**1 round 端到端**(跟 Phase 4j/4k 一致),commit `Phase 4l optimization sprint 2 (HTTP telemetry + retry/backoff + release sanitizer)`。

---

## Step 1 — P1-01 HTTP/cache 诊断遥测

### 背景

当前 cache wrapper(`CachedEdgarHttpGet / CachedXueqiuHttpGet / StockAnalysisUSHttpGet 等`)只返回 body 字符串。用户无法判断:
- 这次返回是从网络抓的还是 24h cache 命中?
- 缓存有多旧?
- 抓数耗时多少?
- HTTP 实际状态码是什么?
- 重试了几次?

→ 加结构化遥测,把这些写进诊断 sheet。

### 实现

**子任务 1A — `模块_工具函数.bas` 新增 `THttpResult` struct**:

```vba
' Phase 4l Step 1: HTTP / cache 调用结构化返回
Public Type THttpResult
    Body As String              ' 响应体, 失败时空字符串
    StatusCode As Long          ' HTTP 状态码 (200 / 404 / 429 / etc.); cache hit 时 = 0
    StatusText As String        ' "OK" / "Not Found" / 自定义错误描述
    Source As String            ' "EDGAR" / "XUEQIU" / "STOCKANALYSIS_KR" / etc.
    UrlHash As String           ' SHA1(URL) 截断 12 字符, 用于诊断关联
    CacheKey As String          ' cache 文件名 (不含路径)
    CacheStatus As String       ' "HIT" / "MISS" / "EXPIRED" / "BYPASS" / "WRITE_ERROR" / "READ_ERROR"
    CacheAgeHours As Double     ' cache hit 时 = 缓存年龄; miss/expired 时 = -1
    ElapsedMs As Long           ' 整次调用耗时(毫秒, 含 cache 读 + HTTP + 写)
    RetryCount As Long          ' Step 2 retry 次数; 本 Step 默认 0
    ErrorStage As String        ' "" / "CACHE_READ" / "HTTP" / "PARSE" / "CACHE_WRITE"
    ErrorText As String         ' 异常时的错误摘要 (截断 200 字符)
End Type
```

**子任务 1B — Cache wrapper 改造(只改 wrapper,核心 HTTP 不动)**:

`ReadLocalHttpCache` / `WriteLocalHttpCache` 不动。**新增** `RunCachedHttpGet(url, cacheKey, source, ttlHours, ByRef result As THttpResult) As String`:

```vba
Public Function RunCachedHttpGet(ByVal url As String, _
                                  ByVal cacheKey As String, _
                                  ByVal source As String, _
                                  ByVal ttlHours As Long, _
                                  ByRef result As THttpResult) As String
    Dim startTime As Double: startTime = Timer
    result.Source = source
    result.CacheKey = cacheKey
    result.UrlHash = ComputeShortHash(url)
    
    ' 1. Try cache
    Dim cacheBody As String, cacheAge As Double
    cacheBody = ReadLocalHttpCacheWithAge(cacheKey, ttlHours, cacheAge)
    If Len(cacheBody) > 0 Then
        result.Body = cacheBody
        result.CacheStatus = IIf(cacheAge >= 0 And cacheAge <= ttlHours, "HIT", "EXPIRED")
        result.CacheAgeHours = cacheAge
        result.StatusCode = 0
        result.StatusText = "CACHE_" & result.CacheStatus
        result.ElapsedMs = CLng((Timer - startTime) * 1000)
        RunCachedHttpGet = cacheBody
        Exit Function
    End If
    
    ' 2. Cache miss → real HTTP
    result.CacheStatus = "MISS"
    result.CacheAgeHours = -1
    
    Dim httpStart As Double: httpStart = Timer
    On Error Resume Next
    Dim body As String, statusCode As Long
    ' ... 调实际 HTTP (按 source 分发: EdgarHttpGet / XueqiuHttpGet / StockAnalysisHttpGet)
    Select Case source
        Case "EDGAR":            body = EdgarHttpGet(url, statusCode)
        Case "XUEQIU":           body = XueqiuHttpGet(url, statusCode)
        Case "STOCKANALYSIS_KR", "STOCKANALYSIS_US": body = StockAnalysisHttpGet(url, statusCode)
        Case Else:               body = ""
    End Select
    Dim httpErr As Long: httpErr = Err.Number
    Dim httpErrDesc As String: httpErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    
    result.StatusCode = statusCode
    If httpErr <> 0 Then
        result.ErrorStage = "HTTP"
        result.ErrorText = Left$(httpErrDesc, 200)
        result.StatusText = "HTTP_ERROR"
        result.Body = ""
        result.ElapsedMs = CLng((Timer - startTime) * 1000)
        RunCachedHttpGet = ""
        Exit Function
    End If
    
    result.Body = body
    result.StatusText = "OK"
    
    ' 3. Write cache (only for HTTP 200)
    If statusCode = 200 And Len(body) > 0 Then
        On Error Resume Next
        WriteLocalHttpCache cacheKey, body
        If Err.Number <> 0 Then
            result.CacheStatus = "WRITE_ERROR"
            result.ErrorStage = "CACHE_WRITE"
            result.ErrorText = Left$(Err.Description, 200)
        End If
        Err.Clear
        On Error GoTo 0
    End If
    
    result.ElapsedMs = CLng((Timer - startTime) * 1000)
    RunCachedHttpGet = body
End Function
```

**子任务 1C — 现有 cache wrapper 改用 RunCachedHttpGet**:

`CachedEdgarHttpGet / CachedXueqiuHttpGet / 等` 这些 wrapper **保留旧签名**(返回 String body),内部调 `RunCachedHttpGet`,把 result 写到一个 module-level 的 `LastHttpResult` 全局变量。fetch 模块需要诊断信息时去读 `LastHttpResult`。

或更优雅:fetch 模块直接调 `RunCachedHttpGet` 拿 result + body,然后在写诊断时一起 emit。

**子任务 1D — `ReadLocalHttpCacheWithAge` helper 新增**:

`ReadLocalHttpCache` 现在只返回 body。新增同名 + Age 后缀版,同时返回 cache 年龄(浮点 hours)。

**子任务 1E — 诊断 sheet 列扩展**:

诊断 sheet 当前 11 列(R2 表头:公司 / 报表 / 输出指标 / 状态 / 数据源 / Taxonomy / 命中字段 / Unit / Score / 匹配方式+备注 / FX_Rate)。

**Phase 4l 扩到 17 列**(追加 6 列):

| 列 | 新列名 | 内容 |
|---|---|---|
| L | CacheStatus | HIT / MISS / EXPIRED / WRITE_ERROR |
| M | CacheAgeHours | cache hit 时填年龄, miss/N/A 留空 |
| N | HTTPStatus | 200 / 404 / 429 / 0(cache hit)|
| O | ElapsedMs | 整次调用耗时(毫秒)|
| P | RetryCount | Step 2 retry 次数, 本 Step 默认 0 |
| Q | ErrorStage | "" / "CACHE_READ" / "HTTP" / "PARSE" / "CACHE_WRITE" |

`tools/install_modules.py` 的 `_make_diagnostic_sheet` + `_refresh_diagnostic_headers` 同步:
- `headers` 列表加 6 项
- `widths` 加对应列宽 [12, 10, 10, 10, 8, 14]
- L:Q 列文本格式 `@`(避免 Excel 把 `200` 误转日期)

`AddDiagnosticRow` helper(grep 一下 现有 ` arr(r, 1) = ...` 模式)扩展 6 个新参数 + 写入新列。

### 验证

- 跑 `一键港股`(00700)2 次:
  - 第 1 次:诊断 sheet `CacheStatus = MISS`, `HTTPStatus = 200`, `ElapsedMs ~ 800-2000`
  - 第 2 次:`CacheStatus = HIT`, `HTTPStatus = 0`, `ElapsedMs < 50`
- 清空 `.cache/` 后再跑:回到 `MISS`
- 故意把 cache 文件 mtime 改成 25 小时前:`CacheStatus = EXPIRED`(注:这步可以 manual 验证或 skip)

### Generator 不要做

- ❌ 不要改核心 HTTP 函数签名(`EdgarHttpGet / XueqiuHttpGet / StockAnalysisHttpGet`)— 只在 wrapper 层加遥测
- ❌ 不要改诊断 sheet 现有 11 列内容(只追加 6 列)
- ❌ 不要把 cache key / URL 原文写诊断(用 hash)— 隐私 + 列宽控制

---

## Step 2 — P1-03 HTTP 重试退避 + SEC 限流

### 背景

当前 HTTP 请求一次失败就报错。SEC EDGAR / 雪球 / stockanalysis 偶发 429 / 503 / 5xx 应该重试。SEC 官方要求 ≤10 req/sec(fair access),需要独立限流。

### 实现

**子任务 2A — `HttpGetWithRetry` helper**(放在 Step 1 cache wrapper 内部 / 之前):

```vba
' Phase 4l Step 2: HTTP 请求带重试 / 退避
'   重试条件: 408 / 429 / 500 / 502 / 503 / 504
'   不重试: 400 / 401 / 403 / 404 (永久失败)
'   退避: 第 1 次 500ms, 第 2 次 1000ms, 第 3 次 2000ms + 0-300ms jitter
Public Function HttpGetWithRetry(ByVal url As String, _
                                  ByVal source As String, _
                                  ByRef result As THttpResult) As String
    Dim retryDelays As Variant: retryDelays = Array(500, 1000, 2000)
    Dim maxRetries As Long: maxRetries = UBound(retryDelays) + 1
    Dim attempt As Long
    Dim body As String, statusCode As Long
    
    For attempt = 0 To maxRetries
        ' SEC 限流: ≤10 req/sec → 至少 100ms 间隔
        If source = "EDGAR" Then EnforceSecRateLimit
        
        On Error Resume Next
        Select Case source
            Case "EDGAR":            body = EdgarHttpGet(url, statusCode)
            Case "XUEQIU":           body = XueqiuHttpGet(url, statusCode)
            Case "STOCKANALYSIS_KR", "STOCKANALYSIS_US": body = StockAnalysisHttpGet(url, statusCode)
            Case Else:               body = ""
        End Select
        Dim errNum As Long: errNum = Err.Number
        Err.Clear
        On Error GoTo 0
        
        result.StatusCode = statusCode
        result.RetryCount = attempt
        
        ' 成功
        If errNum = 0 And statusCode = 200 Then
            result.StatusText = "OK"
            HttpGetWithRetry = body
            Exit Function
        End If
        
        ' 永久失败(不重试)
        Select Case statusCode
            Case 400, 401, 403, 404
                result.StatusText = "HTTP_" & statusCode & "_NO_RETRY"
                result.ErrorStage = "HTTP"
                HttpGetWithRetry = ""
                Exit Function
        End Select
        
        ' 可重试错误(继续 loop)
        If attempt < maxRetries Then
            Dim delayMs As Long: delayMs = CLng(retryDelays(attempt)) + Int(Rnd() * 300)
            Sleep delayMs
        End If
    Next attempt
    
    ' 所有重试用完
    result.StatusText = "RETRY_EXHAUSTED"
    result.ErrorStage = "HTTP"
    HttpGetWithRetry = ""
End Function

' SEC 限流: 至少 100ms 间隔(≤10 req/sec)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private LastSecRequestMs As Double

Private Sub EnforceSecRateLimit()
    Const SEC_MIN_INTERVAL_MS As Long = 110    ' 100ms + 10ms safety
    Dim nowMs As Double: nowMs = Timer * 1000
    Dim elapsed As Double: elapsed = nowMs - LastSecRequestMs
    If LastSecRequestMs > 0 And elapsed < SEC_MIN_INTERVAL_MS Then
        Sleep CLng(SEC_MIN_INTERVAL_MS - elapsed)
    End If
    LastSecRequestMs = Timer * 1000
End Sub
```

**子任务 2B — `RunCachedHttpGet`(Step 1)调用 `HttpGetWithRetry`**:

把 Step 1 §1B 里 `Select Case source` 直接调 `EdgarHttpGet` 等的部分改成调 `HttpGetWithRetry`。retry 自动统计写到 `result.RetryCount`,自动写诊断 P 列。

### 验证

- 跑 `一键美股`(AAPL)正常:`RetryCount = 0`(第 1 次成功)
- **手工模拟限流**:连续跑 `一键美股` 10 个 ticker,看到 SEC 请求间隔 ≥ 100ms(可以加 timing 日志验证)
- 模拟 429:把 EDGAR url 改成会返回 429 的(eg 故意构造非法 ticker / 头部不全),看到 `RetryCount = 3` + `StatusText = RETRY_EXHAUSTED`

### Generator 不要做

- ❌ 不要在 retry 之间调 `Application.Wait`(会冻 Excel UI)— 用 Win32 `Sleep` API
- ❌ 不要把 `Sleep` 用在主线程长时间(eg 60s)— 退避最长 2s + jitter 是合理上限
- ❌ 不要给 雪球 / stockanalysis 加 SEC 那种 100ms hard limit(它们没明确要求,不要无故 throttle)

---

## Step 3 — P0-01 CleanReleaseWorkbook 宏(简化版)

### 背景

GPT 5.5 Pro 的 P0-01 提议改 secret 管理(env var → secrets.json → 输入框 → B5)对个人单用户工具是 over-engineering。Reviewer 简化为:**保留 B5 cookie 作为日常输入,新增 1 个一键宏让用户分享 xlsm 前主动清**。

### 实现

**子任务 3A — 新增 `CleanReleaseWorkbook` Public Sub**(放在 `模块_工具函数.bas` 末尾或新模块 `模块_发布清理.bas`):

```vba
' Phase 4l Step 3: 发布清理 - 用户分享 xlsm 前主动调
'   清空内容: 样本池 B5 cookie + 3 张诊断 sheet 历史行 + .cache/ 本地缓存目录
'   Note: 不清作者元数据 (Excel 内置属性), 也不清 webextension 残留 — 这些需要外部脚本
Public Sub CleanReleaseWorkbook()
    Dim st As TAppState
    On Error GoTo EH
    st = BeginAppState("正在清理发布版...")
    
    Dim wsPool As Worksheet
    Set wsPool = ThisWorkbook.Sheets("样本池")
    
    ' 1. 清 cookie
    On Error Resume Next
    wsPool.Range("E5").Value = ""    ' Phase 4i 后 cookie 在 E5
    wsPool.Range("B5").Value = ""    ' 老版位置兼容
    Err.Clear
    On Error GoTo EH
    
    ' 2. 清 3 张诊断 sheet 历史行 (保留表头 R1-R2, 清 R3+)
    Dim diagNames As Variant: diagNames = Array("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断")
    Dim diagName As Variant
    For Each diagName In diagNames
        On Error Resume Next
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CStr(diagName))
        If Not ws Is Nothing Then
            Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 3 Then
                ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 17)).Clear  ' 17 列覆盖 Phase 4l 新列
            End If
        End If
        Err.Clear
        On Error GoTo EH
    Next diagName
    
    ' 3. 清 .cache/ 目录 (调已有 ClearLocalCache)
    On Error Resume Next
    ClearLocalCache
    Err.Clear
    On Error GoTo EH
    
    ' 4. 提示用户额外手工步骤
    EndAppState st
    MsgBox "发布清理完成:" & vbCrLf & vbCrLf & _
           "已清:" & vbCrLf & _
           "  - 样本池 cookie (E5)" & vbCrLf & _
           "  - 3 张诊断 sheet 历史行" & vbCrLf & _
           "  - 本地 HTTP 缓存 (.cache/)" & vbCrLf & vbCrLf & _
           "建议手工再做(Excel 内置功能):" & vbCrLf & _
           "  - 文件 → 信息 → 检查问题 → 检查文档 → 删除作者/修改者属性" & vbCrLf & _
           "  - 另存为新文件名(eg 上市公司财务数据查询_release.xlsm)", _
           vbInformation, "Phase 4l 发布清理"
    Exit Sub
    
EH:
    EndAppState st
    MsgBox "清理过程出错: " & Err.Description, vbExclamation, "Phase 4l 发布清理"
End Sub
```

**子任务 3B — 工具栏加按钮**(可选):

样本池 Q 列已有 3 个按钮(一键全抓 / 显示隐藏 / 清缓存)。新增 1 个 `BtnCleanRelease` 按钮放在 Q13:Q15(或 R 列):
```python
("BtnCleanRelease", "发布清理", "模块_工具函数.CleanReleaseWorkbook", "Q13:Q15", SECONDARY_FILL, SECONDARY_FG, 10, True),
```

或者**不加按钮**,只让宏在 VBE 中可调用 + 在使用说明 sheet 加说明。**Codex 选不加按钮**(避免按钮区拥挤;用户分享前从 VBE 调即可,场景低频)。

### 验证

- 装完 xlsm,样本池 E5 / B5 填测试 cookie,跑一次抓数生成诊断行
- 在 VBE Immediate 窗口跑 `CleanReleaseWorkbook`
- 验证 E5 / B5 空,3 张诊断 sheet R3+ 为空,`.cache/` 目录为空
- MsgBox 弹出提示

### Generator 不要做

- ❌ 不要清作者 / 修改者元数据(用 VBA 改 `BuiltinDocumentProperties` 不可靠,推荐用户用 Excel 内置功能)
- ❌ 不要清 `.git/` / `.claude/` / 任何项目元数据(那是 dev 文件,不在 xlsm 里)
- ❌ 不要把这宏加到 一键全抓 等正常流程(只在用户主动调用时执行)

---

## Step 4 — 端到端回归 + STATUS §CC 收口

### 4A. 跑 frozen 5 张

```bash
py tools/test_fx_live.py --skip-install
py -u tools/diff_phase4f_step3_lite.py
py -u tools/inspect_phase4g_state.py
py -u tools/inspect_phase4h_state.py
py -u tools/inspect_phase4k_state.py
```

任一退化立即停下。**注意**:Step 1 改诊断 sheet 列数(11 → 17),`inspect_phase4h_state.py` 如果硬编码检查 11 列,需要同步改成 17 列(state-bound inspect 同步规则,顶部加 `# Phase 4l: 同步诊断 sheet 17 列`)。

### 4B. 新增 `tools/inspect_phase4l_state.py`

检查项:
1. 诊断 sheet 17 列表头正确(L=CacheStatus / M=CacheAgeHours / N=HTTPStatus / O=ElapsedMs / P=RetryCount / Q=ErrorStage)
2. cache 命中 smoke:故意写一个 cache 文件 → 调 RunCachedHttpGet → result.CacheStatus="HIT"
3. retry smoke(可选):写一个 macro 模拟 429 响应 → `RetryCount = 3` + `StatusText = RETRY_EXHAUSTED`
4. CleanReleaseWorkbook smoke:set cookie → 调 CleanReleaseWorkbook → 验证 E5/B5 空 + cache 空

### 4C. STATUS §CC 收口

模仿 §BB 格式追加:

```markdown
## CC. Phase 4l 收口: 优化 Sprint 2 — HTTP/cache 诊断遥测 + 重试退避 + 发布清理

执行依据: PHASE_4L_PLAN.md v1。状态: ✅ Codex 已实现并通过 5 张 frozen 回归 + 新增 inspect。

### CC.1 已完成
- [Step 1 P1-01] THttpResult struct + RunCachedHttpGet wrapper, 诊断 sheet 扩 6 新列 (CacheStatus / CacheAgeHours / HTTPStatus / ElapsedMs / RetryCount / ErrorStage)
- [Step 2 P1-03] HttpGetWithRetry helper (指数退避 500/1000/2000ms + jitter), SEC 独立限流 ≤10/s
- [Step 3 P0-01] CleanReleaseWorkbook 宏(清 cookie + 诊断历史 + cache, 提示手工清作者元数据)

### CC.2 验证结果
[5A 回归 + 5B inspect 结果]

### CC.3 已知边界
- 诊断 sheet 列数从 11 → 17, inspect_phase4h_state.py 同步更新断言
- CleanReleaseWorkbook 不清作者元数据, 用户需手工用 Excel 内置功能
- HTTP retry 仅对 408/429/5xx 重试, 4xx (除 408/429) 不重试 (避免无限循环)
- SEC 限流 ≥110ms 间隔, 雪球/stockanalysis 不加额外 throttle (无明确要求)
```

PHASE_4L_PLAN.md v1 → v2,标记 ✅ Phase 4l 全期闭环。

---

## ⚠️ 全 Phase 严禁动

| 文件/区域 | 原因 |
|---|---|
| `模块_抓汇率.bas` | Phase 4f frozen |
| 旧 `GetFxRate` / `GetFxRateStatus` / `GetFxFromSheet` 签名 | Phase 4k frozen |
| 旧 `EdgarHttpGet / XueqiuHttpGet / StockAnalysisHttpGet` 签名 | 本期 wrapper 层加遥测,核心 HTTP 不动 |
| `汇率` sheet 8 列结构 | Phase 4f frozen |
| 4 市场 fetch 模块字段映射 | Phase 4c-4h frozen |
| 16 张分市场 sheet 内容 + 跨市场指标表 | Phase 4j frozen |
| 5 张 frozen 回归驱动核心断言 | Phase 4f-4k 验证基线 |
| 样本池 R14+ 用户数据 | 数据安全 |

## ⚠️ 联系 Planner 触发条件

- Step 1 改诊断 sheet 列数后,任一已有 inspect / smoke macro 因列错位失败
- Step 2 SEC 限流后 `一键美股` 抓数耗时 > 现有 2x(说明 throttle 太激进)
- Step 2 retry 后偶发请求总时长 > 30s(说明退避过长,需要降级)
- Step 3 CleanReleaseWorkbook 宏在某些 Excel 版本(eg 2016)崩溃
- 任一 frozen 回归 PASS → FAIL
- VBE Compile 失败

## State-bound inspect 同步规则

`inspect_phase4h_state.py` / `inspect_phase4k_state.py` 如果硬编码诊断 sheet 列数 = 11,Step 1 后改成 17,顶部加 `# Phase 4l: 同步诊断 sheet 17 列`,不算违反 frozen。

# Phase 4f Step 2 — Generator 细化 task list

> **作者**: Planner sub-agent (Opus 4.7 1M)
> **日期**: 2026-05-03
> **状态**: READY FOR GENERATOR
> **范围**: Plan §Step 2(汇率 sheet 模板)+ §Step 3(VBA helpers)合并执行
> **依据**:
> - `PHASE_4F_RMB_PLAN.md` v2 修订纪要(7 条 must-fix)
> - `samples/RMB_FX_PROBE.md` Step 1 实测纠正
> - 现有 `modules/模块_工具函数.bas` line 608 `XueqiuHttpGet` / line 641 `ReadXueqiuCookie` 不能改

## Step 1 修订 → Step 2 task 映射(强制 must-fix)

| Step 1 must-fix | 落地 task |
|---|---|
| 1. FX symbol 必带 `.FX` | 2B § `FetchFxKline` symbol 拼接 hardcode `.FX`;2C § `GetFxRate` currency↔symbol map 表只有 `USDCNY.FX/HKDCNY.FX/KRWCNY.FX` |
| 2. FX HTTP `Accept-Encoding: gzip, deflate` | 2B § `XueqiuFxHttpGet` 完全独立模块,**不复用** `XueqiuHttpGet`(line 608) |
| 3. 进程启动 GET `https://xueqiu.com/hq` warmup | 2B § `WarmupXueqiuSession()` 拿 `xq_a_token` 后用模块级 `Static` 变量缓存 |
| 4. quote `data[]` 直接 list | 本期 quote 不用,留 Step 4 — 但 2B 注释明确写 schema 防误用 |
| 5. kline `data.item[]` 12 列, ts=item[0], close=item[5] | 2B § `FetchFxKline` ParseJson 后索引 5 取 close(VBA 1-based 对应 6) |
| 6. K 线 `period=day&type=before&count=-2000&begin=今天 CST 0 点 ms` | 2B § `EnsureFxRateCached` 默认 count=-90 单期间足够;2B § `FetchFxKline` begin 计算函数用 CST 0 点 |
| 7. 不动 line 560/619/666 identity | 2B § 新建 `模块_抓汇率.bas` 完全独立,工具函数模块只增不改 |

---

## 子任务 2A — 新建『汇率』sheet 模板

### 具体动作(8 个)

**2A.1 在 `tools/build_template.py` 新增 `build_fx_sheet` 函数**(插入位置: `build_diagnostic_sheet` 之后, `main` 之前):

```python
def build_fx_sheet(ws):
    """
    汇率 sheet 模板:
      Row 1: 8 列表头(深蓝白字)
      Row 2+: 由 VBA 模块_抓汇率 自动填充, 用户也可手填 override
      列宽: A=14 报告期 / B-G=14 数值 / H=40 备注
      冻结 A2
    """
    ws.column_dimensions["A"].width = 14
    for letter in ["B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[letter].width = 14
    ws.column_dimensions["H"].width = 40

    headers = ["报告期", "USDCNY期末", "USDCNY期均",
               "HKDCNY期末", "HKDCNY期均",
               "KRWCNY期末", "KRWCNY期均", "备注/override"]
    fill = PatternFill("solid", fgColor=DARK_BLUE)
    for j, txt in enumerate(headers, start=1):
        col = get_column_letter(j)
        cell = ws[f"{col}1"]
        cell.value = txt
        cell.font = HEADER_FONT
        cell.fill = fill
        cell.alignment = CENTER
        cell.border = BORDER

    for r in range(1, 200):
        ws.cell(row=r, column=1).number_format = "@"   # 报告期文本格式

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
```

**2A.2 在 `tools/build_template.py` `main()` 末尾创建汇率 sheet**:
```python
    # ---- Phase 4f Step 2: 汇率 sheet ----
    ws_fx = wb.create_sheet("汇率")
    build_fx_sheet(ws_fx)
```

**2A.3 在 `tools/install_modules.py` 新增 `_make_fx_sheet` helper**(插入位置: `_make_diagnostic_sheet` 之后):
```python
def _make_fx_sheet(wb, name="汇率"):
    """Phase 4f Step 2: 汇率 sheet (8 列表头, 跨市场共享缓存)"""
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name

    headers = ["报告期", "USDCNY期末", "USDCNY期均",
               "HKDCNY期末", "HKDCNY期均",
               "KRWCNY期末", "KRWCNY期均", "备注/override"]
    widths = [14, 14, 14, 14, 14, 14, 14, 40]
    for j, (txt, w) in enumerate(zip(headers, widths), start=1):
        c = ws.Cells(1, j)
        c.Value = txt
        c.Font.Name = "微软雅黑"
        c.Font.Size = 11
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        ws.Columns(j).ColumnWidth = w

    ws.Columns("A").NumberFormat = "@"
    ws.Rows(1).RowHeight = 22

    try:
        ws.Activate()
        wb.Application.ActiveWindow.SplitColumn = 0
        wb.Application.ActiveWindow.SplitRow = 1
        wb.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass
    return ws
```

**2A.4 在 `ensure_market_sheets` 增加汇率 sheet 创建逻辑**(末尾追加):
```python
    if "汇率" in {sh.Name for sh in wb.Sheets}:
        print("  ~ sheet 已存在: 汇率")
    else:
        _make_fx_sheet(wb, "汇率")
        print("  + sheet 新建: 汇率")
```

**2A.5 在 `reorder_report_sheets` 把汇率 sheet 排到最后**:
```python
    desired_order = [
        "使用说明", "样本池",
        "A股_资产负债表", "A股_利润表", "A股_现金流量表", "A股_指标表",
        "美股_资产负债表", "美股_利润表", "美股_现金流量表", "美股_指标表",
        "美股_抓取诊断",
        "港股_资产负债表", "港股_利润表", "港股_现金流量表", "港股_指标表",
        "港股_抓取诊断",
        "韩股_资产负债表", "韩股_利润表", "韩股_现金流量表", "韩股_指标表",
        "韩股_抓取诊断",
        "汇率",   # ← Phase 4f Step 2 新增
    ]
```

**2A.6** print 列表自动包含,无需改动。

**2A.7 文档同步: `update_intro_sheet` 加汇率说明** (插入 `lines` 列表):
```python
        "",
        "汇率与币种 (Phase 4f Step 2 起)",
        "新增『汇率』sheet 缓存 USDCNY / HKDCNY / KRWCNY 期末与期间平均汇率。",
        "数据源: 雪球 K 线 USDCNY.FX / HKDCNY.FX / KRWCNY.FX, 期间平均 = 区间内日 close 算术平均。",
        "用户可在 汇率 sheet 手填 cell 覆盖系统拉取值; 备注列可写 override 理由。",
        "样本池 B6 切换『显示币种』: 默认『原币』; 选『统一RMB』后 Step 4 跑数会自动 × 汇率写 RMB。",
```

**2A.8 删除旧『统一 RMB 折算留到后续 Phase』占位句**(被 2A.7 覆盖)。

### 验收 criteria

| 检查项 | 方法 |
|---|---|
| 1. 模板新建 `汇率` sheet | `py tools/build_template.py` 后 openpyxl 打开,`"汇率" in wb.sheetnames` |
| 2. 8 列表头精确匹配 | 读 Row 1 cell A1-H1,值 == 列表 |
| 3. 列宽 A=14 / B-G=14 / H=40 | `ws.column_dimensions["A"].width == 14` etc |
| 4. 冻结 A2 | `ws.freeze_panes == "A2"` |
| 5. install 跑后旧 .xlsm 也补出 `汇率` sheet | 用 4e baseline xlsm 跑 install |
| 6. 使用说明 sheet 含汇率新段 | grep "汇率与币种" 出现 1 次 |

---

## 子任务 2B — 新建 `模块_抓汇率.bas`

### 完整文件骨架

文件头:
```vba
Attribute VB_Name = "模块_抓汇率"
Option Explicit

' Phase 4f Step 2: 雪球 FX 抓取核心模块
'
' ⚠️ 关键约束 (Step 1 RMB_FX_PROBE.md 实测锁定):
'   1. FX symbol 必带 .FX: USDCNY.FX / HKDCNY.FX / KRWCNY.FX
'   2. Accept-Encoding 必须 "gzip, deflate" (NOT identity)
'      → 不复用 模块_工具函数.bas line 608 XueqiuHttpGet
'   3. 进程启动需 GET https://xueqiu.com/hq 拿 xq_a_token cookie
'   4. quote response: parsed("data") 直接是 list (本期不用,Step 4 用)
'   5. kline response: parsed("data")("item") 是 list, 12 列
'      VBA 1-based: ts = items(i)(1), close = items(i)(6)
'   6. begin = CST 0 点 ms

' --- 模块级 cookie cache (进程内复用) ---
Private g_strFxWarmupCookie As String
Private g_blnFxWarmupDone As Boolean
```

### 核心函数(完整实现见上 Planner 输出)

**2B.3 `WarmupXueqiuSession()`** — GET `https://xueqiu.com/hq` 拿访客 cookie,模块级 cache,进程内只跑 1 次

**2B.4 `XueqiuFxHttpGet(strUrl)`** — 独立 FX HTTP 客户端,gzip + warmup + 注入 warmup cookie + B5 cookie(enhancement)

**2B.5 `BeginCstMsForToday()` / `PeriodEndCstMs(periodEnd)`** — CST 0 点 → ms helpers,用 `Currency` 防 Long 溢出

**2B.6 `FetchFxKline(currency, beginCstMs, count)`** — 拉某币种 K 线,返回 2D Variant array `(i,0)=ts_ms / (i,1)=close`
关键: VBA 1-based 索引 → `items(i)(1)` 是 ts (Step 1 0-based item[0]), `items(i)(6)` 是 close (Step 1 0-based item[5])

**2B.7 `EnsureFxRateCached(periodEnd, currency)`** — 主入口
- RMB/CNY → True
- 查 sheet 缓存命中 → True
- 缺则拉 K 线 → 算期末 + 期均 → 写 sheet → True

**2B.8 `InferPeriodStart(periodEnd)`** — 假设财年从 01-01 开始

**2B.9 `ComputeEopRate(items, periodEnd)`** — 取 ts <= periodEnd+1day 的最大 ts 对应 close

**2B.10 `ComputeAvgRate(items, periodStart, periodEnd)`** — 区间内 close 算术平均

**2B.11 用户 override 检测** — `IsNumeric(eopVal) And CDbl(v) > 0` 隐含,无需额外代码,加注释

**2B.12 不修改 line 560/619/666** — Evaluator grep 验

**2B.13 限速** — Cache 命中跳 HTTP,自然限频,不显式 sleep

**2B.14 文档** — 已在 2A.7 覆盖

### 验收 criteria

| 检查项 | 方法 |
|---|---|
| 1. 文件存在 + Module name | 首行 `Attribute VB_Name = "模块_抓汇率"` |
| 2. install 后 VBA 模块成功导入 | `+ installed: 模块_抓汇率` |
| 3. **不动** line 560/619/666 | grep `identity` modules/模块_工具函数.bas 仍 3 处 |
| 4. FX HTTP gzip | grep `gzip, deflate` modules/模块_抓汇率.bas 出现 2 次 |
| 5. Symbol .FX | grep `\.FX` 出现 ≥ 3 次 |
| 6. 立即窗口 `EnsureFxRateCached("2024-12-31", "USD")` | 汇率 sheet 出现 2024-12-31 行 + B 列 ≈ 7.0 + C 列 ≈ 7.2 |
| 7. 第二次相同参数毫秒返回 | cache 命中 |
| 8. RMB/CNY → True 不写 sheet | `EnsureFxRateCached("2024-12-31", "RMB") = True` |
| 9. KRW symbol map | F 列 ≈ 0.0049 / G 列 ≈ 0.0050 |
| 10. Warmup 进程内只 1 次 | `Debug.Print "warmup"` 探针, 跑 3 次 EnsureFxRateCached 只 print 1 次 |

### 风险

| 风险 | 兜底 |
|---|---|
| WinHttp 5.1 不自动 inflate gzip | `Debug.Print Left$(raw, 80)` 看是 JSON 还是二进制; 退化用 ServerXMLHTTP + 手工 zlib |
| JsonConverter Collection 1-based 错位 | 注释强调 `items(i)(1)`/`items(i)(6)` 语义对应 Step 1 0-based item[0]/item[5] |
| 节假日 periodEnd 当日无 day bar | `ComputeEopRate` 用 `<=` + 取最大 ts 实现回溯 |
| Cookie 全空 | Step 1 实测无 cookie 也通; 已 `If Len(userCookie) > 0` 才注入 |
| Currency 类型溢出 | 全程用 `Currency`(8 字节, ±9.2e14) |

### 依赖

- 2B 必须晚于 2A(需 sheet 存在)
- 2B 必须晚于 2C(需 `FX_SHEET` / `FX_DATA_ROW` 常量)

---

## 子任务 2C — `模块_工具函数.bas` 新增 helpers

### 具体动作(8 个)

**2C.1 模块顶部新增常量**(`Option Explicit` 之后任何 Sub/Function 之前):
```vba
' --- Phase 4f Step 2: 汇率 sheet 常量 ---
Public Const FX_SHEET As String = "汇率"
Public Const FX_DATA_ROW As Long = 2
```

**2C.2 新增 `ReadDisplayCurrency()`**(`ReadXueqiuCookie()` 之后):
```vba
Public Function ReadDisplayCurrency() As String
    On Error Resume Next
    Dim s As String
    s = Trim$(CStr(ThisWorkbook.Sheets("样本池").Range("B6").Value))
    If Err.Number <> 0 Or Len(s) = 0 Then s = "原币"
    Err.Clear
    On Error GoTo 0
    ReadDisplayCurrency = s
End Function
```

**2C.3 新增 `GetFxRate()` + 2 个 Private helpers** — 完整实现见 Planner output(LookupFxRowByPeriod / LookupFxColForCurrency)

**2C.4 `FxColForCurrency` 提到 `模块_抓汇率` 改 Public 共用** — 推荐(避免 drift),否则 2C 自己 Private 副本

**2C.5 不修改任何现有 helper** — 全是追加

**2C.6 验证 line 560/619/666 不变** — grep `Accept-Encoding.*identity` 仍 3 行

**2C.7-2C.8 文档注释强调返回 0 的语义 + Step 4 调用契约**

### 验收 criteria

| 检查项 | 方法 |
|---|---|
| 1. `Public Const FX_SHEET` 出现 1 次 |  |
| 2. ReadDisplayCurrency 默认 "原币" | 删 B6 后 `?ReadDisplayCurrency()` → `原币` |
| 3. GetFxRate("RMB", *, *) = 1 |  |
| 4. GetFxRate("USD", "2024-12-31", True) > 6.5 |  |
| 5. GetFxRate("USD", "2024-12-31", False) > 6.5 |  |
| 6. GetFxRate("XYZ", *, *) = 0 |  |
| 7. line 560/619/666 不变 | grep `identity` 仍 3 行 |

---

## 子任务 2D — 样本池 B6 toggle『显示币种』

### 具体动作(6 个)

**2D.1 修改 `tools/build_template.py` `build_sample_pool` 加 A6/B6**:
- A6 label "显示币种" + 浅蓝
- B6 default "原币" + 浅黄 + 数据验证下拉 `"原币,统一RMB"`
- `range(1, 6)` 改 `range(1, 7)`(row height 22)

**2D.2 `tools/install_modules.py` 新增 `install_currency_toggle_cell` helper**:
- A6 label / B6 default "原币" + 数据验证下拉 + comment

**2D.3 `layout_sample_pool` 加 row 6 重画**:
- label tuple 加 `("A6", "显示币种")`
- row height range 改 `range(1, 7)`

**2D.4 `layout_sample_pool` 末尾调 `install_currency_toggle_cell(ws_pool)`**

**2D.5 `migrate_old_sample_pool` 兜底** — 天然幂等, `if not val_cell.Value` 检查不覆盖用户值

**2D.6 旧 .xlsm 升级路径** — B6 是新增, 旧没值 → 写默认 "原币", 完全向后兼容

### 验收 criteria

| 检查项 | 方法 |
|---|---|
| 1. build_template 后 B6 = "原币" |  |
| 2. B6 数据验证 list `"原币,统一RMB"` |  |
| 3. install 跑后 B6 cell 有 comment |  |
| 4. 旧 4e xlsm 升级后 B6 = "原币" |  |
| 5. 用户改 B6 = 统一RMB,install 不覆盖 | `if not val_cell.Value` |
| 6. ReadDisplayCurrency() 与 B6 同步 |  |

---

## Generator 推荐执行顺序

```
1. 2C.1-2C.2 (常量 + ReadDisplayCurrency)        ← 最快 0 风险
2. 2A.1-2A.6 (build_template + install 加 sheet)
   → py tools/build_template.py + py tools/install_modules.py
3. 2D.1-2D.6 (B6 toggle)
   → install + 验证 B6 默认 "原币"
4. 2C.3-2C.4 (GetFxRate, 暂时 stub return 0)
5. 2B.1-2B.14 (模块_抓汇率 全部)
   → install + VBA F5 编译 + 立即窗口跑 EnsureFxRateCached("2024-12-31", "USD")
6. 2C 回填 (GetFxRate 真正连 EnsureFxRateCached)
7. 2A.7-2A.8 (使用说明)
```

## 全局约束: 7 条 Step 1 修订强制对照(Evaluator grep 检查表)

| # | Step 1 修订 | 验收 grep |
|---|---|---|
| 1 | FX .FX | `grep -E "USDCNY\.FX\|HKDCNY\.FX\|KRWCNY\.FX" modules/模块_抓汇率.bas` ≥ 3 行 |
| 2 | gzip+deflate | `grep "Accept-Encoding" modules/模块_抓汇率.bas` 全 gzip; `grep "identity" modules/模块_工具函数.bas \| wc -l` = 3 |
| 3 | warmup `xueqiu.com/hq` | `grep "xueqiu\.com/hq" modules/模块_抓汇率.bas` ≥ 1 |
| 4 | quote schema (本期不实施) | N/A |
| 5 | kline 1-based ts/close | `grep "items(i)(1)\|items(i)(6)" modules/模块_抓汇率.bas` 各 1 |
| 6 | begin CST 0 点 ms | `grep "BeginCstMsForToday\|PeriodEndCstMs" modules/模块_抓汇率.bas` ≥ 1 |
| 7 | 不动 line 560/619/666 | `git diff` line 552-680 区域 0 改动 |

## Smoke Tests(Evaluator 端到端验证)

1. **模板生成**: `rm *.xlsm *.xlsx; py tools/build_template.py; py tools/install_modules.py` → xlsm 成功 + sheet 列表末尾 `汇率` + B6="原币" + 汇率 Row 1 = 8 列表头
2. **VBA 编译**: Alt+F11 → Debug → Compile VBAProject → 0 error
3. **立即窗口手测**:
   ```
   ?ReadDisplayCurrency() → "原币"
   ?GetFxRate("RMB", "2024-12-31", True) → 1
   ?GetFxRate("USD", "2024-12-31", True) → ~7.0
   ?GetFxRate("USD", "2024-12-31", False) → ~7.1
   ?GetFxRate("HKD", "2024-09-30", True) → ~0.91
   ?GetFxRate("KRW", "2024-12-31", False) → ~0.005
   ```
4. **Cache 命中**: 跑 3 次同 USD → 第二/三次毫秒返回, 探针 `Debug.Print` 只 1 行
5. **用户 override**: 手填 B5=6.50 → GetFxRate 返 6.50 不重拉
6. **B6 toggle**: 改 "统一RMB" → ReadDisplayCurrency 返回; 清空 → "原币" 默认

---

## 备注: Step 4 接口契约预告

```vba
Dim displayCurrency As String: displayCurrency = ReadDisplayCurrency()
Dim fxRate As Double: fxRate = 1#
If displayCurrency = "统一RMB" Then
    fxRate = GetFxRate(企业报告币种, 报告期, (报表类型 = "BalanceSheet"))
    If fxRate <= 0 Then fxRate = 1#    ' 失败兜底
End If
cellValue = cellValue * fxRate
```

企业报告币种来源:
- A股 4 家: "RMB"
- 美股: EDGAR unit "USD"
- 港股: `finance/hk/balance.json` `currency` 字段(02519 = "CNY")
- 韩股: "KRW"

本期 Step 2/3 只准备 `GetFxRate` API,Step 4 才接入。

## 完成后产出文件清单

新增:
- `modules/模块_抓汇率.bas` (~250-300 行)
- `汇率` sheet

修改:
- `tools/build_template.py` (~+50 行)
- `tools/install_modules.py` (~+80 行)
- `modules/模块_工具函数.bas` (~+80 行追加, 0 改动现有)

不动:
- `modules/模块_工具函数.bas` line 552-676 区域
- 所有 `modules/模块_抓*.bas`(Step 4 才改)
- 所有现有 sheet(除汇率新增)

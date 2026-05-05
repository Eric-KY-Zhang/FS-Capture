# Phase 4f: 统一 RMB 跨市场对标(#5 实施 + 可选 #2 合表)

> **版本**: v3(2026-05-03,Step 2 联网验证通过后再修订)
> **状态**: ✅ **Phase 4f 全期闭环** — Step 3-7 已实现;联网 FX 回归 5/5 PASS,本地 RMB hook / 诊断 FX_Rate / A1 注释 / R1 币种 tag smoke 回归通过。
> **作者**: Claude(planner)+ Codex 离线后改用 Claude Code 三角架构

## v3 修订纪要(Step 2 联网验证后)

| # | 修订 | 详情 |
|---|---|---|
| 1 | **HTTP 实现路径锁定**: PowerShell shell-out | Codex 实测 `WinHttp.WinHttpRequest.5.1` + `Accept-Encoding: gzip, deflate` 拿到的是 gzip 字节,**不会自动 inflate**;`ServerXMLHTTP.6.0` 同样症状。解决: shell out 到 `powershell.exe` + `System.Net.Http.HttpClient` (`AutomaticDecompression` + `UseCookies`)。PS 把响应 JSON 写临时文件,VBA 用 `ADODB.Stream` UTF-8 解码读回。**单次 PS 调用开销 ~1.2s**(含 warmup `https://xueqiu.com/hq` + cookie session + kline fetch);缓存命中场景 ~0ms 跳过 HTTP |
| 2 | **联网验证实测汇率值**(2026-05-03 跑) | USDCNY 2024-12-31 EOP=7.3002 / AVG=7.1966; HKDCNY 2024-12-31 EOP=0.93971 / AVG=0.9223; KRWCNY 2024-12-31 EOP=0.00494 / AVG=0.005280; USDCNY 2023-12-31 EOP=7.0999 / AVG=7.0828; HKDCNY 2023-12-31 EOP=0.90914 / AVG=0.9047 — 全部与公开汇率一致 |
| 3 | **B5 cookie 可空** | 实测 B5 cookie 为空时 FX 仍工作(PS 内 warmup 拿访客 `xq_a_token`)。用户填 B5 是 enhancement 不是 dependency |
| 4 | **install_modules B6 toggle 保留用户值** | install 时先记 B6 现值,layout 清 A1:Q9 后写回,避免覆盖用户已选的"原币/统一RMB" |
| 5 | **EnsureFxRateCached 部分成功也返 True** | EOP 或 AVG 任一计算成功 + 写入即返 True (容错,Step 4 调用方能拿到至少一个 rate) |
> **背景**: Phase 4e UX 闭环。用户实际场景 = 床垫/家居跨境电商同业对标 — 5 家 A 股(含港股傲基 `02519.HK`)+ 1 家韩股 Zinus → 需 RMB 统一才能横向对比。

## v2 修订纪要(Step 1 实测后,Generator + Evaluator 联合 must-fix)

| # | v1 假设 | Step 1 实测纠正 | 来源 |
|---|---|---|---|
| 1 | symbol = `USDCNY,HKDCNY,KRWCNY` | **必加 `.FX` 后缀**: `USDCNY.FX,HKDCNY.FX,KRWCNY.FX`(无 .FX → `data:[]` 静默空) | Generator 1A symbol_variants 对照 |
| 2 | `Accept-Encoding: identity` | **FX 端必须 `gzip, deflate`**(identity → body_len=0)。**仅 FX 端**;美股 finance 端继续 identity 不动 | Generator 1A plan_call vs working_call 对照 |
| 3 | 直接 fetch | **进程启动需 GET `https://xueqiu.com/hq` 暖 session**,拿访客 `xq_a_token` cookie 注入后续请求 | Generator 1A 实测 |
| 4 | response schema = `parsed["data"]["items"]` | **quote**: `parsed["data"]` 直接 list;**kline**: `parsed["data"]["item"]` 是 list,每项 12 列(`item[0]`=ts,`item[5]`=close) | Generator 1A/1B dump |
| 5 | 傲基股份代码待 probe(候选 02519/003031/09927) | **= 02519.HK**(quote_name 已 2025-04 改名"傲基股份",原"傲基控股")。**Trading currency=HKD(港币挂牌)vs Reporting currency=CNY(人民币编报)** — Step 4 港股 hook 必须从 finance API 读 `currency` 字段,不能复用 quote API | Generator 1C 双端 confirm |
| 6 | begin 时区不明 | **CST 端点 ms**(每个 day bar 落 04:00 UTC = 12:00 CST) | Generator 1B `attempts[0].first_item[0]` mod 86400000 = 14400000 ms |
| 7 | K 线 count 上限不明 | **~2147 条硬上限**;FX 数据起 2018-02-07;production 用 `count=-2000` + `begin=今天 CST` 单次覆盖 7.7 年,**不要分段** | Generator 1B count_cap_probe |

## Step 1 → Step 2 must-fix(Planner 下一轮必须写进 task list)

1. **FX symbol 硬编码必须带 `.FX` 后缀**
2. **FX HTTP 客户端**: `gzip+deflate` + warmup `https://xueqiu.com/hq` + cookie env 作 enhancement(不是 dependency)
3. **Response 解析**: quote `data[]` / kline `data.item[]` 12 列, ts=item[0] / close=item[5]
4. **傲基代码 hardcode**: `02519`(5 位裸码,不带 `.HK`);**reporting currency 从 finance API 读, 不复用 quote**
5. **K 线参数**: `period=day&type=before&count=-2000&begin=<今天 CST 0 点 ms>`
6. **期间平均算法**: 算术平均 `item[5]` (close),节假日缺口自动跳过;期末汇率取 `item[5]@period_end`,缺则向前回溯到最近交易日

## ⚠️ 不可盲改

- `modules/模块_工具函数.bas` line 560/619/666 的 `Accept-Encoding: identity` **保留不动**(美股 finance 端 4a-4e 已验证工作);**FX 端用独立 HTTP 客户端 / 独立 Headers**,不混用

## Context — 用户实际场景

| 公司 | 代码 | 市场 | 报告币种 |
|---|---|---|---|
| 安克创新 | 300866 | A股(深交所创业板) | RMB |
| 梦百合 | 603313 | A股(上交所) | RMB |
| 喜临门 | 603008 | A股(上交所) | RMB |
| 致欧科技 | 301376 | A股(深交所创业板) | RMB |
| **傲基股份** | **代码待确认** | **市场待确认** ⚠️ | **币种待确认** ⚠️ |
| Zinus | 013890 | 韩股(KOSDAQ) | KRW |

⚠️ **傲基股份**(Aoji Tech,跨境电商,跟安克同业): 代码可能是 `02519.HK`(2024-02 港股 IPO)/ 也可能 A 股新上代码。**Codex Step 1 需 probe 三个候选**(02519 港股 / 003031 A股 / 其它),实测哪个能命中。

**跨市场对标核心痛点**: 现状 6 家公司在 4 张 sheet 各市场分开 + 不同币种,无法直接比较"安克 vs Zinus 的资产规模"。

## 已锁定的决策(用户已确认)

| 决策 | 选项 |
|---|---|
| 汇率策略 | **期末/期间平均混合(会计准则推荐)** — BS 用期末汇率(时点数);IS / CF 用期间平均汇率(发生数) |
| 数据源 | 雪球 quote API + K 线 endpoint(`USDCNY` / `HKDCNY` / `KRWCNY`) + **用户手动 override 兜底**(在新 sheet `汇率` 某个 cell 改值) |
| UX 切换 | 样本池新增 toggle `B6 = 显示币种(原币 / 统一RMB)`,默认"原币"(向后兼容) |
| 实施范围 | **#5 RMB 统一**(本期主线)+ **可选 #2 合表**(plan 末尾决策点)|
| Sheet 命名 | 新增 `汇率` sheet(单一,跨市场共享)|

## 实施方案(分 Step,每 Step 完成 commit + 我 review)

### Step 1 — 探查雪球汇率 API + 锁傲基代码(必须先做并暂停 review)

#### 1A 雪球 quote 实时汇率
测:
```
https://stock.xueqiu.com/v5/stock/realtime/quotec.json?symbol=USDCNY,HKDCNY,KRWCNY
```
- 确认是否需要 cookie(应该不要,quote 是公开)
- 确认返回字段(`current` / `last_close` / `change`)
- dump 到 `samples/xueqiu_quote_currency.json`

#### 1B 雪球 K 线历史汇率
测:
```
https://stock.xueqiu.com/v5/stock/chart/kline.json?symbol=USDCNY&begin=1672531200000&period=day&type=before&count=-365
```
- 拉 USDCNY 一年 daily 看格式(`item` 数组,每个 `[timestamp, volume, open, high, low, close, ...]`)
- 验证能否拉到 2020-01-01 → 2025-12-31 的全部 daily
- 确认时区(UTC vs CST)
- dump 到 `samples/xueqiu_kline_USDCNY.json`(只 dump 一年的 sample)

#### 1C 傲基股份代码探查
测候选(用现有 fetch 框架):
- `02519` 港股(雪球 HK / stockanalysis HKEX)
- `003031` A 股(新浪 A 股)
- `09927` 港股(可能性低,作为兜底)

dump 到 `samples/aoji_probe_*.json/html` + 报告里写"傲基股份 = ?"

#### 1D 报告
写 `samples/RMB_FX_PROBE.md`:
- 1A quote 是否通 / 字段格式
- 1B K 线历史能拉多远 / 数据连续性
- **傲基股份确认代码 + 市场**
- 期间平均汇率算法建议(月度算术平均 / 加权 / etc)

**暂停等 Claude review**

### Step 2 — `汇率` sheet 模板

新建 sheet `汇率`,布局:

```
       A          B            C            D            E            F            G            H
Row 1: 报告期    USDCNY期末   USDCNY期均   HKDCNY期末   HKDCNY期均   KRWCNY期末   KRWCNY期均   备注/override
Row 2: 2020-12-31  6.5249      6.8979      0.8409      0.8898      0.005965    0.005844
Row 3: 2021-03-31  6.5500      6.4798      0.8423      0.8344      0.005777    0.005757
...
Row N: 2025-12-31  ...         ...         ...         ...         ...         ...
```

**列宽**: A=14 / B-G=14 / H=40

**自动填充**:
- 第一次用户跑跨市场抓数时,VBA 检测 `汇率` sheet 是否有缺失报告期
- 如缺失,调雪球 K 线 endpoint 拉取 + 缓存到 sheet
- 后续跑数复用缓存,不重复拉

**用户 override**:
- 用户改某 cell B5 → VBA 跑数时优先用 cell 值,不再调 API
- H 列 备注/override 可写理由(eg "Bloomberg 中间价 vs 雪球收盘差 0.5%, 用 BBG")

### Step 3 — VBA 汇率读取 helpers

新增于 `模块_工具函数.bas`:

```vba
Public Const FX_SHEET As String = "汇率"
Public Const FX_DATA_ROW As Long = 2

Public Function ReadDisplayCurrency() As String
    ' 读样本池 B6 toggle: "原币" 或 "统一RMB" (默认 "原币")
    ' B6 必须在 install 时初始化
End Function

Public Function GetFxRate(ByVal currency As String, _
                          ByVal periodEnd As String, _
                          ByVal isInstant As Boolean) As Double
    ' currency: "USD" / "HKD" / "KRW" / "RMB" (RMB 直接返 1)
    ' periodEnd: "yyyy-mm-dd" 报告期
    ' isInstant: True = 期末汇率(BS 用) / False = 期间平均(IS/CF 用)
    ' 1) 检查 汇率 sheet 缓存 → 命中返回
    ' 2) 不命中 → 调 EnsureFxRateCached → 返回
    ' 3) 全失败 → 抛错或返回 0
End Function

Private Sub EnsureFxRateCached(ByVal currency As String, ByVal periodEnd As String)
    ' 调雪球 K 线拉该报告期的 USDCNY/HKDCNY/KRWCNY 数据
    ' 写到 汇率 sheet 对应 row
    ' 期间平均 = 该报告期月份的当月日均(BS 期末 = 季末当日 close;IS/CF 期间平均 = 季度内每日 close 平均)
End Sub
```

### Step 4 — 各市场 Run* 加汇率换算 hook

修改:
- `模块_抓资产负债表.bas` 等 4 个 A 股 Main(A 股本来 RMB,toggle = RMB 时不动)
- `RunUSStatement` (美股 USD → RMB)
- `RunHKStatement` (港股 CNY/HKD/USD → RMB,大多数港股已经是 CNY 不需换算)
- `RunKRStatement` (韩股 KRW → RMB)

实现位置:
- `WriteWideTable` 写 cell 前判断 `ReadDisplayCurrency() = "统一RMB"` 时,把 value `× GetFxRate(originalCurrency, periodEnd, isInstant)`
- `isInstant = True` if statement = "BalanceSheet" else False

诊断 sheet 加列:
- 旧 10 列(`Unit` 在第 8 列)
- 新 11 列:第 11 列 `FX_Rate` 显示换算用的汇率(toggle = 原币时为 1.0)

### Step 5 — 样本池 toggle + 模板更新

`模块_工具函数.bas` 新增:
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

`tools/build_template.py` `build_sample_pool` 加:
- A6 label = "显示币种"
- B6 default value = "原币" + dropdown validation list `"原币,统一RMB"`

`tools/install_modules.py` 同步加 `install_currency_toggle_cell` helper。

`build_template.py` 加 `build_fx_sheet` 函数 + `main` 末尾创建 `汇率` sheet。

`install_modules.py` `ensure_market_sheets` 加 `汇率` sheet。

### Step 6 — UI 反馈

各市场 sheet 的 A1 cell comment 改成动态:
```vba
' 在 Run*Statement 末尾(WriteWideTable 之后):
Dim displayCurrency As String: displayCurrency = ReadDisplayCurrency()
On Error Resume Next
If Not wsTarget.Range("A1").Comment Is Nothing Then wsTarget.Range("A1").Comment.Delete
If displayCurrency = "统一RMB" Then
    wsTarget.Range("A1").AddComment "单位: 百万 RMB (汇率源: 见 汇率 sheet, 期末/期间平均混合)"
Else
    ' 原币:沿用既有注释 (港股 = 各家公司报告币种 / 美股 = 百万USD / 韩股 = 十亿KRW)
End If
```

诊断 sheet J 列备注追加 `; fx=USDCNY期均=6.92`(toggle = 统一RMB 时)

### Step 7 — 可选 #2 合表(决策点)

**完成 #5 后用户可选**:
- 选项 A: 不做 #2,保留 16 张分市场表 + toggle
- 选项 B: 做 #2,删 16 张分市场表,合并 4 张主表(BS/IS/CF/Indicator),每张横向铺 N 公司列(不分市场)

**推荐 A**(谨慎):
- 16 张表已经稳定,合表破坏面大
- 用户可以手动复制粘贴或写 Excel 公式做 ad-hoc 合表
- 实在需要再做 #2

**如选 B**(本期不做,留 Phase 4g):
- 数据列 = 公司 × 报告期(同 #3 样本池布局,横向铺)
- R1 公司名 + 市场 tag(如 `300866 安克 (A)` / `013890 Zinus (KR)` / `02519 傲基 (HK)`)
- R2 报告期降序
- 4 张合并主表 sheet 名:`资产负债表` / `利润表` / `现金流量表` / `指标表`(无市场前缀)

---

## 验收方案

### Test 1 — 汇率 sheet 自动拉取
- 删除 `汇率` sheet 数据(Row 2+),保留 Row 1 表头
- 跑 任意 跨市场抓数(美股/港股/韩股)
- **期望**: `汇率` sheet 自动填充缺失报告期的 USDCNY/HKDCNY/KRWCNY 期末+期间平均

### Test 2 — 用户 override 优先
- `汇率` sheet 某 cell 手动改值(eg B5 期末汇率 改成 6.50)
- 跑数,**期望**: VBA 用 6.50 换算,不再调 API

### Test 3 — toggle 切换
- `B6 = 原币` → 跑数 → 各 sheet 数值 = 原币(USD/HKD/KRW/RMB),A1 注释原币
- `B6 = 统一RMB` → 跑数 → 数值 × 汇率 = RMB,A1 注释 RMB

### Test 4 — 用户实际场景
样本池(用户实际同业对标 6 家):
- A 股: 300866 安克 / 603313 梦百合 / 603008 喜临门 / 301376 致欧 / [傲基股份代码 — Step 1 探查后填]
- 韩股: 013890 Zinus
- (可加港股傲基 02519 if Step 1 探明)

配置: `A2=2024 / A4=Q4 / B6=统一RMB`

跑 一键全抓 4 市场

**期望**:
- A 股: 5 家公司 BS/IS/CF/Indicator,数值跟原 RMB 一致(× 1.0)
- 韩股: Zinus,数值 = 韩元值 × KRWCNY 期末/期间平均 → 百万 RMB
- 诊断 sheet `FX_Rate` 列各家公司各报告期都有汇率值
- 用户能在 A股_资产负债表 / 韩股_资产负债表 看 RMB 数值,有意义对比

---

## 文件改动清单

| 文件 | 改动 | 责任 |
|---|---|---|
| `samples/xueqiu_quote_currency.json` | quote API dump | Codex Step 1 |
| `samples/xueqiu_kline_USDCNY.json` | K 线 dump | Codex Step 1 |
| `samples/aoji_probe_*.json/html` | 傲基代码 probe | Codex Step 1 |
| `samples/RMB_FX_PROBE.md` | 探查报告 | Codex Step 1 |
| `modules/模块_工具函数.bas` | `ReadDisplayCurrency / GetFxRate / EnsureFxRateCached / FX_SHEET` 常量 | Codex Step 3 |
| `modules/模块_抓汇率.bas`(新建) | 雪球 K 线拉取 + `汇率` sheet 写入 | Codex Step 3 |
| `modules/模块_抓资产负债表.bas` 等 4 A 股 | 写 cell 前 hook 汇率换算(A 股大多 ×1.0) | Codex Step 4 |
| `modules/模块_抓美股财报.bas` `RunUSStatement` | hook | Codex Step 4 |
| `modules/模块_抓港股财报.bas` `RunHKStatement` | hook | Codex Step 4 |
| `modules/模块_抓韩股财报.bas` `RunKRStatement` | hook | Codex Step 4 |
| `modules/模块_工具函数.bas` `EnsureDiagnosticSheet` | 加 11 列 `FX_Rate` | Codex Step 4 |
| `modules/模块_工具函数.bas` `WriteWideTable` | 写 cell 前 hook | Codex Step 4 |
| `tools/build_template.py` `build_sample_pool` | 加 A6/B6 toggle | Codex Step 5 |
| `tools/build_template.py` 新 `build_fx_sheet` | 模板 | Codex Step 5 |
| `tools/install_modules.py` `install_currency_toggle_cell` | install 时设 toggle | Codex Step 5 |
| `tools/install_modules.py` `ensure_market_sheets` | 加 `汇率` sheet | Codex Step 5 |
| `tools/install_modules.py` `reorder_report_sheets` | `汇率` sheet 放最后(诊断之后) | Codex Step 5 |
| 「使用说明」refresh | 加汇率 toggle 说明 | Codex Step 5 |

## Codex 工作顺序

```
Day 1:
  Step 1 (汇率 API + 傲基代码 probe) → 暂停 review
Day 2:
  Step 2 (汇率 sheet 模板) + Step 3 (VBA helpers + 雪球 K 线拉取)
Day 3:
  Step 4 (各 Run* hook + 诊断 11 列) + Step 5 (toggle + 模板)
Day 4:
  Step 6 (UI 反馈) + Test 1/2/3/4 + STATUS.md §T 收口
```

## 风险点

1. **雪球 K 线时区** — UTC 还是 CST?periodEnd 转 timestamp 要正确否则差 8 小时拉错日
2. **期间平均算法** — 月度日均 vs 季度日均 vs 加权?业内常用"期间日数算术平均",建议 **`AVG(daily close from quarter_start to quarter_end)`**
3. **K 线 endpoint 反爬** — 雪球 K 线可能有限速,Codex 间隔 1s 拉每个币种
4. **汇率历史范围** — 雪球可能只有 5-10 年,2015 年前数据可能缺;Step 1 实测
5. **傲基股份代码** — Step 1 必须确认,否则用户实际场景跑不全
6. **WriteWideTable 改动 大面积** — 触及 4 市场所有 sheet,小心回归;建议 Side: 用现有 4b-14a baseline 跑 diff,验证 toggle = 原币时 0 mismatches

## Phase 4g 待办(future)

- **#2 合表 4 市场同表** — 等用户用 #5 一段时间反馈后决定是否做
- 数据可视化(雷达图 / 时序折线 / 结构饼图)— 等 #5 + 可能的 #2 完成后,价值最大化

## 收口: #2 合表决策

Phase 4f 全期闭环,#2 合表 defer 到 Phase 4g (等待用户实际使用反馈)。

判断依据 (2026-05-04, Step 6 完成后):
- 当前能力已经提供 4 市场分表 + B6 `原币/统一RMB` toggle;RMB 模式下 A股/港股/韩股分市场主表已经可做跨市场数值对比。
- 本阶段未出现“必须立刻在一张表里看 4 市场”的明确需求;合表涉及公司横铺/纵向堆叠、报告期对齐、跨市场指标命名差异等新设计点,留到 Phase 4g 基于实际使用反馈再定。

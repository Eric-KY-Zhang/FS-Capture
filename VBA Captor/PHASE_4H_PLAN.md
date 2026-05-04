# Phase 4h: 跨市场全合表 + B6 实时 toggle + 缓存层 + 中概美股 fallback + 4g 收账

> **版本**: v2(2026-05-04,Phase 4h 全期闭环)
> **状态**: ✅ Phase 4h 全期闭环
> **作者**: Claude(planner) + Codex(generator)
> **背景**: Phase 4g 已交付『跨市场_指标表』,把 18 标准指标 跨 4 市场合到 1 张 sheet。Phase 4h 把 BS/IS/CF 也合(完成跨市场对标工作流闭环),同时清理 4 个累积的 backlog:B6 切换需重抓的 UX、雪球/EDGAR 重复抓数浪费的性能、雪球 cookie 失效时的 fallback、Phase 4g 顺手修的 HK 双年 fetch 文档化。

## 项目语境(给 Generator 的 anchor 段)

本项目是个人用户(财务/审计专业人士)用的 Excel 桌面报表工具,用于审阅几家公司的报告期可比指标。**全部 5 件工作的网络部分都是 Phase 4c-4g 已有数据源(雪球/EDGAR/stockanalysis)的延续**:Step 4 的中概美股 fallback 是 Phase 4d 韩股 stockanalysis 路径的同站点 generalize;Step 3 的缓存是把已经发出的 HTTP 响应落到本地 `.cache/` 让重复跑数不重发请求,**降低**而不是新增网络流量;Step 1 / 2 / 5 完全本地。所有抓样维持 Phase 4d/4f 的单线程顺序 + 2s 间隔节奏,**没有任何并发循环或自动化扩面**。

## 用户已锁定的决策(2026-05-04 AskUserQuestion 收集)

| 决策点 | 选项 |
|---|---|
| 主线 | A. 跨市场 BS/IS/CF 全合表(完成 Phase 4g 工作流闭环) |
| 范围 | 主线 + 多件并行(A + B + C + D + E,5 件全做) |
| BS/IS/CF 合表布局 | **4 张子表并行**(跨市场_资产负债表 / _利润表 / _现金流量表 / _指标表 各一张) |
| 行项目对齐策略 | **Codex 提案**:实现时根据 4 张分市场表的 IndicatorsByInsertion 顺序提案 mapping,review 时再定 |
| 并行项 | B(B6 实时 toggle)+ C(stockanalysis 中概美股 fallback)+ D(磁盘缓存)+ E(4g 小尾巴收账) |

## Step 总览

| Step | 内容 | 估时 | 阻塞依赖 |
|---|---|---|---|
| 1 | E. Phase 4g 小尾巴收账(纯 doc:HK 双年 fetch + AAPL fiscal fuzzy match) | 0.3h | 无 |
| 2 | A. 跨市场 BS/IS/CF 3 张合表 sheet + VBA 通用化(Phase 4g BuildCrossMarketIndicatorSheet 抽象成 4-statement 形态) | 4-5h | 无 |
| 3 | A. 4 张合表统一刷新逻辑(一键全抓末尾 + 单独 4 个按钮 + 1 个汇总按钮) | 1.5h | Step 2 |
| 4 | B. B6 原币↔RMB 实时 toggle(WriteWideTable 重构成幂等 + 缓存原币 dump,B6 切换时只刷公式) | 4-6h | 无(独立) |
| 5 | D. 磁盘 JSON 缓存层(雪球/EDGAR/stockanalysis 原始 HTTP 响应落盘,24h TTL) | 2-3h | 无(独立) |
| 6 | C. stockanalysis 中概美股 fallback 落地(BABA/JD/PDD BS+CF 子页面抓样 + 字段映射 + 默认关闭的 fallback hook) | 4-5h | Step 5(走缓存层) |
| 7 | 端到端回归 + STATUS §W 收口 + plan v2 升级 | 1.5h | 依赖 1-6 |

**Codex 工作流建议(5 rounds)**:

| Round | Step | Commit 描述 |
|---|---|---|
| Round 1 | Step 1 | `Phase 4h step 1 phase 4g tail documentation` |
| Round 2 | Step 2 + Step 3 | `Phase 4h step 2-3 cross-market BS/IS/CF merge sheets + refresh wiring` |
| Round 3 | Step 4 | `Phase 4h step 4 B6 realtime toggle (cache原币 dump, B6 switch reformulates)` |
| Round 4 | Step 5 | `Phase 4h step 5 disk JSON cache layer (24h TTL)` |
| Round 5 | Step 6 + Step 7 | `Phase 4h step 6-7 stockanalysis 中概美股 fallback + closure` |

每轮 commit 后停下等 Planner review。Round 间互不依赖代码层(Step 4 / 5 / 6 都是独立子系统),但**回归测试要求每轮都跑** Phase 4f frozen 驱动 + Phase 4g 的 inspect。

---

## Step 1 — E. Phase 4g 小尾巴收账(纯 doc)

### 背景

Phase 4g 的 `db500e5 Fix US and HK indicator growth formulas` commit 顺手修了 2 个 Phase 4f 残留 bug:
1. `FindPriorSamePeriodStatementColumn` 加 ±31 天 fuzzy match,解决 AAPL fiscal year 漂移(2025-09-27 vs 2024-09-28)导致指标增长率公式找不到 prior period
2. 港股 BS/Income 加 lngYear-1 双年 fetch,让指标增长率公式有可比期数据

这 2 个改动当时严格说**违反 plan §⚠️ frozen 规则**(港股 fetch 是 Phase 4c/4d frozen)但被 reviewer 接受。Phase 4g 收口时 §V 已落账,本 Step 把它进一步文档化。

### 子任务 1A — README.md 更新

**1A.1** 在 README.md 现有「常见问题」或「实现说明」section 追加:

```markdown
### 美股 fiscal year fuzzy match
Apple / Microsoft 等公司的 fiscal year 不是 12 月 31 日(eg AAPL FY 通常是 9 月最后一个周六)。指标增长率公式查找 prior period 时,先做精确日期匹配,匹配不到再退到 ±31 天最近期间。该容错范围足以覆盖周末漂移和会计期口径微调,不会跨年误命中。

### 港股双年 fetch
为让港股指标增长率公式有可比期数据,『一键港股』在 BS/IS 抓数时会同时拉当年 + 前一年。CashFlow 不双拉(CF 增长率不在标准 18 指标里)。这会让港股抓数耗时 ~2x,但不影响 throttle(仍 1s/请求)。
```

### 子任务 1B — STATUS.md §V.1 注脚补充

**1B.1** STATUS §V.1 在『本阶段已完成』列表后追加:

```markdown
### V.4 附带修复(Phase 4g 期间发现并修复,不在原 plan 范围内)

- [bug] A 股 statement kind 编译错误:WriteWideTable 收到 "balance"/"profit"/"cash" 但 Phase 4f 期望 "BalanceSheet"/"Income"/"CashFlow"。已在 8341306 commit 加 hookKind 映射。
- [bug] AAPL fiscal year 漂移导致指标增长率公式找不到 prior period。已在 db500e5 commit 给 FindPriorSamePeriodStatementColumn 加 ±31 天 fuzzy match。
- [enhance] 港股 BS/IS 增长率公式因为只拉当年没数据。已在 db500e5 commit 加双年 fetch loop。
```

### 子任务 1C — `tools/inspect_phase4g_state.py` 增加注释

**1C.1** 在 inspect 脚本顶部 docstring 加一行说明它会同时验证 Phase 4g 主线 + db500e5 附带修复(eg AAPL 4 期数据存在 = fuzzy match 工作)。

### 验证

无需代码改动验证,doc 提交即完成。

### Generator 不要做

- ❌ 不要修改 db500e5 / 8341306 commit 的代码本身
- ❌ 不要新增独立测试驱动(已经在 Phase 4g inspect 里间接覆盖)

---

## Step 2 — A. 跨市场 BS/IS/CF 3 张合表 sheet + VBA 通用化

### 目标

把 Phase 4g `BuildCrossMarketIndicatorSheet` 抽象成 generic `BuildCrossMarketStatementSheet(statementKind)`,产生 3 张新 sheet:

- `跨市场_资产负债表`(从 `A股_资产负债表` / `美股_资产负债表` / `港股_资产负债表` / `韩股_资产负债表` 合并)
- `跨市场_利润表`(同上,源是 `*_利润表`)
- `跨市场_现金流量表`(源是 `*_现金流量表`)

这 3 张表的 layout 跟 `跨市场_指标表` 完全相同(R1 公司名 + R2 报告期 + R3+ 行项目;每个数据 cell 公式 cell-ref 到分市场表),**唯一不同是行项目集合**:

- 指标表(已有):18 项标准指标,4 市场对齐(因为分市场指标表已经统一过了)
- BS/IS/CF(本期):**行项目跨市场不对齐**,需要 mapping 策略

### 行项目对齐策略 — Codex 提案 + Reviewer 决策

实现时 Codex 必须做以下提案,以注释形式写在 `BuildCrossMarketStatementSheet` 上方:

```vba
' ========== Phase 4h Step 2: 跨市场 BS/IS/CF 行项目 mapping 提案 ==========
' 现状: 4 张分市场 X 表的行项目通过 IndicatorsByInsertion 收集,顺序 = 第一次出现顺序
'       4 市场行项目集合 各自不同 (中美 GAAP / IFRS / K-IFRS 差异)
' 提案 (Codex 实现时根据实际 dump 调整):
'   方案 P1: 严格对齐 — 只展示 4 市场都有的字段 (≈ 公共子集)
'   方案 P2: 全展示并集 — 把 4 市场所有字段都列出来,某市场没有就空白
'   方案 P3: 分组展示 — 按"全 4 市场都有 / 仅 N 市场有"分块,块内按 BS 标准结构排序
' Codex 当前实现选: <P1/P2/P3>, 理由: <实际 4 市场字段交集 / 并集大小>
' ========================================================================
```

Reviewer 收到 Round 2 commit 后会基于实际 dump 决定是否接受 / 要求换方案。

### 子任务 2A — `build_template.py` + `install_modules.py` 新增 3 张 sheet 模板

**2A.1** `build_template.py` 复用 Phase 4g 的 `build_cross_market_indicator_sheet` 函数体,extract 成 `_build_cross_market_statement_sheet(ws, statement_label)`,然后:

```python
# main() 末尾
for label in ("资产负债表", "利润表", "现金流量表"):
    ws_x = wb.create_sheet(f"跨市场_{label}")
    _build_cross_market_statement_sheet(ws_x, label)
```

**2A.2** `install_modules.py`:
- 复用 `_make_cross_market_indicator_sheet` 抽象成 `_make_cross_market_statement_sheet(wb, name, kind)`,kind 用于 freeze pane / col widths(BS/IS/CF 跟 Indicator 一样 D3 freeze)
- `ensure_market_sheets` 加 idempotent install 4 张(注意:`跨市场_指标表` 已经在 Phase 4g 装好,不重装)
- `reorder_report_sheets` 的 `desired_order` 调整,4 张跨市场表放在一起,顺序: BS / IS / CF / Indicator

### 子任务 2B — VBA `BuildCrossMarketStatementSheet` 通用化

**2B.1** 在 `模块_工具函数.bas` 的 `BuildCrossMarketIndicatorSheet` 之后新增:

```vba
' --------- Phase 4h Step 2: 把 4 张分市场 BS/IS/CF 合并展示到对应『跨市场_<报表>』 ---------
'   - statementKind: "BalanceSheet" / "Income" / "CashFlow"
'   - 行项目集合: 从 4 张分市场表的 IndicatorsByInsertion 顺序合并 (Codex 在此实现时选 P1/P2/P3 之一)
'   - 列布局: 横向铺 公司×报告期 perCompanyPeriods=True (跟指标表一致)
Public Sub BuildCrossMarketStatementSheet(ByVal statementKind As String)
    Dim targetSheet As String, statementLabel As String
    Select Case UCase$(Trim$(statementKind))
        Case "BALANCESHEET":
            targetSheet = "跨市场_资产负债表": statementLabel = "资产负债表"
        Case "INCOME":
            targetSheet = "跨市场_利润表":     statementLabel = "利润表"
        Case "CASHFLOW":
            targetSheet = "跨市场_现金流量表": statementLabel = "现金流量表"
        Case Else:
            Err.Raise vbObjectError + 581, "BuildCrossMarketStatementSheet", _
                "未知 statementKind: " & statementKind
    End Select
    
    ' ... 复用 BuildCrossMarketIndicatorSheet 主体逻辑, 把 IndicatorSheetName 替换成 StatementSheetName
    ' ... 行项目集合逻辑根据 mapping 提案 P1/P2/P3 实现
End Sub

Private Function MarketStatementSheetName(ByVal market As String, ByVal statementLabel As String) As String
    Select Case UCase$(Trim$(market))
        Case "A":  MarketStatementSheetName = "A股_" & statementLabel
        Case "US": MarketStatementSheetName = "美股_" & statementLabel
        Case "HK": MarketStatementSheetName = "港股_" & statementLabel
        Case "KR": MarketStatementSheetName = "韩股_" & statementLabel
    End Select
End Function
```

**2B.2** 重构 `BuildCrossMarketIndicatorSheet`:抽出公共 helper(eg `WriteCrossMarketHeaders`、`WriteCrossMarketRowFormulas`),让 indicator / BS / IS / CF 4 个 caller 都用同一套底层。**保持 Phase 4g `BuildCrossMarketIndicatorSheet` 的对外 API 不变**(Phase 4g 验证脚本已 freeze 它的 OnAction)。

### 验证(Generator 自测)

- 装完模板,4 张『跨市场_*』sheet 都存在,Tab 顺序 BS / IS / CF / Indicator
- 在样本池填示例 4 公司(300866 / AAPL / 00700 / 005930),跑 `一键全抓 4 市场`
- 4 张跨市场表的 R1 公司名 + R2 报告期 都写入,R3+ 行项目数量 ≈ Codex 提案的 P1/P2/P3 大小
- 抽 3 个 cell 检查公式 = `'A股_资产负债表'!D3` 形式

### Generator 不要做

- ❌ 不要动 `BuildStandardIndicatorSheet`(分市场指标表是 source of truth,合表只 lookup)
- ❌ 不要动现有 4 市场 fetch 模块的字段映射
- ❌ 不要在合表里再做一次 RMB 换算(分市场表 Phase 4f 已经换好)
- ❌ 不要把 4 张合表压到 1 张 sheet(用户明确选了 4 张子表布局)

---

## Step 3 — A. 4 张合表统一刷新逻辑

### 目标

让用户能:
- 一键全抓 4 市场后**自动刷** 4 张合表
- 单独点 1 个『合并跨市场全部 4 表』按钮统一刷
- 单独 4 个按钮 each 刷 1 张(BS / IS / CF / Indicator)

### 子任务 3A — VBA 总入口加汇总刷新 Sub

**3A.1** 在 `模块_工具函数.bas` 新增:

```vba
Public Sub BuildAllCrossMarketSheets()
    BuildCrossMarketStatementSheet "BalanceSheet"
    BuildCrossMarketStatementSheet "Income"
    BuildCrossMarketStatementSheet "CashFlow"
    BuildCrossMarketIndicatorSheet
End Sub
```

**3A.2** 修改 `模块_总入口.bas` `一键全抓` 末尾(Phase 4g 现有 `BuildCrossMarketIndicatorSheet` 调用):

```vba
' Phase 4g Step 2 (modified Phase 4h Step 3): 一键全抓后自动刷新 4 张跨市场表
On Error Resume Next
BuildAllCrossMarketSheets    ' 替换 Phase 4g 的 BuildCrossMarketIndicatorSheet
Err.Clear
On Error GoTo CleanUp
```

### 子任务 3B — 5 个按钮(install_modules.py BUTTONS list)

**3B.1** Q5:Q7 现有的『合并跨市场指标表』按钮**保留**(Phase 4g 已验收,不动 OnAction);新增 4 个独立按钮 + 1 个汇总按钮:

```python
# Phase 4g 已有 (不动): ("BtnBuildCrossInd", "合并跨市场指标表", ..., "Q5:Q7", ...)
# Phase 4h 新增:
("BtnBuildCrossBS",  "合并跨市场资产负债表",  "模块_工具函数.BuildCrossMarketBalanceSheetWrapper",  "S5:S7", PRIMARY_FILL, PRIMARY_FG, 11, True),
("BtnBuildCrossIS",  "合并跨市场利润表",      "模块_工具函数.BuildCrossMarketIncomeWrapper",        "S8:S10", PRIMARY_FILL, PRIMARY_FG, 11, True),
("BtnBuildCrossCF",  "合并跨市场现金流量表",  "模块_工具函数.BuildCrossMarketCashFlowWrapper",      "S11:S13", PRIMARY_FILL, PRIMARY_FG, 11, True),
("BtnBuildCrossAll", "合并 4 张跨市场表",     "模块_工具函数.BuildAllCrossMarketSheets",            "S1:S3", PRIMARY_FILL, PRIMARY_FG, 12, True),
```

**3B.2** 加 wrapper Subs(因为 BuildCrossMarketStatementSheet 是带参数的,Excel Shape OnAction 不能传字符串参数):

```vba
Public Sub BuildCrossMarketBalanceSheetWrapper()
    BuildCrossMarketStatementSheet "BalanceSheet"
End Sub
Public Sub BuildCrossMarketIncomeWrapper()
    BuildCrossMarketStatementSheet "Income"
End Sub
Public Sub BuildCrossMarketCashFlowWrapper()
    BuildCrossMarketStatementSheet "CashFlow"
End Sub
```

### 验证

- 装完看到 6 个 Q/S 列按钮(`BtnRunAll`, `BtnBuildCrossInd`, `BtnHideAll`, `BtnBuildCrossBS/IS/CF/All`)
- 跑一键全抓 4 市场,4 张跨市场表都自动写入
- 单独点『合并跨市场资产负债表』,只刷 BS,其他 3 张不动

### Generator 不要做

- ❌ 不要改 Phase 4g `BtnBuildCrossInd` 按钮的 OnAction
- ❌ 不要在新按钮里塞 MsgBox(完成无声操作)

---

## Step 4 — B. B6 原币↔RMB 实时 toggle

### 目标

用户切换 B6『原币』↔『统一RMB』后,**不需要重抓**,数值立即重算并刷新所有展示区(分市场 BS/IS/CF/Indicator + 跨市场合表)。当前痛点:每次切换要重点 4 个『一键 X 股』再花 ~3min 重抓全部 4 市场。

### 实现思路

**核心想法**:把 WriteWideTable 改成两步:
- Step 1: 写**原币 raw dump**到 sheet 的「隐藏区」(eg Row 100+,字体 0pt 或 hidden cells)
- Step 2: 写**展示区公式**(R3+ 用户能看到的区域),公式 cell-ref 到 raw dump,**根据 B6 决定是否乘汇率**

B6 切换后:
- `Worksheet_Change` 事件监听 B6 变化
- B6 变化时,触发 `RebuildDisplayFormulas` 遍历 4 市场 BS/IS/CF/Indicator,把展示区公式重写(只动 ~20-50 行公式,不重抓数)
- ~1s 完成

### 子任务 4A — `WriteWideTable` 重构成两层

**4A.1** 重构现有 `WriteWideTable`:

```vba
Public Sub WriteWideTable(...)
    ' 原有参数全保留, 新增:
    '   useRawDumpLayer: Boolean = True (Phase 4h 默认开启, Phase 4f 行为可通过 False 关闭做回归对比)
    
    If useRawDumpLayer Then
        ' Step 1: 写原币 raw dump 到隐藏区
        Call WriteRawDumpZone(ws, arrCodes, dictData, ...)    ' Row 200+, 字体 0pt
        
        ' Step 2: 写展示区公式 (Row 1-N), cell-ref 到 raw dump
        Call WriteDisplayFormulas(ws, arrCodes, displayMode, dictReportingCurrency, statementKind, ...)
    Else
        ' Phase 4f legacy 路径: 直接写 Value (用于回归对比)
        Call WriteValuesLegacy(ws, ...)
    End If
End Sub
```

**4A.2** 新增 `WriteRawDumpZone`:把所有 BS/IS/CF/Indicator 的原币 raw 数据写到 sheet 的 Row 200+(避开 Row 1-100 用户视觉区),字体颜色 = 背景色 / 行高 = 0(用户看不到但 Excel 公式能 ref)。

**4A.3** 新增 `WriteDisplayFormulas`:展示区 cell 公式形式为 `=IF(样本池!$B$6="原币", _RawDumpCell, _RawDumpCell * 汇率!$B$N)`。这样 B6 一变 Excel 自动重算,不需要 VBA 跑。

### 子任务 4B — 样本池 `Worksheet_Change` 事件钩子(可选 fast path)

**4B.1** 如果 4A 实现成功(纯公式 toggle),就不需要 Worksheet_Change。如果 4A 实现遇到性能问题(eg 公式太多导致 Excel 卡),改用 Worksheet_Change 监听 B6 变化,触发 `RebuildDisplayFormulas` Sub 主动批量改写。

### 子任务 4C — 跨市场合表 inherit toggle

**4C.1** Phase 4g `BuildCrossMarketIndicatorSheet` + Phase 4h Step 2 `BuildCrossMarketStatementSheet` 已经是 cell-ref 形式(eg `=A股_指标表!D3`),所以分市场表 toggle 后跨市场表自动刷,**不需要额外改动**。验证一下即可。

### 子任务 4D — 性能 budget

- 4A 完整改写后,B6 切换响应时间 < 2s(用户感知"立即")
- 如果超 2s,改用 Worksheet_Change 路径

### 验证

- 跑一遍 4 市场抓数(B6=原币),检查分市场 + 跨市场表数据正确
- B6 → 统一RMB,**不点任何按钮**,1-2s 内所有数值变成 RMB 等价值
- B6 → 原币,数值复原
- 跨 5 次 toggle,数据稳定无 drift

### Generator 不要做

- ❌ 不要把 raw dump zone 暴露给用户(Row 200+ 隐藏)
- ❌ 不要在 toggle 时弹 MsgBox / Application.StatusBar 长期更新
- ❌ 不要触碰 Phase 4f `RefreshA1CurrencyComment` 主体(A1 注释逻辑已经依赖 displayMode 自动刷)
- ❌ 不要重写汇率获取(GetFxRate / 模块_抓汇率 frozen)

### 风险点

- raw dump zone 占用 sheet 行号大量空间(eg 18 指标 × 4 公司 × 5 期 = 360 cell × 4 张表 = 1440 cell,可控)
- 公式 cross-sheet ref 在 sheet 改名时会失效 — 4 张跨市场表 sheet 名 frozen 后保持不变 OK
- Excel 1M 行限制 远超本期需求,无风险

---

## Step 5 — D. 磁盘 JSON 缓存层(降低重复 HTTP)

### 目标

每次跑数,雪球 / EDGAR / stockanalysis 的原始 HTTP 响应落到本地 `.cache/` 目录,**24 小时内**重复跑同一公司同期数据时直接读本地 JSON,跳过 HTTP 请求。**这是降低开发调试时网络流量的本地工具,不增加抓数频次**。

### 实现思路

在 PowerShell shell-out 那一层(Phase 4f Step 2 已有的 `RunWinHttpGet` / 同等 helper)插一层:

```python
def cached_get(url, cache_key, ttl_hours=24):
    cache_path = Path(".cache") / f"{hash(cache_key)}.json"
    if cache_path.exists() and (now - cache_path.mtime).hours < ttl_hours:
        return cache_path.read_text()    # cache hit, no HTTP
    
    response = http_get(url)             # cache miss, real HTTP
    cache_path.write_text(response)
    return response
```

VBA 层调用方完全不变,只是 PowerShell helper 内部多一步。

### 子任务 5A — PowerShell shell-out 加 cache layer

**5A.1** 修改 `tools/winhttp_get.ps1`(Phase 4f Step 2 已有,frozen 但可加新参数):

```powershell
# 加 -CacheKey 和 -CacheTtlHours 参数
# 默认 -CacheKey "" 关闭缓存 (向后兼容 Phase 4f 行为)
param(
    [string]$Url,
    [string]$CacheKey = "",
    [int]$CacheTtlHours = 24,
    [string]$CacheDir = ".cache"
)

if ($CacheKey -ne "") {
    $hash = (Get-FileHash -Algorithm MD5 -InputObject ([System.Text.Encoding]::UTF8.GetBytes($CacheKey))).Hash
    $cachePath = Join-Path $CacheDir "$hash.json"
    if ((Test-Path $cachePath) -and ((Get-Date) - (Get-Item $cachePath).LastWriteTime).TotalHours -lt $CacheTtlHours) {
        Get-Content $cachePath -Raw
        exit 0
    }
}

# ... 原有 HTTP 请求逻辑
$response = ... # 跟 Phase 4f 一致

if ($CacheKey -ne "") {
    New-Item -ItemType Directory -Force -Path $CacheDir | Out-Null
    Set-Content -Path $cachePath -Value $response
}

Write-Output $response
```

### 子任务 5B — VBA 调用层加 CacheKey

**5B.1** `模块_工具函数.bas` 现有 `RunWinHttpGet` 函数(或同等),加 Optional cacheKey 参数:

```vba
Public Function RunWinHttpGet(ByVal url As String, _
                              Optional ByVal cacheKey As String = "") As String
    ' 把 cacheKey 透传给 PowerShell -CacheKey 参数
End Function
```

**5B.2** 雪球 / EDGAR / stockanalysis 的各 fetch 模块在调用 `RunWinHttpGet` 时,把 cacheKey 设成 `"<source>_<ticker>_<period>"` 形式(eg `"xueqiu_HK_00700_2024Q4_balance"`)。

### 子任务 5C — `.cache/` 加入 .gitignore

**5C.1** `.gitignore` 追加 `.cache/`,不入库。

### 子任务 5D — UX:样本池新增『清空缓存』按钮(可选)

**5D.1** Q14 加一个小按钮『清空缓存』,调用 `Public Sub ClearLocalCache()`,删除 `.cache/` 目录(用 PowerShell 调用)。

### 验证

- 第一次跑数 → 正常耗时(eg 3min)+ `.cache/` 目录生成 ~30 个 JSON
- 1 小时内立即再跑同一组数据 → 耗时 < 30s(全部 cache hit)
- 25 小时后跑 → 耗时回到 3min(TTL 过期 cache miss 重抓)

### Generator 不要做

- ❌ 不要 cache 雪球 cookie(cookie 是用户敏感信息)
- ❌ 不要把 cache TTL 设超过 7 天(财报数据可能更新)
- ❌ 不要 cache 失败响应(eg HTTP 4xx/5xx,只 cache 200)

### 风险点

- Cache key 设计要 stable:同一公司同期数据多次跑必须命中同一 key
- Windows 文件系统 path 长度 260 限制:hash 化 cache key 避免长 URL 入文件名

---

## Step 6 — C. stockanalysis 中概美股 fallback 落地

### 背景

Phase 4g Step 4 已对 BABA / JD / PDD 收入表做了覆盖度调研(3/3 HTTP 200,但只覆盖 income statement,BS/CF 在子页面)。本 Step **完成 BS/CF 子页面抓样 + 字段映射 + 接入 VBA 作为雪球失效时的备用路径**。

**重要语境**:这是 Phase 4d 已有数据源(stockanalysis.com 韩股路径)的同站点 generalize,不是新数据源。维持 Phase 4d 同样的单线程 + 2s 间隔节奏,默认 fallback hook 关闭(B7 cookie cell 旁边新增 B8『中概美股 fallback』开关,默认空 = 关)。

### 子任务 6A — 子页面 sample 收集(再跑 Phase 4g probe)

**6A.1** 复用 `tools/probe_4a_stockanalysis_coverage.py`,加 BS/CF 路径:

```python
TARGETS_BS_CF = [
    ("BABA_balance",   "https://stockanalysis.com/stocks/baba/financials/balance-sheet/"),
    ("BABA_cashflow",  "https://stockanalysis.com/stocks/baba/financials/cash-flow-statement/"),
    ("JD_balance",     "https://stockanalysis.com/stocks/jd/financials/balance-sheet/"),
    ("JD_cashflow",    "https://stockanalysis.com/stocks/jd/financials/cash-flow-statement/"),
    ("PDD_balance",    "https://stockanalysis.com/stocks/pdd/financials/balance-sheet/"),
    ("PDD_cashflow",   "https://stockanalysis.com/stocks/pdd/financials/cash-flow-statement/"),
]
# 同样 2s 间隔, sequential, identity encoding, no cookie, 6 个 URL 写死
```

跑一次,sample HTML 落地 `samples/stockanalysis_<ticker>_<statement>.html`,加入 `.gitignore`(已有规则覆盖)。

### 子任务 6B — HTML 表格解析复用 Phase 4d 韩股代码

**6B.1** Phase 4d 已经有 stockanalysis HTML 表格解析(韩股路径用),文件 `modules/模块_抓韩股*.bas` 或 `tools/parse_stockanalysis.py`(具体文件名 grep 确认)。Phase 4h **复用其解析 logic**,只是字段映射换成中概美股 taxonomy。

**6B.2** 新增 `tools/parse_stockanalysis_us.py`(或 VBA `模块_抓中概美股fallback.bas`),读 sample HTML → 提取 BS/CF/IS 字段 → 输出 JSON 同 EDGAR 路径格式。

### 子任务 6C — VBA fallback hook(默认关)

**6C.1** 样本池 B8 新增『中概美股 fallback 开关』(下拉选项: `关` / `开`)。装表 idempotent install。

**6C.2** `模块_抓美股*.bas` 在主路径(EDGAR + 雪球 fallback)失败后,如果 B8 = `开`,再退到 stockanalysis fallback:

```vba
' 主路径失败后追加 stockanalysis fallback
If ReadStockAnalysisFallbackEnabled() = "开" Then
    On Error Resume Next
    FetchUSFromStockAnalysis strCode, strKind, ...
    On Error GoTo CleanUp
End If
```

**6C.3** 诊断 sheet 新增 fallback 触发记录:`数据源` 列写 `stockanalysis (fallback)` 区分主路径。

### 子任务 6D — README + STATUS 文档化

**6D.1** README 新增段落:中概美股 fallback 默认关,启用方式 + 局限性(只覆盖 BABA/JD/PDD 测试过,其他中概股可能字段不齐)。

### 验证

- B8 = `关`:中概美股抓数行为不变(主路径)
- B8 = `开`,故意把 B5 cookie 设为无效,跑 `一键美股` BABA → 主路径失败 → 走 stockanalysis fallback → 数据落表
- 诊断 sheet 显示 `stockanalysis (fallback)` 来源

### Generator 不要做

- ❌ 不要把 stockanalysis 设成主路径(用户决策:`继续用雪球 / EDGAR 主路径`)
- ❌ 不要扩展 fallback 到非中概美股(本期只覆盖 3 ticker 测试集)
- ❌ 不要重写 Phase 4d 韩股 stockanalysis 解析(那是 frozen 的)
- ❌ 不要并发抓数(单线程 sequential + 2s 间隔)
- ❌ 不要把抓样 HTML 提交 git(.gitignore 已有规则)

### 风险点

- stockanalysis 站点 反爬:Phase 4g 调研 6 个 URL 都返回 200(港股 4 个除外是 404,不算反爬)。维持 Chrome UA + identity encoding + 2s 间隔 风险可控。
- 中概美股字段单位 / 币种:stockanalysis 默认 USD,RMB 换算走 Phase 4f hook(报告币种 = USD)
- BABA/JD/PDD 之外中概股(eg BIDU、TCOM)字段可能不齐 — 本期不保证

---

## Step 7 — 端到端回归 + STATUS §W 收口 + plan v2

### 子任务 7A — 跑现有回归驱动(每轮都已跑过,本步是终局确认)

```bash
cd "VBA Captor"
py tools/test_fx_live.py --skip-install         # Phase 4f Step 2 frozen (5/5)
py -u tools/diff_phase4f_step3_lite.py          # Phase 4f Step 3-5 frozen (smoke)
py -u tools/inspect_phase4g_state.py            # Phase 4g 跨市场指标表 + hide-tab + 诊断 11 列
```

期望:**全部 PASS**。任一退化触发 ⚠️ 联系 Planner 触发条件。

### 子任务 7B — Phase 4h 新建 inspect 驱动

**7B.1** 新增 `tools/inspect_phase4h_state.py`,dump:

- 4 张跨市场表 sheet 存在 + R1 公司名 + R2 报告期 + Row 3-5 头 3 行公式 + 计算值(Step 2)
- 5 个新按钮位置 + caption + OnAction(Step 3)
- B6 toggle:跑数后切 B6 立即看分市场表 D5 cell 数值变化(Step 4)
- `.cache/` 目录存在,清空后重跑命中率 = 0%,1 分钟内重跑命中率 ≈ 100%(Step 5)
- B8 = 开 + 故意失效 cookie:中概美股诊断 sheet 出现 `stockanalysis (fallback)`(Step 6)

### 子任务 7C — `STATUS.md` §W 收口

**7C.1** 模仿 §V 格式追加 §W:

```markdown
## W. Phase 4h 收口: 跨市场全合表 + 实时 toggle + 缓存 + fallback + 4g 收账

执行依据: `PHASE_4H_PLAN.md` v1。状态: ✅ Codex 已实现 5 件主线 + 1 件文档,通过本地回归 + 无网络验收;Phase 4h 全期闭环。

### W.1 本阶段已完成
- [Step 1] Phase 4g 小尾巴文档化: README + STATUS §V.4
- [Step 2] 跨市场 BS/IS/CF 3 张合表 + VBA BuildCrossMarketStatementSheet 通用化, 行项目对齐方案 = <Codex 实际选择>
- [Step 3] 4 张合表统一刷新: BuildAllCrossMarketSheets + 5 个新按钮
- [Step 4] B6 实时 toggle: WriteWideTable 双层 (raw dump + display formula), 切换 ~1s 完成
- [Step 5] 磁盘 JSON 缓存: 24h TTL, 第二次跑数耗时降到 ~30s
- [Step 6] stockanalysis 中概美股 fallback: 默认关, B8 开关启用; 通过 BABA/JD/PDD 验证

### W.2 验证结果
[7A 回归 + 7B inspect 结果]

### W.3 已知边界
- 跨市场合表行项目对齐: <P1/P2/P3 决策实际效果>
- B6 实时 toggle 性能: 4 公司 × 5 期场景下 ~1s; 超过 10 公司可能要降级
- 缓存只覆盖 .cache/ 内 24h 数据, cookie / Excel state 不缓存
- stockanalysis fallback 仅覆盖 BABA/JD/PDD; 其他中概股字段可能不齐
```

**7C.2** `PHASE_4H_PLAN.md` 状态行 v1 → v2,标记 `✅ Phase 4h 全期闭环`。

### 子任务 7D — `README.md` 全部新功能使用说明

**7D.1** 4 张跨市场表使用说明、B6 实时 toggle 提示、缓存目录说明、B8 fallback 开关说明。

---

## ⚠️ 全 Phase 严禁动的东西

| 文件/区域 | 原因 |
|---|---|
| `modules/模块_抓汇率.bas` | Phase 4f Step 2 frozen |
| `modules/模块_工具函数.bas` line 535-820(HTTP + GetFxRate 区) | Phase 4f frozen |
| `modules/模块_工具函数.bas` `BuildCrossMarketIndicatorSheet` 主体 + 公开 API | Phase 4g frozen,只能扩展底层 helper,不能改对外 Sub 名/参数 |
| `modules/模块_工具函数.bas` `WriteWideTable` 的 `dictReportingCurrency` / `statementKind` 参数语义 | Phase 4f Step 3-5 frozen,Step 4 重构只能加新参数 |
| `tools/test_fx_live.py` / `tools/diff_phase4f_step3_lite.py` / `tools/inspect_phase4g_state.py` | Phase 4f / 4g 验证驱动 frozen |
| Phase 4d 韩股 stockanalysis 解析代码 | Step 6 只复用其 helper,不改 |
| Phase 4f `RefreshA1CurrencyComment` | Step 4 toggle 后 A1 注释自动刷,不动主体 |
| 现有 4 市场 fetch 模块的 fetch 主路径(雪球 / EDGAR) | Step 6 fallback 是追加,主路径不变 |

## ⚠️ 联系 Planner 触发条件

- Step 2 跨市场 BS/IS/CF 行项目交集 < 5 个(说明 4 市场字段命名差异太大,严格对齐方案不可行,需要重新决策)
- Step 4 B6 toggle 响应时间 > 5s(说明纯公式路径性能不达标,需要决定是否降级 Worksheet_Change 主动重写)
- Step 5 缓存 cache key 设计冲突(eg 不同抓数路径用同 key 导致数据 corrupt)
- Step 6 stockanalysis 中概美股 BS/CF 子页面 HTTP 失败率 > 30%(说明 URL 模式不稳定,fallback 不可靠)
- Step 7 任一回归驱动退化(eg test_fx_live 跑出 < 5/5)

## ⚠️ Codex 安全分类器友好的 framing

Step 4(B6 toggle)/ Step 5(缓存)/ Step 6(stockanalysis fallback)涉及网络 + 数据持久化,**generator prompt 必须明确**:

- **不引入新数据源**:Step 6 是 Phase 4d 已有 stockanalysis 韩股路径的同站点 generalize
- **不增加抓数频次 / 并发**:Step 5 缓存是为了**降低**重复 HTTP,Step 6 fallback 维持 Phase 4d 单线程 + 2s 间隔
- **不做自动化扩面**:Step 6 fallback 测试集只有 BABA/JD/PDD 3 ticker,B8 开关默认关
- **数据用途**:个人 audit / research,不是商业自动化采集
- **抓样规模**:Step 6 一次性 6 个子页面 sample 收集,跟 Phase 4g Step 4 同等规模

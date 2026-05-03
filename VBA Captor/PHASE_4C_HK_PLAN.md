# Phase 4c: 港股 (HK) 重启 — 走雪球 HK API

> **版本**: v2(2026-05-03,Step 1 探查后调整)
> **状态**: 🚧 Step 3 + Step 4 已实现,等待 Claude Code review
> **作者**: Claude(planner)+ Codex(executor)
> **背景**: Phase 4b-14a 美股已闭环。港股是用户记忆里"A股 / 美股 / 港股 / 韩股"4 市场目标的第三块。早期 Phase 4a 走新浪 HK 失败已废弃,本期改走**雪球 HK API**,复用 Phase 4b-5 起雪球 fallback 的成熟框架。

## v2 修订纪要(Step 1 探查发现的 3 个 plan 假设破产点)

| # | v1 假设 | v2 调整 | 来源 |
|---|---|---|---|
| 1 | 单位 = 百万港元 (HKD) | **不强转币种**,sheet A1 改"单位:百万(各家公司报告币种,见 港股_抓取诊断 Unit 列)" | Step 1 发现:腾讯雪球返回 `data.currency=CNY`,港股互联网公司多数报告币种是 RMB,部分是 HKD/USD,**写死币种会误导**;诊断 sheet `Unit` 列已能承载币种信息 |
| 2 | HK 字段名跟美股雪球类似 | **完全独立 field map** — HK 雪球字段是中文项首字母缩写(`ta` / `tlia` / `tto` / `nocf` / `gp`),与美股 (`total_assets` / `revenue`) 完全异构 | Step 1 发现:HK BS 41 字段 vs US POM BS 54 字段,公共 key 仅 5 个(全是 metadata `ed/sd/report_date/...`)|
| 3 | 季度过滤沿用 `report_annual` + `fp` | **改用 `month_num` + `ed` 后缀** — HK 雪球没有 `report_annual` / `report_type_code`,只有 `sd/ed/report_name/month_num` | Step 1 发现:`first record = {"report_name": "2025年报", "month_num": 12, "ed": "2025-12-31"}`,无 fy/fp 字段 |

## Context

**用户当前状态**: A股 / 美股已稳定;港股 sheet 模板早就在工作簿里(空的)、市场识别(`ResolveMarket` 5 位数字 → HK)早就支持,但**无抓数模块**。雪球 HK 接口在浏览器里能查(已确认 https://xueqiu.com/snowman/S/00700/detail#/ZCFZB),只是没接进 VBA。

**为什么走雪球而非新浪/东财/HKEX**:
- 新浪 HK 财报数据稀疏(Phase 4a 跳过的根因)
- HKEX 自家只提供 PDF 年报,无结构化 API
- 东方财富 HK 接口未验证,字段映射要重做
- **雪球 HK** 已有 cookie + HTTP wrapper + JSON parser + 诊断框架(美股 fallback 用了大半年),复用成本最低

## 已锁定的决策(用户已确认)

| 决策 | 选项 |
|---|---|
| 数据源 | **雪球 HK API**(`stock.xueqiu.com/v5/stock/finance/hk/{balance|income|cash_flow|indicator}.json`) |
| 入口架构 | **新写 `RunHKStatement`** — 独立模块,不复用 `RunUSStatement`(避免污染美股代码 + 没有三级递进需要,只走雪球) |
| 测试公司 | **互联网 H 股 × 4**: `00700 腾讯` / `09988 阿里 H` / `01024 快手` / `03690 美团` |
| 货币单位 | ~~百万港元~~ → **不强转币种**(v2 修订 #1)。Sheet 数据列 = 雪球原值 / 1,000,000(scale=1e6);A1 注释 = "单位:百万(各家公司报告币种,见 港股_抓取诊断 Unit 列)";诊断 sheet `Unit` 列填 `data.currency` (CNY/HKD/USD)。**不做汇率换算**(留 Phase G 跨市场对标做)|
| Sheet 命名 | `港股_资产负债表` / `港股_利润表` / `港股_现金流量表` / `港股_指标表`(对称 A股_/美股_) |
| 指标体系 | 沿用 18 个标准指标(`BuildStandardIndicatorSheet "HK"`)— 不抓雪球的原始指标表 |
| 诊断 sheet | **新建 `港股_抓取诊断`**(独立 sheet,不与 `美股_抓取诊断` 混)— 列结构相同 10 列 |
| 一键全抓 | 升级到 12 张表(A 股 4 + 美股 4 + 港股 4),顺序追加,不打散现有 8 张 |

## 实施方案(分步,每步建议 Codex 提交一次,我 review 一次)

### Step 1 — 探查雪球 HK API ✅ 已完成(2026-05-03)

**实际执行结果**(`samples/HK_API_PROBE.md`):
- ✅ 4 endpoint 全部 HTTP 200 + `error_code=0` + `data.list` len 8(`00700` 格式;`700` 不带前导 0 → list 空)
- ⚠️ Tencent `data.currency=CNY` — 不是预期的 HKD,触发 v2 修订 #1
- ⚠️ HK 字段名是缩写 (`ta/tlia/tto/nocf/gp/cceq/...`),与美股雪球完全异构 — 触发 v2 修订 #2
- ⚠️ HK records 缺 `report_annual` / `report_type_code`,只有 `sd/ed/report_name/month_num` — 触发 v2 修订 #3
- ✅ Codex 提供了字段对照表(BS 9 项 / IS 7 项 / CF 7 项),作为 Step 2 实施基线

Step 1 → Step 2 review 通过,Codex 按 v2 修订 + 下面 Step 2 强化指引执行。

### Step 2 — 字段映射 + 主流程模块(v2 调整后) ✅ 已实现(2026-05-03)

**新建** `modules/模块_抓港股财报.bas`,核心函数:

```vba
Public Sub RunHKStatement(strKind, targetSheet, conceptMap, maxPeriods)
    ' 类似 RunUSStatement 但只走雪球, 无 EDGAR
    ' Step 2 临时使用港股模块内私有诊断 writer,确保不污染美股诊断
    ' Step 4 再统一切到 g_diagnosticSheetName = "港股_抓取诊断"
    ' 创建 collDiagRows, 调 FetchHKFromXueqiu, 末尾 WriteHKDiagnosticForKind
End Sub

Private Sub FetchHKFromXueqiu(strTicker, strKind, conceptMap, collDiagRows, ...)
    ' 直接调 https://stock.xueqiu.com/v5/stock/finance/hk/{kind}.json
    ' 字段映射用 XueqiuFieldMapForKindHK (HK 专用)
    ' 不走 ifrs-full, 不走 us-gaap
    ' 单位 = data.currency (CNY/HKD/USD), 写到诊断行 Unit 列
End Sub

Private Function XueqiuFieldMapForKindHK(strKind) As Object
    ' BS/IS/CF 各一组 Dictionary
    ' 字段名基于 Step 1 dump 的 HK 缩写 (ta/tlia/tto/nocf/gp/cceq/...)
End Function

Private Function MatchHKPeriod(periodEnd, monthNum, strQuarter, lngYear) As Boolean
    ' 类比 MatchXueqiuPeriod, 但用 month_num + ed 后缀, 不用 report_annual
    ' 全部 → 通过
    ' Q1 → monthNum=3 AND ed 月日="-03-31"
    ' Q2 → monthNum=6 AND ed 月日="-06-30"  (港股 H1 半年报落这里)
    ' Q3 → monthNum=9 AND ed 月日="-09-30"
    ' Q4 → monthNum=12 (年报, ed 月日按公司财年 03-31/06-30/12-31 都可能)
End Function
```

**conceptMap 形态**: 6-tuple `(cat, label, primary_xq_csv, "—", 1000000#, alt_xq_csv)`
- 第 4 槽位(unit)填 `"—"`(占位,实际单位由 `data.currency` 决定,写到诊断 Unit 列)
- 第 6 槽位(alt_csv)填备选雪球字段 CSV(若有)
- **建议给港股写专门的 `MapEntryXueqiuConcepts` helper**(语义清晰),不复用 `MapEntryUsGaapConcepts`

**核心字段(用于诊断"核心字段是否抓到",HK 没 fallback,只用于验证 + 写 status)**:
- BS: `Total assets` → `ta` / `Total liabilities` → `tlia` / `Total equity` → `teqy`
- IS: `Revenue` → `tto` / `Net income` → `ploashh`
- CF: `Cash from operations` → `nocf` / `Cash at end of period` → `cceqeyr`

**Step 1 字段对照表(v1 实施基线,直接用)**:

| 输出标签 | HK 字段 | 备选 |
|---|---|---|
| **BS** | | |
| Cash & equivalents | `cceq` | |
| Inventory | `inv` | |
| Total current assets | `ca` | |
| Total current liabilities | `clia` | |
| Total assets ⭐ | `ta` | |
| Total liabilities ⭐ | `tlia` | |
| Total equity ⭐ | `teqy` | |
| Long-term debt | `ltdt` | |
| Short-term debt | `stdt` | |
| **IS** | | |
| Revenue ⭐ | `tto` | |
| Gross profit | `gp` | |
| Operating income | `opeplo` | `opeploinclfincost` |
| Net income ⭐ | `ploashh` | |
| R&D expense | `rshdevexp` | |
| Selling expense | `slgdstexp` | |
| Admin expense | `admexp` | |
| **CF** | | |
| Cash from operations ⭐ | `nocf` | |
| Cash from investing | `ninvcf` | |
| Cash from financing | `nfcgcf` | |
| Cash at beginning | `cceqbegyr` | |
| Cash at end ⭐ | `cceqeyr` | |
| FX effect | `ncfdchexrateot` | |
| Capex / PPE | `fxdiodtinstr` | `rpafxdiodtinstr` |

⭐ = 核心字段(决定诊断 status)

**HK 专属变化**:
- 季度过滤靠 `month_num` + `ed` 后缀(v2 修订 #3)— 见 `MatchHKPeriod` 伪代码
- canonical entry grouping key 改用 `month_num + "|" + Left(ed,4)` 替代 `fy + "|" + fp`
- 财年差异:阿里 H(03 月)、腾讯/美团/快手(12 月)— canonical 用 `max(end)+min(start)` 自然处理
- 单位:**不强转**,sheet 数据列 = 原值/1e6,诊断 sheet Unit 列写 `data.currency`

### Step 3 — 4 张 thin wrapper 模块 ✅ 已实现(2026-05-03)

```
modules/模块_抓港股资产负债表.bas
modules/模块_抓港股利润表.bas
modules/模块_抓港股现金流量表.bas
modules/模块_抓港股指标表.bas
```

每个模块的 `Public Sub Main` 调 `RunHKStatement` + 自己的 conceptMap(BS/IS/CF)或 `BuildStandardIndicatorSheet "HK"`(Indicator)。

### Step 4 — 工具函数升级 ✅ 已实现(2026-05-03)

`modules/模块_工具函数.bas`:
- `BuildStandardIndicatorSheet` 第二参数 `market` 加分支 `"HK"` — 拷源表头从 `港股_资产负债表` 而非 `美股_资产负债表`
- `AppendStandardIndicators` 第二参数 `market` 加 `"HK"` 分支 — 公式引用 `'港股_资产负债表'!...` 等
- 新增 `Public g_diagnosticSheetName As String`(默认 "美股_抓取诊断", HK 主流程入口设为 "港股_抓取诊断"),让 `WriteDiagnosticForKind` / `EnsureDiagnosticSheet` / `ClearDiagnosticSheet` / `DeleteDiagnosticRowsForKind` 全部按这个 var 决定写到哪张 sheet

执行状态:✅ 已实现(2026-05-03)。

- `RunUSStatement` 入口显式设置 `g_diagnosticSheetName = "美股_抓取诊断"`。
- `RunHKStatement` 入口显式设置 `g_diagnosticSheetName = "港股_抓取诊断"`。
- 公共诊断函数已按 `CurrentDiagnosticSheetName()` 路由。
- Step 2 临时私有 HK 诊断 writer 已删除,HK 诊断写入改走通用 `WriteDiagnosticForKind`。
- `BuildStandardIndicatorSheet` / `AppendStandardIndicators` / `StandardRowMap` / `StandardTargetPeriodWanted` / `StandardDataStartCol` 已支持 `HK`。

### Step 5 — install_modules.py + build_template.py

`tools/build_template.py`:
- `main` 加创建 `港股_资产负债表` / `港股_利润表` / `港股_现金流量表` / `港股_指标表` 4 张 sheet(复用 `build_wide_table`)
- 加创建 `港股_抓取诊断` sheet(复用 `build_diagnostic_sheet`,只换 sheet 名 + Row 1 标题文字)

`tools/install_modules.py`:
- `ensure_us_sheets` → 重命名 `ensure_market_sheets`,加 4 张港股 sheet + `港股_抓取诊断`
- `_make_diagnostic_sheet` 接收 `name` 参数(已经是),复用
- `reorder_report_sheets` desired_order 加港股 4 张 + `港股_抓取诊断`(放美股 4 张 + 美股诊断之后)
- `BUTTONS` 加 4 个港股按钮(`BtnRunHKBalance` / `BtnRunHKProfit` / `BtnRunHKCash` / `BtnRunHKInd`),颜色用第三色(深绿 `#548235` / 白字 / 11pt,与 A 股蓝 + 美股红区分)
- 「使用说明」refresh 加港股一段

### Step 6 — `模块_总入口.一键全抓` 升到 12 张表

```vba
Public Sub 一键全抓(...)
    ' [1/12] - [4/12] A 股 4 表
    ' [5/12] - [8/12] 美股 4 表
    ' [9/12] - [12/12] 港股 4 表
    ' g_diagnosticAppendOnly = True 期间,清空两张诊断 sheet (美股_抓取诊断 + 港股_抓取诊断)
End Sub
```

`ClearDiagnosticSheet` 调 2 次(美股 + 港股),或重构成接收 sheet 名参数。

## 验收方案

### Test 1 — 4 互联网 H 股一键全抓
- 样本池: 00700 腾讯 / 09988 阿里 H / 01024 快手 / 03690 美团(可与 A 股 + 美股共存)
- 配置: `A2=2024, A4=Q4`
- 跑: 一键全抓
- **期望**:
  - 港股 4 张表都填好,4 家公司分别有数据列(美团 12 月末 / 阿里 H 03 月末)
  - 港股_抓取诊断 sheet 显示 BS Total assets / IS Revenue+Net income / CF Cash from operations 等核心字段为 `OK_XUEQIU`
  - 弹窗失败数 = 0
  - 12 张表的 cell 数据都非空(每家公司至少有 1 列)

### Test 2 — 财年差异验证
- 单跑 `更新港股资产负债表`
- 阿里 H(财年 3 月底)和 美团(财年 12 月底)报告期不一致
- **期望**: R2 报告期降序排列,两家公司各自有自己的财年末日期(03-31 / 12-31),不强对齐
- **附加验证(v2)**: 港股_抓取诊断 sheet 的 `Unit` 列每家公司可能不同(腾讯=CNY,其它待 Step 1 后续公司探测)

### Test 3 — 边界 case
- 故意填明显非法 HK 代码 `99999`(超出常规 HK 编号范围)
- **期望**: 失败但不中断,诊断 sheet 整行 MISSING

## 文件改动清单

| 文件 | 改动 | 责任 |
|---|---|---|
| `samples/xueqiu_HK_00700_*.json`(4 个) | dump 调试样本 | Codex Step 1 |
| `samples/HK_API_PROBE.md` | Step 1 探查报告 | Codex Step 1 |
| `modules/模块_抓港股财报.bas` | 新建 — RunHKStatement / FetchHKFromXueqiu / XueqiuFieldMapForKindHK | Codex Step 2 |
| `modules/模块_抓港股资产负债表.bas` | 新建 thin wrapper | Codex Step 3 |
| `modules/模块_抓港股利润表.bas` | 新建 | Codex Step 3 |
| `modules/模块_抓港股现金流量表.bas` | 新建 | Codex Step 3 |
| `modules/模块_抓港股指标表.bas` | 新建,调 `BuildStandardIndicatorSheet "HK"` | Codex Step 3 |
| `modules/模块_工具函数.bas` | `BuildStandardIndicatorSheet/AppendStandardIndicators` 加 HK 分支;新增 `g_diagnosticSheetName` | Codex Step 4 |
| `modules/模块_总入口.bas` | 一键全抓 升到 12 张表 + 双诊断 sheet 清空 | Codex Step 6 |
| `tools/build_template.py` | 加 5 张港股 sheet(4 表 + 1 诊断)模板 | Codex Step 5 |
| `tools/install_modules.py` | ensure_market_sheets 升级 + reorder + 加 4 个深绿按钮 + 使用说明 refresh | Codex Step 5 |

## 给 Codex 的执行指南

1. **Step 1 必须先做并暂停等 review** — 不要一条龙到底,API 探查结果决定后续字段映射的正确性
2. 字段映射保守:**核心字段**(BS Total assets/IS Revenue+Net income/CF CFO+End Cash)优先,其它字段先写 1 个候选就够,后续按诊断 sheet 推荐回填
3. **不要给港股加 ifrs-full / fuzzy 推荐**(雪球 snake_case 字段名 fuzzy 命中率低,且无 EDGAR 兜底必要)
4. 诊断 sheet 用 `g_diagnosticSheetName` 全局变量切换 — 这是本期最关键的解耦,**否则港股诊断会写错到美股 sheet**
5. 一键全抓加港股段时,**注意 `g_diagnosticAppendOnly` flag 的作用域** — 需要在港股段开始前重置 `ClearDiagnosticSheet`(港股版),否则会把美股诊断和港股诊断混在一张
6. 测试时建议 cookie 提前刷新一次(雪球 cookie 1 个月过期)
7. 跑通 Test 1/2/3 后 commit + 回信我做 review

## 风险点 / 未知点

- ✅ **HK ticker 格式**: Step 1 确认 — 必须 `00700` 5 位带前导 0,`700` 返回空
- ⚠️ **雪球 HK 单位**: Step 1 确认腾讯 = CNY 不是 HKD;v2 修订 #1 已应对(不强转 + 诊断标 Unit)
- ⚠️ **阿里 H 财年 03 月底**: canonical entry `max(end)+min(start)` + grouping key `month_num + ed_year` 应自然处理,Test 2 验证
- **快手数据完整性**: 上市晚(2021 IPO),早期数据可能稀疏(`data.list` 长度 < 8 是正常)
- **雪球限速**: HK API 与美股共享同一域名 `stock.xueqiu.com`,Codex 沿用美股 1 秒间隔 `Application.Wait`
- ⚠️ **HK 季度报告稀疏**: A4=Q1/Q3 在港股语境下大概率 0 命中(港股公司多数只出 H1+年报),诊断 sheet 应留空但不报错

---

# 零碎事(并行,Codex 在 Phase 4c 主线推进时穿插做)

## Side 1 — 修 CF a(32) ifrs 槽位错配

**File**: `modules/模块_抓美股现金流量表.bas` 第 95-96 行

**Before**:
```vba
a(32) = Array("", "Cash at beginning of period", _
              "NoEdgarConceptCashAtBeginning,CashAndCashEquivalentsAtBeginningOfPeriod")
```

**After**(把 `CashAndCashEquivalentsAtBeginningOfPeriod` 移到 ifrs-full 槽位 index 5):
```vba
a(32) = Array("", "Cash at beginning of period", _
              "NoEdgarConceptCashAtBeginning", _
              "USD", 1000000#, _
              "CashAndCashEquivalentsAtBeginningOfPeriod,CashAndCashEquivalents")
```

**理由**: `CashAndCashEquivalentsAtBeginningOfPeriod` 是 ifrs-full 风格名,不是 us-gaap 标准 concept。放对槽位后,IFRS filer 的 Tier 2 兜底可能会命中。

**验收**: AAPL/AMZN cell-level diff = 0(us-gaap 端永远命中占位 → MISSING,行为不变); BABA / 港股暂不受影响(港股不走 EDGAR)

## Side 2 — Test 1 cell-level 严格回归脚本

**新建 `tools/diff_xlsm.py`**(Python + openpyxl):

```python
"""
对比两个 .xlsm 工作簿的报表数据区每个 cell value, 输出 mismatches。

用法:
    py tools/diff_xlsm.py <new.xlsm> <baseline.xlsm>

例:
    py tools/diff_xlsm.py 上市公司财务数据查询.xlsm 新浪财经行业数据查询V3_稳定版_20260503.xlsm

对比范围: 6 张主表 (A股 BS/IS/CF + 美股 BS/IS/CF) 的 R3+ 所有 cell
跳过: 指标表(公式触发, 数值随时间变化), 诊断表(本来就只有新版有), 样本池/使用说明
"""
```

**期望输出**(理想情况 0 diff):
```
A股_资产负债表: ✅ 0 mismatches
A股_利润表:     ✅ 0 mismatches
A股_现金流量表: ✅ 0 mismatches
美股_资产负债表: ✅ 0 mismatches
美股_利润表:     ✅ 0 mismatches
美股_现金流量表: ✅ 0 mismatches

ALL PASS — Phase 4b-13 stable baseline preserved.
```

**如果出现 diff**: 列出 sheet / cell / new_val / baseline_val,Codex 排查根因(是预期改进还是回归)

## Side 3 — STATUS.md NONUSD 双行说明

**位置**: `STATUS.md` §O.3 已知边界,加一段:

```markdown
- **同一 (公司, 指标) 在诊断 sheet 出现两行属预期行为**:
  当 ifrs-full 命中某 concept 但单位不是 USD,会先 emit 一行 `MISSING_NON_USD`(留下 ifrs taxonomy 有该字段的痕迹);随后 Tier 3 雪球如果命中,emit `OK_XUEQIU` 第二行。两行表示"我们看到 ifrs 有这个字段但单位不对,所以走了雪球",**这是 feature 不是 bug**,留给后续 Phase 4b-14b 决定是否做币种换算。
```

**同步**:`tools/build_template.py` 的 `build_intro` 函数加一行解释(让模板新建的工作簿也有此说明)

---

## Codex 工作顺序建议

```
Day 1:
  Step 1 (HK API 探查) → 提交 → Claude review
  并行做 Side 1 (CF a(32) 修复)

Day 2:
  Step 2-4 (主流程 + 4 wrapper + 工具函数升级)
  并行做 Side 3 (STATUS.md 双行说明)

Day 3:
  Step 5-6 (Python tools + 一键全抓)
  Test 1/2/3 跑通
  并行做 Side 2 (diff_xlsm.py + 跑一次回归确认 Side 1 不破)
  commit + 回信 Claude review
```

总计 ~3 天工作量。每个 Step / Side 完成后 commit + 回信我 review。

---

## Claude (reviewer) 关注重点

| 审查点 | 重点 |
|---|---|
| Step 1 dump | 字段名结构 vs 美股雪球 — 决定字段映射 reuse / rewrite |
| Step 2 RunHKStatement | 是否真的没污染 RunUSStatement;诊断 sheet 切换是否清晰 |
| Step 4 g_diagnosticSheetName | 全局 var 切换的作用域 — 警惕跨表污染 |
| Step 5 install/build | 诊断 sheet 列宽/字色三处一致性(港股 vs 美股 sheet 命名外完全相同) |
| Step 6 一键全抓 | 12 张表顺序;两张诊断 sheet 都正确清空 |
| Side 1 | AAPL/AMZN cell-level 0 diff(配合 Side 2 验证) |
| Side 2 | diff 脚本对 6 张主表的覆盖完整性(指标表正确跳过) |

# Phase 4d: 韩股 (KR) 接入 — 雪球 KR + stockanalysis.com 双源探查

> **版本**: v1(2026-05-03)
> **状态**: ✅ Step 1-6 已由 Codex 实现并通过本地验证,等待 Claude Code review
> **作者**: Claude(planner)+ Codex(executor)
> **背景**: Phase 4c 港股已闭环,A 股 / 美股 / 港股三市场跑通。**韩股是用户记忆里 4 市场目标的最后一块**,完成后即可启动 Phase G 跨市场对标。

## Context

用户提示新增数据源候选:**https://stockanalysis.com/**(金融数据聚合站,英文字段名清晰,支持 US/Korea/Japan 等多市场)。本期 Step 1 同时 probe **雪球 KR API** 与 **stockanalysis.com KR**,实测后选优。

## 执行收口(2026-05-03)

| 项目 | 状态 | 结果 |
|---|---|---|
| Step 1 双源 probe | ✅ 完成 | 雪球 KR 8 路径不可用;stockanalysis.com KRX HTML 表格可用 |
| Step 2 主流程 | ✅ 完成 | 新增 `模块_抓韩股财报.bas`,使用 `htmlfile` DOM + Chrome UA |
| Step 3 thin wrapper | ✅ 完成 | 新增韩股 BS/IS/CF/指标表 4 个 wrapper |
| Step 4 工具函数 | ✅ 完成 | `BuildStandardIndicatorSheet "KR"` 接入 18 个标准指标 |
| Step 5 模板/安装脚本 | ✅ 完成 | 新增韩股 4 表 + `韩股_抓取诊断` + 深紫按钮 `#7030A0` |
| Step 6 一键全抓 | ✅ 完成 | 升级为 16 张正式表,三张诊断表独立清空和写入 |
| 验证 | ✅ 完成 | 5 家韩股 Q4、Q1/Q3、无效代码、完整一键全抓均已跑通 |

最终选择 stockanalysis.com 作为韩股主数据源。原 HTML 表格单位为百万韩元,正式表按十亿韩元(KRW billions)写入,即 `/1000`。诊断状态仅使用 `OK_STOCKANALYSIS` / `MISSING`,不做 fuzzy。

**为什么不直接锁雪球**:
- 雪球 KR 框架现成(类比 4c),但**未实测覆盖度**
- stockanalysis.com 字段名英文清晰、不需要 cookie、可能更稳定 — 但**未实测反爬 + 韩股深度**
- 双源 probe 1 天内就能定,**避免走错路**

## 已锁定的决策(用户已确认)

| 决策 | 选项 |
|---|---|
| 测试公司 | **大盘科技 4 + 床垫 1**: `005930 三星电子` / `000660 SK 海力士` / `035420 NAVER` / `035720 Kakao` / **`013890 Zinus`(KOSDAQ 床垫,与用户原家居 4 公司同业)** |
| 入口架构 | **新写 `RunKRStatement`** — 类比 `RunHKStatement`,独立模块 |
| 货币单位 | **已定:十亿韩元(KRW billions)** — stockanalysis.com HTML 表格为百万韩元,写入正式表时除以 `1,000` |
| Sheet 命名 | `韩股_资产负债表` / `韩股_利润表` / `韩股_现金流量表` / `韩股_指标表`(对称 A股_/美股_/港股_) |
| 指标体系 | 沿用 18 个标准指标(`BuildStandardIndicatorSheet "KR"`)|
| 诊断 sheet | **新建 `韩股_抓取诊断`**(列结构相同 10 列) |
| 一键全抓 | 升级到 16 张表(A 股 4 + 美股 4 + 港股 4 + 韩股 4) |

## 实施方案

### Step 1 — 双源 probe(Codex 必须先做并暂停等 review)

**目标**: 实测雪球 KR 与 stockanalysis.com KR 谁更适合做主数据源。

#### 1A. 雪球 KR 探查(类比 4c Step 1)

测以下 4 个 endpoint(005930 三星 + 013890 Zinus 各试一次,验证 KOSPI + KOSDAQ 都通):
```
https://stock.xueqiu.com/v5/stock/finance/kr/balance.json?symbol=005930&type=all&is_detail=true&count=8
https://stock.xueqiu.com/v5/stock/finance/kr/income.json?symbol=005930&type=all&is_detail=true&count=8
https://stock.xueqiu.com/v5/stock/finance/kr/cash_flow.json?symbol=005930&type=all&is_detail=true&count=8
https://stock.xueqiu.com/v5/stock/finance/kr/indicator.json?symbol=005930&type=all&is_detail=true&count=8
```
**注意**: 雪球 KR endpoint 路径可能不是 `/kr/`,可能是 `/kosdaq/` `/kospi/` 或别的。先试 `/kr/`,失败再改。

dump 到 `samples/xueqiu_KR_005930_{kind}.json` + `samples/xueqiu_KR_013890_{kind}.json`(8 个文件)。

#### 1B. stockanalysis.com 探查

试以下 URL(浏览器先打开看能否直接看到数据,然后 curl 看 HTML 结构):
```
https://stockanalysis.com/quote/krx/005930/financials/                  (BS - 实际 URL 待定)
https://stockanalysis.com/quote/krx/005930/financials/balance-sheet/
https://stockanalysis.com/quote/krx/005930/financials/cash-flow-statement/
https://stockanalysis.com/quote/krx/005930/financials/ratios/
```
- 确认是否有 KRX(韩国交易所)股票页面
- 确认是否需要 login / cookie / Cloudflare 是否拦
- 确认数据是 HTML table 还是 inline JSON 还是 fetch API
- dump HTML 到 `samples/stockanalysis_KR_005930_{kind}.html`

#### 1C. 报告

写 `samples/KR_API_PROBE.md`:
- 雪球 KR:HTTP status / `error_code` / `data.list` 长度 / `data.currency` / 字段名清单 / Zinus(KOSDAQ)是否通
- stockanalysis.com:HTTP status / 是否有 Cloudflare 拦截 / 数据载体 / 字段名清单 / 是否需 cookie
- **比较 + 推荐选哪个**
- 数值规模:三星 Total assets 实际数值 → 决定 scale(韩元百万 还是 十亿 还是 兆)

**暂停等 Claude review**,review 通过后再进 Step 2。

### Step 2-6 — 类比 Phase 4c 实施

**Step 2 主流程**: `模块_抓韩股财报.bas`,核心:
- `RunKRStatement` — 入口设 `g_diagnosticSheetName = "韩股_抓取诊断"`
- `FetchKRFromXueqiu` 或 `FetchKRFromStockAnalysis`(看 Step 1 选哪源)
- `XueqiuFieldMapForKindKR` 或 `StockAnalysisFieldMapForKindKR`
- `MatchKRPeriod` — 韩股大多 12 月底财年,Q1/Q2/Q3/Q4 都披露(比港股完整),逻辑应该比 HK 简单

**Step 3 - thin wrappers**: 4 个 `模块_抓韩股*.bas`

**Step 4 - 工具函数**: `BuildStandardIndicatorSheet "KR"` + `AppendStandardIndicators "KR"` + `g_diagnosticSheetName` 已支持

**Step 5 - 工具脚本**:
- `tools/build_template.py`: 加 4 韩股 sheet + 韩股_抓取诊断
- `tools/install_modules.py`:
  - `ensure_market_sheets` 加韩股 sheet
  - `BUTTONS` 加 4 个韩股按钮 — **第四色:深紫 `#7030A0` / 白字**(与 A蓝 / 美红 / 港绿 区分)
  - `reorder_report_sheets` 加韩股段
  - 「使用说明」refresh

**Step 6 - 一键全抓 16 表**

### 关键差异点(vs Phase 4c)

| 点 | Phase 4c 港股 | Phase 4d 韩股 |
|---|---|---|
| 数据源 | 雪球 HK(已确认) | 雪球 KR + stockanalysis.com 双源 probe 选优 |
| 财年 | 12月/3月混合(阿里 H 等) | 几乎全 12月 |
| Q1/Q3 披露 | 大多没有 | 大多有(韩股季报完整) |
| ticker | 5 位带前导 0(00700) | 6 位带前导 0(005930) |
| 货币 | CNY/HKD/USD 混合 | KRW(单一) |
| 数值规模 | 百万 | 待定(可能十亿/兆,看实际值) |
| 按钮颜色 | 深绿 #548235 | 深紫 #7030A0 |

## 验收方案

### Test 1 — 5 韩股一键全抓
样本池: 005930 三星 + 000660 SK 海力士 + 035420 NAVER + 035720 Kakao + 013890 Zinus(可与现有 A股+美股+港股共存)
配置: `A2=2024, A4=Q4`
跑: 一键全抓
**期望**: 16 张正式表 + 3 张诊断 sheet 全填好,5 家韩股核心字段(BS Total assets / IS Revenue+Net income / CF CFO+End Cash)为 `OK_STOCKANALYSIS`,失败不应中断流程

### Test 2 — KOSPI vs KOSDAQ 覆盖
- 005930(KOSPI 三星)+ 013890(KOSDAQ Zinus)同时跑
- **期望**: 两板都能抓到,Zinus 不会因 KOSDAQ 板块特殊性失败

### Test 3 — Q1/Q3 季报覆盖
- 005930 三星 + A4=Q1 / Q3 各跑一次
- **期望**: 跟港股不同,韩股 Q1/Q3 应该有数据(韩股季报完整)

### Test 4 — 边界
- 故意填 `999999` → 失败但不中断 + 诊断全 MISSING

## 文件改动清单

| 文件 | 改动 | 责任 |
|---|---|---|
| `samples/xueqiu_KR_005930_*.json`(4) + `xueqiu_KR_013890_*.json`(4) + `stockanalysis_KR_*.html` | dump 调试样本 | Codex Step 1 |
| `samples/KR_API_PROBE.md` | Step 1 探查 + 数据源选型报告 | Codex Step 1 |
| `modules/模块_抓韩股财报.bas` | 新建主流程 | Codex Step 2 |
| `modules/模块_抓韩股资产负债表.bas` | 新建 thin wrapper | Codex Step 3 |
| `modules/模块_抓韩股利润表.bas` | 新建 | Codex Step 3 |
| `modules/模块_抓韩股现金流量表.bas` | 新建 | Codex Step 3 |
| `modules/模块_抓韩股指标表.bas` | 新建 | Codex Step 3 |
| `modules/模块_工具函数.bas` | `BuildStandardIndicatorSheet/AppendStandardIndicators` 加 KR 分支 | Codex Step 4 |
| `modules/模块_总入口.bas` | 一键全抓 升到 16 张表 + 3 张诊断 sheet 清空 | Codex Step 6 |
| `tools/build_template.py` | 加 5 张韩股 sheet 模板 | Codex Step 5 |
| `tools/install_modules.py` | ensure_market_sheets 升 + 4 个紫色按钮 + reorder + 使用说明 | Codex Step 5 |

## Codex 工作顺序

```
Step 1 (双源 probe) → 暂停 review → 选源
  并行做 Side 1+2+3 (零碎事)
Step 2 (主流程)
Step 3-4 (wrapper + 工具)
Step 5-6 (Python tools + 一键全抓)
Test 1/2/3/4
```

## 风险点 / 未知点

- **雪球 KR endpoint 是否真存在**: `/v5/stock/finance/kr/...` 未实测,可能是 `/kosdaq/` 或别的;Step 1 实测
- **stockanalysis.com 反爬**: Cloudflare 可能拦 WinHttp;Step 1 必测
- **KRW 数值规模**: 三星 Total assets ~400 兆 KRW,如果 scale=1e6 cell 会是 4e8(8 位数),可读性差;scale=1e9(十亿)显示更友好;scale=1e12(兆)对于小公司又太粗。**待 Step 1 后定 scale**,初步推荐 1e9(十亿 KRW)
- **Zinus(013890)股价低 + 流动性低**: 雪球可能数据稀疏,作为 KOSDAQ 测试样本是有意识的压力测试

## Claude (reviewer) 关注重点

| 审查点 | 重点 |
|---|---|
| Step 1 双源选型 | 数据完整性 + cookie/反爬 + 字段命名清晰度 — 选优时给清晰理由 |
| Step 2 主流程 | 不污染美股/港股代码;诊断 sheet `g_diagnosticSheetName` 切换正确 |
| Step 5 install/build | 4 个紫色按钮颜色;sheet 顺序;使用说明 4 市场段都齐 |
| Step 6 一键全抓 16 表 | 顺序对;3 张诊断 sheet 都正确清空;`g_diagnosticAppendOnly` 在 Clear 之后才设 True |
| Test 3 Q1/Q3 | 韩股季报应有数据(对比港股大多 0 命中)|

---

# 零碎事(并行,Codex 主线推进时穿插做)

## Side 1 — 重生成 4b-14a 严格回归 baseline

**目标**: 闭环 Phase 4c Side 2 留下的"为什么 diff 不为 0"问题。

**步骤**:
1. 临时把样本池改成 4 公司测试集:`AAPL / AMZN / POM / HTT`(行 8-11),A2=2025,A4=全部,清掉港股/A股 公司
2. 跑 `tools/install_modules.py` + 一键全抓
3. 把生成的 `上市公司财务数据查询.xlsm` 复制到 `archive/上市公司财务数据查询_4b14a_baseline_20260503.xlsm`
4. 还原样本池(加回 A股 4 + 港股 4 + 美股原 4)
5. 更新 `tools/diff_xlsm.py` 的 `DEFAULT_BASELINE` 指向新 baseline
6. 验证:再跑一次同样 4 公司 → diff 应为 0

**预期**: Side 2 验收闭环,后续任何小改动都能通过 0 diff 回归确认无破坏

## Side 2 — `PHASE_4C_HK_PLAN.md` 状态改 ✅ 全部完成

**位置**: 文件头部 line 4 状态行

**Before**:
```
> **状态**: 🚧 Step 5 + Step 6 已实现,等待 Claude Code review
```

**After**:
```
> **状态**: ✅ Phase 4c 全部完成并通过 Claude Code 闭环 review(2026-05-03)
```

每个 Step 标题后追加 ✅ 已闭环 标记(Step 1-6 + Side 1/2/3)。

## Side 3 — `STATUS.md` §M 加备份说明

**位置**: §M.2 末尾(Phase 4b-13 总回归验证段)

**追加一段**:

```markdown
### M.3 baseline 备份位置

为便于后续严格回归验证, Phase 4b-13 的稳定版备份保存在:

- `archive/新浪财经行业数据查询V3_稳定版_20260503.xlsm` (项目改名前的快照, Phase 4b-13 完成态)
- `archive/新浪财经行业数据查询V3_更名前备份_20260503.xlsm` (项目改名前同期备份, 与稳定版数据一致)

注意:这两份文件名仍是旧的"新浪财经行业数据查询V3", 但内容已经是 Phase 4b-13 完成态。Phase 4d 起 baseline 改用 `archive/上市公司财务数据查询_4b14a_baseline_20260503.xlsm`(由本期 Side 1 生成, 4b-14a 完成态 + 4 美股测试公司)。
```

---

## Codex 工作顺序建议

```
Day 1:
  Side 1 (4b-14a baseline 重生成) — 0.5 天
  Side 2 + 3 (文档收口) — 0.5 天
  Step 1 (双源 probe + 报告) — 0.5 天 → 暂停等 review

Day 2:
  (Step 1 review 通过后)
  Step 2-3 (主流程 + 4 wrapper)
  Step 4-5 (工具函数 + Python tools)

Day 3:
  Step 6 (一键全抓 16 表)
  Test 1/2/3/4
  最终 commit + 报告

总计 ~2.5-3 天工作量
```

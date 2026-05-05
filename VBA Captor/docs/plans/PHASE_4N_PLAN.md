# Phase 4n: 优化 Sprint 4 收尾 — AppStateGuard 扩展 + 架构文档 + release notes

> **版本**: v2(2026-05-05,Phase 4n 优化 Sprint 4 闭环)
> **状态**: ✅ Phase 4n 全期闭环
> **作者**: Claude(planner) + Codex(generator)
> **背景**: Phase 4k-4m 三个 sprint 已交付优化主线(数据准确性 / live FX / HTTP 可观测 / 重试 / 发布清理 / 离线测试 / 数据质量 QA / cache 分源 TTL)。Phase 4n 是**收尾性质 sprint** — 把 Phase 4k 的 AppStateGuard 扩到剩余 5 个入口(一键X股 4 个 + 跨市场指标表)+ 写一份 ARCHITECTURE.md + README 加 Phase 4f-4n release notes 段。**不做大重构,不动业务逻辑**。

## 项目语境(给 Generator 的 anchor 段)

本期 3 项工作**全部本地,零网络改动,零业务逻辑改动**。Step 1 复用 Phase 4k 已 proven 的 BeginAppState/EndAppState pattern apply 到 5 个剩余入口,Step 2 写新文档,Step 3 跑回归 + 收口。

## 3 项任务清单

| # | Task | 价值 | 工作量 |
|---|---|---|---|
| 1 | AppStateGuard 扩展到剩余 5 个入口(一键A/美/港/韩股 + BuildCrossMarketIndicatorSheet)| 防御性一致(Phase 4k 只做了 一键全抓,其他入口出错仍可能脏 Excel 状态)| 1.5h |
| 2 | 新增 `ARCHITECTURE.md` + README 加 release notes 段(Phase 4f → 4n)| 维护性(给 Codex 后续 / 将来回顾提供 high-level 视图)| 1.5h |
| 3 | 端到端回归 + STATUS §EE 收口 + plan v2 升级 | — | 0.5h |

总耗时 ~3.5h Codex,~30min Reviewer。

## Step 总览

| Step | 内容 | 估时 | 阻塞依赖 |
|---|---|---|---|
| 1 | AppStateGuard 扩展:`一键A股 / 一键美股 / 一键港股 / 一键韩股` + `BuildCrossMarketIndicatorSheet` 5 个 Sub 包 BeginAppState/EndAppState | 1.5h | 无 |
| 2 | `ARCHITECTURE.md` 新建(模块依赖图 + sheet inventory + 数据流 + 关键 invariants)+ `README.md` 末尾加 Phase 4f-4n release notes 表 | 1.5h | 无 |
| 3 | 8 张回归驱动全 PASS + STATUS §EE 收口 | 0.5h | 1, 2 |

**Codex 工作流建议**:**1 round 端到端**(跟 Phase 4j-4m 一致),commit `Phase 4n optimization sprint 4 closure (AppStateGuard rollout + architecture docs)`。

---

## Step 1 — AppStateGuard 扩展到剩余 5 入口

### 背景

Phase 4k Step 2 只给 `一键全抓` apply 了 BeginAppState/EndAppState。剩余 5 个数据入口仍然依赖各 Sub 自己手写恢复(部分恢复了 ScreenUpdating + StatusBar,但 Calculation / EnableEvents / DisplayAlerts 漏)。

### 实施目标

把 `BeginAppState/EndAppState` pattern apply 到:

1. `一键A股`(`模块_总入口.bas`)
2. `一键美股`(`模块_总入口.bas`)
3. `一键港股`(`模块_总入口.bas`)
4. `一键韩股`(`模块_总入口.bas`)
5. `BuildCrossMarketIndicatorSheet`(`模块_工具函数.bas`)

### Pattern(Phase 4k 已 proven)

```vba
Public Sub 一键A股()
    Dim st As TAppState
    On Error GoTo EH
    st = BeginAppState("正在抓取 A 股...")
    
    ' ... 现有抓数 + UnhideMarketTabs 等 logic 不动 ...
    
CleanExit:
    EndAppState st
    Exit Sub
EH:
    Application.StatusBar = "一键 A 股出错: " & Err.Description
    Resume CleanExit
End Sub
```

### 子任务 1A — 4 个一键X股 Sub 包夹

`模块_总入口.bas` 现有结构(grep 验证):
- `Public Sub 一键A股()` — 抓 A 股 4 张正式表
- `Public Sub 一键美股()` — 同上
- `Public Sub 一键港股()` — 同上
- `Public Sub 一键韩股()` — 同上

每个 Sub:
1. 开头加 `Dim st As TAppState` + `On Error GoTo EH` + `st = BeginAppState("正在抓取 X 股...")`
2. 末尾加 `CleanExit:` label + `EndAppState st` + `Exit Sub` + `EH:` handler
3. 现有 `UnhideMarketTabs` 调用 移到 `CleanExit:` 之前(确保即使错误也展开 sheet — 或者放 `CleanExit:` 之后只在成功时展开,**Codex 选保守的:错误时不展开** 让用户先看错误)

### 子任务 1B — BuildCrossMarketIndicatorSheet 包夹

`模块_工具函数.bas` `BuildCrossMarketIndicatorSheet` 当前已经手动管理 `Application.DisplayAlerts = False/True`(Phase 4j.1 抑制合并弹窗)和 `Application.ScreenUpdating = False/True`。**改成统一用 BeginAppState/EndAppState** 替代:

```vba
Public Sub BuildCrossMarketIndicatorSheet()
    Dim st As TAppState
    On Error GoTo EH
    st = BeginAppState("正在合并跨市场指标表...")
    
    ' ... 现有 unmerge / clear / 写公式 / 写 QA(Phase 4m)等 logic 不动 ...
    
CleanExit:
    EndAppState st
    Exit Sub
EH:
    Application.StatusBar = "合并跨市场指标表出错: " & Err.Description
    Resume CleanExit
End Sub
```

**注意**:`BeginAppState` 内部已经 set `DisplayAlerts = False`,所以现有 `Application.DisplayAlerts = False` 显式调用可以删除(但保留也无害)。

### 验证

- VBE 编译通过
- 故意把 `BuildCrossMarketIndicatorSheet` 内调用的 `BuildCrossMarketIndicatorSheetCore`(假设)改成不存在的名字,跑 `一键全抓` → Excel `Application.Calculation` 在错误后恢复到 `xlCalculationAutomatic`
- 跑 `一键A股` 正常完成 → StatusBar 恢复成默认值

### Generator 不要做

- ❌ 不要改 `切换X股tabs`(4 个)/ `切换所有分市场tabs` / `切换跨市场tabs` 等 toggle Sub(逻辑轻量,不需要 AppStateGuard)
- ❌ 不要改各 fetch helper(`FetchUSFromXueqiu` / `FetchHKFromXueqiu` 等)— AppStateGuard 只 apply 到 entry-level Sub
- ❌ 不要在 `EH:` handler 里 MsgBox(用 StatusBar 即可,避免抓数失败时弹窗打断用户操作)
- ❌ 不要把 `On Error Resume Next` 范围扩大(只在 EndAppState 内部用,Phase 4k 已设计)

---

## Step 2 — 架构文档 + README release notes

### 子任务 2A — 新建 `ARCHITECTURE.md`

放在 `VBA Captor/ARCHITECTURE.md`,~250-400 行,目录:

```markdown
# 上市公司财务数据查询 — 架构

> Last updated: 2026-05-05 (Phase 4n)

## 1. 项目目标

一句话:Excel + VBA 桌面工具,让财务/审计专业人士在 1 分钟内完成 A 股 / 美股 / 港股 / 韩股 同期可比财报抓取 + 跨市场对标。

## 2. 模块依赖图(VBA)

```
入口层 (模块_总入口)
  ↓
  ├── 抓数层
  │   ├── 模块_抓资产负债表 / 模块_抓利润表 / 模块_抓现金流量表 / 模块_抓指标表 (A 股, 新浪)
  │   ├── 模块_抓美股财报 (EDGAR + 雪球 fallback + stockanalysis 中概 fallback)
  │   ├── 模块_抓港股财报 (雪球 HK API)
  │   └── 模块_抓韩股财报 (stockanalysis.com KRX)
  ↓
  ├── 写表层 (模块_工具函数.WriteWideTable + BuildStandardIndicatorSheet)
  │     使用 GetFxFromSheet UDF 实时换算原币 ↔ RMB
  ↓
  ├── 跨市场层 (模块_工具函数.BuildCrossMarketIndicatorSheet)
  │     合并 4 张分市场指标表到 1 张跨市场视图
  ↓
  └── 跨切面服务
       ├── 模块_AppStateGuard (Phase 4k) — Excel 状态守护
       ├── 模块_工具函数.RunCachedHttpGet (Phase 4l) — HTTP/cache 统一入口
       ├── 模块_工具函数.HttpGetWithRetry (Phase 4l) — retry/backoff/SEC 限流
       ├── 模块_工具函数.RunDataQualityChecks (Phase 4m) — BS 平衡 / FX missing / 关键字段 3 条 QA
       ├── 模块_抓汇率 (Phase 4f) — USDCNY / HKDCNY / KRWCNY 汇率抓取
       └── 模块_测试 (Phase 4m) — 8 个 Test_Offline_* 离线测试
```

## 3. Sheet Inventory

| Sheet | 用途 | 数据来源 |
|---|---|---|
| 使用说明 | 用户文档(7 section + TOC)| 装表时硬编码 |
| 样本池 | 输入 + 操作 | 用户录入 |
| A股_资产负债表 / _利润表 / _现金流量表 / _指标表 | A 股报表 | 模块_抓X 写入 |
| 美股_* | 美股报表 | 模块_抓美股财报 写入 |
| 港股_* | 港股报表 | 模块_抓港股财报 写入 |
| 韩股_* | 韩股报表 | 模块_抓韩股财报 写入 |
| 美股_抓取诊断 / 港股_抓取诊断 / 韩股_抓取诊断 | 17 列遥测(默认隐藏)| WriteDiagnosticForKind / RunDataQualityChecks |
| 跨市场_指标表 | 4 市场 18 标准指标合表(公式 ref)| BuildCrossMarketIndicatorSheet |
| 汇率 | USDCNY / HKDCNY / KRWCNY 期末 + 期均 cache | 模块_抓汇率 写,GetFxFromSheet UDF 读 |

## 4. 数据流(从用户录入到看到数值)

```
1. 用户在样本池录入代码 + 简称, 选 B6 = 原币 / 统一RMB
2. 点 一键全抓 4 市场 → 模块_总入口.一键全抓
   ↓ BeginAppState
3. 顺序调 4 市场抓数 Sub → 各调 模块_抓X股财报
   ↓ Cache wrapper (RunCachedHttpGet) 优先读 .cache/ 24h 内的 JSON
   ↓ 否则调 HttpGetWithRetry → 真打 HTTP (SEC ≤10/s 限流)
   ↓ 拿到 body → 解析字段 → 写 4 张分市场表
4. 4 市场全部完成后, 自动调 BuildCrossMarketIndicatorSheet
   ↓ 收集 4 张分市场指标表 R3-R20 (18 标准指标)
   ↓ 写公式 cell-ref 到对应分市场表
   ↓ 公式包含 GetFxFromSheet UDF, B6 切换时实时刷新
5. 末尾自动调 RunDataQualityChecks → 写 GLOBAL_QA 3 行到美股诊断 sheet
6. EndAppState — 恢复 Excel 状态
```

## 5. 关键 Invariants(代码必须维护的不变量)

- Phase 4f frozen:`模块_抓汇率.bas` 不动,`GetFxRate` 旧签名兼容
- Phase 4k frozen:`GetFxRate / GetFxRateStatus / GetFxFromSheet` 签名,FX 缺失永不 fallback 到 1
- Phase 4l frozen:`THttpResult` 12 字段,核心 HTTP 函数(`EdgarHttpGet / XueqiuHttpGet / StockAnalysisHttpGet`)签名
- 诊断 sheet 17 列(L:Q 文本格式 `@`)
- 跨市场指标表 18 行标准指标 + 公式 cell-ref
- 样本池 R14+ 用户数据安全(任何 layout 重构都不能清掉)
- AppStateGuard:所有数据入口必须 BeginAppState/EndAppState 包夹

## 6. 关键文件路径

```text
VBA Captor/
├── 上市公司财务数据查询.xlsm          # 主工作簿
├── README.md                          # 用户文档
├── STATUS.md                          # 开发追溯
├── ARCHITECTURE.md                    # 架构(本文档)
├── modules/                           # VBA 源码 (装表脚本注入到 xlsm)
│   ├── 模块_AppStateGuard.bas         # Phase 4k
│   ├── 模块_工具函数.bas              # 主工具集(含 HTTP / cache / FX / QA / write)
│   ├── 模块_总入口.bas                # 4 个一键X股 + 一键全抓
│   ├── 模块_抓资产负债表 / 利润表 / ...   # A 股
│   ├── 模块_抓美股财报.bas
│   ├── 模块_抓港股财报.bas
│   ├── 模块_抓韩股财报.bas
│   ├── 模块_抓汇率.bas
│   └── 模块_测试.bas                   # 8 个 Test_Offline_* (Phase 4m)
├── tests/fixtures/                    # 离线测试 fixture (.gitignore)
├── tools/                             # Python 装表 + 回归驱动
│   ├── install_modules.py             # 装表 (核心)
│   ├── build_template.py              # 生成空模板
│   ├── test_fx_live.py                # Phase 4f frozen 回归
│   ├── diff_phase4f_step3_lite.py     # RMB 短路验证
│   ├── inspect_phase4{g,h,k,l,m}_state.py  # 各 phase state 验证
│   └── run_offline_tests.py           # Phase 4m 离线测试 runner
└── samples/                           # 调研期收集的真实 sample (.gitignore 部分)
```

## 7. 数据源声明

| 市场 | 数据源 | URL 模式 | Cookie 需求 |
|---|---|---|---|
| A 股 | 新浪财经 | `https://money.finance.sina.com.cn/...` | 不需要 |
| 美股 | SEC EDGAR + 雪球 fallback + stockanalysis 中概 | `https://data.sec.gov/api/xbrl/companyfacts/CIK*.json` | 雪球 fallback 需要 cookie |
| 港股 | 雪球 HK API | `https://stock.xueqiu.com/v5/stock/finance/cn/...` | 需要 cookie |
| 韩股 | stockanalysis.com KRX | `https://stockanalysis.com/quote/krx/...` | 不需要 |
| 汇率 | 雪球 K 线 | `https://stock.xueqiu.com/v5/stock/chart/kline.json?symbol=USDCNYC...` | 不需要 |

## 8. Cache 分源 TTL(Phase 4m)

| Source | TTL(小时)| 理由 |
|---|---|---|
| SEC_TICKER_MAP | 168 | 几乎不变 |
| EDGAR | 24 | 财报季度更新 |
| XUEQIU | 12 | 实时数据稍频繁 |
| STOCKANALYSIS_KR / _US | 24 | HTML 变更慢 |
| FX_KLINE | 24 | 每日 close 后稳定 |
| (default) | 24 | 兜底 |
```

### 子任务 2B — README 加 release notes

`README.md` 末尾追加 section:

```markdown
## 版本历史

从 Phase 4f 起,本工具进入"多市场 + 跨市场对标 + 优化"阶段。

| Phase | 日期 | 主线交付 |
|---|---|---|
| 4f | 2026-05 | RMB 换算 hook + 汇率 sheet + B6 toggle scaffold |
| 4g | 2026-05 | 跨市场指标合表 + hide-tab 按钮 + POOL_DATA_START_ROW 迁移 |
| 4h | 2026-05 | B6 实时 toggle + 磁盘 JSON 缓存 + stockanalysis 中概美股 fallback |
| 4i / 4i.1 / 4i.2 | 2026-05 | UX 抛光(样本池布局 / 使用说明商务化 / 汇率说明区 / 单表按钮删除 / fallback 自动化)|
| 4j / 4j.1 - 4j.4 | 2026-05 | 简化跨市场对比 + 抑制合并弹窗 + 样本池视觉 1:1 还原 |
| 4k | 2026-05 | 优化 Sprint 1:数据准确性(FX missing / KR Score)+ live FX UDF + AppStateGuard |
| 4l | 2026-05 | 优化 Sprint 2:HTTP/cache 诊断遥测 + 重试退避 + 发布清理宏 |
| 4m | 2026-05 | 优化 Sprint 3:8 个离线测试 + 数据质量 QA + cache 分源 TTL |
| 4n | 2026-05 | 优化 Sprint 4 收尾:AppStateGuard 全入口覆盖 + ARCHITECTURE.md |

详细 commit 历史见 `STATUS.md`。
```

### 验证

- 打开 `ARCHITECTURE.md` 在编辑器渲染,8 个 section + 模块依赖图 + sheet inventory + 数据流图 + 关键 invariants 表都正常
- README 末尾 release notes 表 9 行(4f → 4n)

### Generator 不要做

- ❌ 不要把 ARCHITECTURE.md 写超 600 行(太长没人看)
- ❌ 不要在 ARCHITECTURE.md 里 dump 完整代码片段(用 prose + 简短 pseudo code)
- ❌ 不要新建 CHANGELOG.md / TEST_CASES.md / RELEASE_CHECKLIST.md 等(STATUS.md 已经覆盖追溯,本期不新增)

---

## Step 3 — 端到端回归 + STATUS §EE 收口

### 3A. 跑 frozen 8 张

```bash
py tools/test_fx_live.py --skip-install
py -u tools/diff_phase4f_step3_lite.py
py -u tools/inspect_phase4g_state.py
py -u tools/inspect_phase4h_state.py
py -u tools/inspect_phase4k_state.py
py -u tools/inspect_phase4l_state.py
py -u tools/inspect_phase4m_state.py
py tools/run_offline_tests.py
```

任一退化立即停下。

### 3B. 不需要新增 inspect_phase4n_state.py

本期改动是 AppStateGuard 扩展 + 文档,无新 state。`inspect_phase4k_state.py` 已经验证 AppStateGuard 模块存在 + Begin/End 函数定义。Step 1 扩展完后,**手工**在 VBE 跑 1 次 `?Application.Calculation`(应为 -4105 = xlCalculationAutomatic),跑 `一键A股` 故意制造错误,再跑 `?Application.Calculation` 看是否还是 -4105 — 这步**让 Codex 在 commit message 报告手工验证结果**,不写自动 inspect。

### 3C. STATUS §EE 收口

模仿 §DD 格式追加:

```markdown
## EE. Phase 4n 收口: 优化 Sprint 4 — AppStateGuard 全入口覆盖 + 架构文档

执行依据: PHASE_4N_PLAN.md v1。状态: ✅ Codex 已实现并通过 8 张 frozen 回归 + 手工 AppStateGuard 错误恢复验证。

### EE.1 已完成
- [Step 1] AppStateGuard 扩展到 5 个入口: 一键A股 / 一键美股 / 一键港股 / 一键韩股 + BuildCrossMarketIndicatorSheet
- [Step 2] 新增 ARCHITECTURE.md (~300 行,8 section);README 末尾加 4f-4n release notes 表 9 行

### EE.2 验证结果
- 8 张 frozen 全 PASS
- 手工 AppStateGuard 验证: 故意 corrupt `一键A股` 调用链, 错误后 Application.Calculation = xlCalculationAutomatic ✓

### EE.3 已知边界
- 切换tabs sub (5 个) 不加 AppStateGuard (逻辑轻量, 不需要)
- 各 fetch helper 内部不加 AppStateGuard (只 apply 到 entry-level Sub)
- ARCHITECTURE.md 不含 CHANGELOG/TEST_CASES/RELEASE_CHECKLIST (STATUS.md 已覆盖)
- 优化 backlog 剩余 P0-04 / P1-04 / P1-07 / P2-01 / P2-04 全部 defer 或不做(详见 §EE.4)

### EE.4 优化 backlog 总账

GPT 5.5 Pro 静态审阅 14 项 backlog,Phase 4k-4n 4 个 sprint 的处置:

| 项 | 状态 | 处置理由 |
|---|---|---|
| P0-01 凭证清理 | ✅ Phase 4l(简化 CleanReleaseWorkbook 宏)|
| P0-02 KR Score 日期化 | ✅ Phase 4k |
| P0-03 AppStateGuard | ✅ Phase 4k(1 入口)+ Phase 4n(扩到 5 入口)|
| P0-04 PowerShell 加固 | ⏸️ defer(P2,企业部署再说;cookie 不在 cmdline)|
| P0-05 FX missing 不 fallback 1 | ✅ Phase 4k(critical fix)|
| P1-01 HTTP/cache 诊断遥测 | ✅ Phase 4l |
| P1-02 cache 分源 TTL | ✅ Phase 4m |
| P1-03 retry/backoff/限流 | ✅ Phase 4l |
| P1-04 汇率 fiscal period | ⏸️ defer(Phase 4g fuzzy match 已 80% 兜住,ROI 不够)|
| P1-05 公式 live ref 汇率表 | ✅ Phase 4k |
| P1-06 数据质量 QA | ✅ Phase 4m(精简 3 条)|
| P1-07 OERN 全收敛 | ⏸️ defer(顺手做即可,不专 sprint)|
| P2-01 拆分模块_工具函数 | ⏸️ defer(几个月后再说)|
| P2-02 MarketAdapter | ❌ 不做(无新市场需求)|
| P2-03 离线 fixture | ✅ Phase 4m |
| P2-04 发布与版本管理 | ⏸️ 部分 ✅ Phase 4n(ARCHITECTURE.md + README release notes;CHANGELOG/RELEASE_CHECKLIST defer)|

**14 项中 9 项 ✅,5 项 defer/不做。优化阶段全部完成。**
```

PHASE_4N_PLAN.md v1 → v2,标记 ✅ Phase 4n 全期闭环。

---

## ⚠️ 全 Phase 严禁动

| 文件/区域 | 原因 |
|---|---|
| 任何 fetch / cache / WriteWideTable / FX / QA 业务逻辑 | 优化阶段已闭环,本期纯收尾 |
| Phase 4k-4m 的 helper 签名 | frozen |
| 16 张分市场 sheet + 跨市场指标表 + 诊断 sheet 内容 / 列结构 | frozen |
| 8 张 frozen 回归驱动 | frozen |
| 样本池 R14+ 用户数据 | frozen |
| 切换tabs 等 toggle Sub(5 个)| 逻辑轻量,不加 AppStateGuard |

## ⚠️ 联系 Planner 触发条件

- Step 1 改 5 个入口后,任一 frozen 回归 PASS → FAIL
- Step 1 引入 EH handler 导致 fetch 失败时 Excel 状态恢复**不完整**(eg Calculation 仍 manual)
- Step 2 ARCHITECTURE.md 某 section 引用了不存在的 helper / sheet / 模块名
- VBE Compile 失败

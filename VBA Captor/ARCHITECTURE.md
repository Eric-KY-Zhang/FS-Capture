# 上市公司财务数据查询 — 架构

> Last updated: 2026-05-05 (Phase 4n)

本文档描述 `VBA Captor` 目录下的 Excel/VBA 桌面工具架构。它面向后续维护者和 Codex/Reviewer 交接使用,不替代 `STATUS.md` 的逐 phase 追溯记录。

## 1. 项目目标

本工具是给财务、审计和投研专业人士使用的本地 Excel 桌面报表工具。

核心目标:

- 在一个工作簿里维护 A 股、美股、港股、韩股样本池。
- 一键抓取并写出四市场上市公司的资产负债表、利润表、现金流量表和指标表。
- 统一展示 18 个标准指标,支持跨市场同业对标。
- 通过 `原币 / 统一RMB` 切换支持跨市场金额口径比较。
- 尽量把网络、缓存、汇率、诊断和 QA 都留在本地,不引入服务端依赖。

非目标:

- 不是 Web 服务。
- 不是通用爬虫框架。
- 不是证券行情终端。
- 不写入外部数据库。
- 不做自动交易或投资建议。

设计取舍:

- 数据落在 Excel sheet 中,便于用户二次编辑和审计留痕。
- VBA 作为运行时,Python 只负责装表和回归驱动。
- 入口宏尽量少,用户日常只需要样本池上的一键按钮。
- 诊断 sheet 默认隐藏,保留给排错和 Reviewer 验收使用。

## 2. 模块依赖图(VBA)

```text
入口层
└─ 模块_总入口
   ├─ 一键A股 / 一键美股 / 一键港股 / 一键韩股
   ├─ 一键全抓
   ├─ 一键跨市场指标表
   ├─ 一键清空所有数据
   └─ 显示/隐藏按钮入口

抓数层
├─ A 股
│  ├─ 模块_抓资产负债表
│  ├─ 模块_抓利润表
│  ├─ 模块_抓现金流量表
│  └─ 模块_抓指标表
├─ 美股
│  ├─ 模块_抓美股财报
│  ├─ 模块_抓美股资产负债表
│  ├─ 模块_抓美股利润表
│  ├─ 模块_抓美股现金流量表
│  └─ 模块_抓美股指标表
├─ 港股
│  ├─ 模块_抓港股财报
│  ├─ 模块_抓港股资产负债表
│  ├─ 模块_抓港股利润表
│  ├─ 模块_抓港股现金流量表
│  └─ 模块_抓港股指标表
└─ 韩股
   ├─ 模块_抓韩股财报
   ├─ 模块_抓韩股资产负债表
   ├─ 模块_抓韩股利润表
   ├─ 模块_抓韩股现金流量表
   └─ 模块_抓韩股指标表

写表和通用服务
└─ 模块_工具函数
   ├─ WriteWideTable
   ├─ BuildStandardIndicatorSheet
   ├─ BuildCrossMarketIndicatorSheet
   ├─ GetFxRate / GetFxRateStatus / GetFxFromSheet
   ├─ RunCachedHttpGet / HttpGetWithRetry
   ├─ WriteDiagnosticForKind / AddDiagnosticRow
   ├─ RunDataQualityChecks
   └─ ClearLocalCache / CleanReleaseWorkbook

跨切面模块
├─ 模块_AppStateGuard
│  ├─ BeginAppState
│  └─ EndAppState
├─ 模块_抓汇率
│  └─ EnsureFxRateCached
├─ 模块_测试
│  ├─ Phase 4f-4l smoke macros
│  └─ Test_Offline_* macros
└─ JsonConverter
   └─ JSON parser
```

依赖方向:

- `模块_总入口` 调用抓数层和工具层。
- 抓数层调用 `模块_工具函数` 做 HTTP/cache、字段诊断和写表。
- `模块_工具函数` 读取 `模块_抓汇率` 写出的汇率 sheet,但不改汇率抓取主路径。
- `模块_测试` 只用于本地验证,不被用户入口依赖。

## 3. Sheet Inventory

| Sheet | 用途 | 数据来源 | 可见性 |
|---|---|---|---|
| 使用说明 | 用户说明和操作步骤 | `tools/install_modules.py` 写入 | 可见 |
| 样本池 | 参数、样本公司、按钮区 | 用户录入 + 装表脚本布局 | 可见 |
| 汇率 | USDCNY/HKDCNY/KRWCNY 期末和期均汇率 | `模块_抓汇率` | 可见 |
| 跨市场_指标表 | 四市场 18 项标准指标合并视图 | `BuildCrossMarketIndicatorSheet` | 可见/可隐藏 |
| A股_资产负债表 | A 股资产负债表宽表 | 新浪财经路径 | 可见/可隐藏 |
| A股_利润表 | A 股利润表宽表 | 新浪财经路径 | 可见/可隐藏 |
| A股_现金流量表 | A 股现金流量表宽表 | 新浪财经路径 | 可见/可隐藏 |
| A股_指标表 | A 股 18 项标准指标 | A 股三张基础表公式 | 可见/可隐藏 |
| 美股_资产负债表 | 美股资产负债表宽表 | EDGAR / fallback | 可见/可隐藏 |
| 美股_利润表 | 美股利润表宽表 | EDGAR / fallback | 可见/可隐藏 |
| 美股_现金流量表 | 美股现金流量表宽表 | EDGAR / fallback | 可见/可隐藏 |
| 美股_指标表 | 美股 18 项标准指标 | 美股三张基础表公式 | 可见/可隐藏 |
| 港股_资产负债表 | 港股资产负债表宽表 | 雪球 HK API | 可见/可隐藏 |
| 港股_利润表 | 港股利润表宽表 | 雪球 HK API | 可见/可隐藏 |
| 港股_现金流量表 | 港股现金流量表宽表 | 雪球 HK API | 可见/可隐藏 |
| 港股_指标表 | 港股 18 项标准指标 | 港股三张基础表公式 | 可见/可隐藏 |
| 韩股_资产负债表 | 韩股资产负债表宽表 | stockanalysis KRX | 可见/可隐藏 |
| 韩股_利润表 | 韩股利润表宽表 | stockanalysis KRX | 可见/可隐藏 |
| 韩股_现金流量表 | 韩股现金流量表宽表 | stockanalysis KRX | 可见/可隐藏 |
| 韩股_指标表 | 韩股 18 项标准指标 | 韩股三张基础表公式 | 可见/可隐藏 |
| 美股_抓取诊断 | 美股字段、HTTP、FX、QA 诊断 | `WriteDiagnosticForKind` / QA | 默认隐藏 |
| 港股_抓取诊断 | 港股字段、HTTP、FX 诊断 | `WriteDiagnosticForKind` | 默认隐藏 |
| 韩股_抓取诊断 | 韩股字段、HTTP、FX 诊断 | `WriteDiagnosticForKind` | 默认隐藏 |

诊断 sheet 固定 17 列:

```text
A 公司
B 报表
C 输出指标
D 状态
E 数据源
F Taxonomy
G 命中字段
H Unit
I Score
J 匹配方式+备注
K FX_Rate
L CacheStatus
M CacheAgeHours
N HTTPStatus
O ElapsedMs
P RetryCount
Q ErrorStage
```

## 4. 数据流

### 4.1 用户到分市场报表

```text
1. 用户在「样本池」选择年度、季度、显示币种。
2. 用户从 Row 14 起按市场录入代码和简称。
3. 用户点击一键按钮。
4. 入口宏调用 BeginAppState,保存并设置 Excel 运行状态。
5. 对应市场抓数模块读取样本池。
6. 抓数模块优先走 RunCachedHttpGet。
7. cache 命中时直接使用本地响应。
8. cache 未命中时走 HttpGetWithRetry。
9. 抓数模块解析字段,构造 dictData / concept map / 诊断 rows。
10. WriteWideTable 写出宽表。
11. BuildStandardIndicatorSheet 生成 18 项标准指标表。
12. EndAppState 恢复 Excel 状态。
```

### 4.2 跨市场指标表

```text
1. BuildCrossMarketIndicatorSheet 读取 4 张分市场指标表。
2. CollectCompaniesFromIndicatorSheet 收集公司和报告期区块。
3. 跨市场_指标表 横向铺 公司×报告期。
4. 每个数值 cell 写公式引用分市场指标表。
5. A1 注释写入显示币种和 QA 提示。
6. RunDataQualityChecks 追加 GLOBAL_QA 三行。
```

### 4.3 RMB toggle

```text
1. 样本池 E6 = 原币 / 统一RMB。
2. WriteWideTable 写 raw dump layer 和展示公式。
3. 原币模式直接显示 raw dump。
4. 统一RMB 模式调用 GetFxFromSheet(currency, period, EOP/AVG)。
5. 用户手改「汇率」sheet 后,公式随重算刷新。
6. 汇率缺失时显示空值,并写 FX_MISSING 诊断。
```

### 4.4 发布清理

```text
CleanReleaseWorkbook
├─ 清样本池 cookie
├─ 清三张诊断历史
├─ 清 .cache/
└─ 提醒用户用 Excel 内置检查器清作者元数据
```

## 5. 关键 Invariants

这些不变量是后续修改必须维持的边界。

### 5.1 VBA API

- `GetFxRate(curCode, periodEnd, useEop)` 签名必须保留。
- `GetFxRateStatus(curCode, periodEnd, useEop, outRate)` 签名必须保留。
- `GetFxFromSheet(currencyCode, periodEnd, rateKind)` UDF 签名必须保留。
- `EdgarHttpGet / XueqiuHttpGet / StockAnalysisHttpGet` 签名必须保留。
- `RunCachedHttpGet` 的遥测语义必须保留。
- `THttpResult` 字段语义必须保留。
- `BuildCrossMarketIndicatorSheet` 仍只生成 `跨市场_指标表`。

### 5.2 Sheet 和布局

- 样本池 Row 14+ 是用户录入区,装表和布局刷新不能清空。
- 诊断 sheet 是 17 列结构,L:Q 必须保持文本格式。
- `跨市场_指标表` 只保留指标合表,不恢复已删除的 BS/IS/CF 合表。
- `汇率` sheet 8 列结构不能改。
- 分市场 16 张正式表由抓数写入,显隐按钮只控制正式表。
- 诊断 sheet 默认隐藏,全局显隐按钮不展开诊断。

### 5.3 数据准确性

- RMB/CNY 永远短路为 1.0。
- 非 RMB 汇率缺失不能 fallback 为 1。
- 资产负债表换汇使用期末汇率。
- 利润表和现金流量表换汇使用期间平均汇率。
- stockanalysis 中概美股 fallback 仅对白名单测试集自动触发。
- QA 只追加诊断,不阻断写表。

### 5.4 运行状态

- 数据入口必须用 `BeginAppState / EndAppState` 包夹。
- 错误路径必须恢复 `Calculation / EnableEvents / DisplayAlerts / ScreenUpdating / StatusBar`。
- toggle 入口不加 AppStateGuard,保持轻量。
- fetch helper 内部不直接管理全局 Excel 状态。

## 6. 关键文件路径

```text
VBA Captor/
├── 上市公司财务数据查询.xlsm
├── README.md
├── STATUS.md
├── ARCHITECTURE.md
├── PHASE_4N_PLAN.md
├── modules/
│   ├── JsonConverter.bas
│   ├── 模块_AppStateGuard.bas
│   ├── 模块_工具函数.bas
│   ├── 模块_总入口.bas
│   ├── 模块_抓汇率.bas
│   ├── 模块_抓资产负债表.bas
│   ├── 模块_抓利润表.bas
│   ├── 模块_抓现金流量表.bas
│   ├── 模块_抓指标表.bas
│   ├── 模块_抓美股财报.bas
│   ├── 模块_抓港股财报.bas
│   ├── 模块_抓韩股财报.bas
│   └── 模块_测试.bas
├── tools/
│   ├── build_template.py
│   ├── install_modules.py
│   ├── test_fx_live.py
│   ├── diff_phase4f_step3_lite.py
│   ├── inspect_phase4g_state.py
│   ├── inspect_phase4h_state.py
│   ├── inspect_phase4k_state.py
│   ├── inspect_phase4l_state.py
│   ├── inspect_phase4m_state.py
│   └── run_offline_tests.py
├── tests/
│   └── fixtures/
│       ├── .gitignore
│       └── README.md
├── samples/
└── .cache/
```

`modules/*.bas` 是 VBA 源码真源。`tools/install_modules.py` 会把这些模块注入到 xlsm。

`tests/fixtures/` 只提交 scaffold。fixture payload 由 `tools/run_offline_tests.py` 本地生成,不提交到 git。

`.cache/` 是 HTTP 响应缓存目录,不提交到 git。

## 7. 数据源声明

| 数据范围 | 数据源 | 入口 | Cookie |
|---|---|---|---|
| A 股财报 | 新浪财经公开页面 | `HttpGet` | 不需要 |
| 美股 EDGAR | SEC companyfacts JSON | `EdgarHttpGet` / `CachedEdgarHttpGet` | 不需要 |
| 美股 fallback | 雪球 / stockanalysis | `CachedXueqiuHttpGet` / stockanalysis wrapper | 雪球需要 |
| 港股财报 | 雪球 HK API | `CachedXueqiuHttpGet` | 需要 |
| 韩股财报 | stockanalysis.com KRX 页面 | `StockAnalysisHttpGet` wrapper | 不需要 |
| 汇率 | 雪球 K 线 | `模块_抓汇率.EnsureFxRateCached` | 不需要 |

HTTP 请求原则:

- 只使用公开 GET 请求。
- 不绕过登录或权限控制。
- 不抓取私人数据。
- 不保存雪球 cookie 到 `.cache/`。
- SEC 请求有最小间隔控制。
- 失败响应不写入 HTTP cache。

## 8. Cache 分源 TTL

| Source | TTL(小时) | 用途 | 理由 |
|---|---:|---|---|
| `SEC_TICKER_MAP` | 168 | ticker 到 CIK 映射 | 变化很慢 |
| `EDGAR` | 24 | companyfacts JSON | 财报更新低频 |
| `EDGAR_COMPANYFACTS` | 24 | EDGAR 兼容别名 | 默认日级 |
| `XUEQIU` | 12 | 雪球财报 API | 页面/API 可能更频繁变化 |
| `XUEQIU_HK` | 12 | 港股雪球兼容别名 | 与雪球主路径一致 |
| `XUEQIU_US` | 12 | 美股雪球 fallback 兼容别名 | 与雪球主路径一致 |
| `XUEQIU_KR` | 12 | 韩股雪球历史兼容别名 | 与雪球主路径一致 |
| `STOCKANALYSIS` | 24 | stockanalysis 通用路径 | HTML 更新低频 |
| `STOCKANALYSIS_KR` | 24 | 韩股 stockanalysis | 日级足够 |
| `STOCKANALYSIS_US` | 24 | 中概美股 fallback | 日级足够 |
| `FX_KLINE` | 24 | 汇率 K 线 | 每日 close 后稳定 |
| default | 24 | 未识别 source | 保守兜底 |

TTL 由 `GetTtlHoursForSource(sourceName)` 统一返回。调用方仍通过 `RunCachedHttpGet` 读写缓存和遥测字段。

### 验证矩阵

常规收口回归:

```powershell
py tools/test_fx_live.py --skip-install
py -u tools/diff_phase4f_step3_lite.py
py -u tools/inspect_phase4g_state.py
py -u tools/inspect_phase4h_state.py
py -u tools/inspect_phase4k_state.py
py -u tools/inspect_phase4l_state.py
py -u tools/inspect_phase4m_state.py
py tools/run_offline_tests.py
```

每项职责:

- `test_fx_live.py`: 汇率 sheet、`GetFxRate` 兼容性、RMB/CNY 短路。
- `diff_phase4f_step3_lite.py`: A 股原币和统一 RMB 短路一致性。
- `inspect_phase4g_state.py`: 跨市场指标表和显隐按钮。
- `inspect_phase4h_state.py`: B6 实时 toggle、本地 cache、自动 fallback。
- `inspect_phase4k_state.py`: Score 文本、FX missing、live FX UDF。
- `inspect_phase4l_state.py`: HTTP/cache 遥测、SEC 限流、发布清理。
- `inspect_phase4m_state.py`: 离线 fixture、QA、分源 TTL。
- `run_offline_tests.py`: 8 个本地离线 smoke。

新增 phase 后应优先复用上述 frozen 驱动,只有 state 结构变化时才新增 inspect。

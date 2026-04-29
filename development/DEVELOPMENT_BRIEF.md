# FS Capture — 开发简报与代码审核委托书

> 给 Codex 的审核任务说明。本文件包含项目背景、架构决策、模块清单、已验证的链路、已知风险，以及希望审核重点关注的问题。

---

## 1. 项目背景

### 1.1 用户画像与原始诉求

用户是一名中国金融从业者（极可能是审计 / 卖方研究方向，多次提到"瑞华底稿"——瑞华是国内知名会计师事务所）。其工作流深度依赖 Excel + VBA。

用户提交的需求原文：

> 我想要开发一个一键抓取 A股 / 港股 / 美股 / 韩股上市公司**审计报告和年报（公司自己发布的）**的工具，具体的需求如下：
>
> 1. **交互界面**：用 EXE 打开的 GUI 界面，前端设计美观、扁平化风格、边框过渡用圆角而非矩形、色彩有设计感；
> 2. **界面内容与要实现的功能**：勾选 A股 / 港股 / 美股 / 韩国交易所，每个交易所下面有单独的 section，可以点击加行来加入需要抓取的公司，每行可以输入股票代码，等所有要抓的公司代码输入完毕后，点击一个按钮确认后就可以**自动生成公司全名给到用户去确认**；还需要有一个 section 确认需要抓取的报告口径：包含报告年份、季报、年报；
> 3. **自动抓取财务数据及财务指标的功能**：用户提供了一个基于 VBA 的 Excel 底稿（`新浪财经行业数据查询V2.2.xlsm`），但其局限是只能抓 A 股；建议升级或重做。

### 1.2 既有 VBA 工具的现状（被替代对象）

文件：`E:\Eric Nutshell\4. 效率工具\新浪财经行业数据查询V2.2.xlsm` (3.3 MB)

通过解压 ZIP + 提取 vbaProject.bin 字符串得到的现状：

- **数据源**：新浪财经 (`vip.stock.finance.sina.com`)
- **HTTP 客户端**：`WinHttp.WinHttpRequest.5.1`（Windows COM）
- **HTML 解析**：`htmlfile` COM + `VBScript.Regexp`（正则）
- **编码**：GB2312
- **覆盖范围**：仅 A 股（`sh600000` / `sz000001` 主板）；不含 H 股、ADR、北交所等
- **8 个 sheet**：使用说明 / 样本池 / 指标表 / 资产负债表 / 利润表 / 现金流量表 / 上市公司基本资料 / 瑞华底稿

主要痛点：单市场、HTML 正则脆弱、Windows-only COM 锁死、不下载 PDF。

### 1.3 通过 AskUserQuestion 锁定的关键决策

| 议题 | 决策 |
|---|---|
| GUI 框架 | **PySide6 + 自定义 QSS**（用户接受推荐） |
| 韩股 DART API Key | **由用户自行注册**并在设置中粘贴 |
| 多期间支持 | **年份区间 + 期间类型勾选**（如 2022→2024 + 年报 + 三季报 一次跑完） |
| 审计报告口径 | **完整年报 PDF + 独立审计报告 PDF（如有）** |

### 1.4 上下文

- 工作目录：`E:\Claude+CODEX Project\FS Capture`（本次开发前为空目录）
- 用户邮箱：`kaiyu199602@gmail.com`（已用作 SEC EDGAR User-Agent）
- 用户操作系统：Windows 11，Python 3.14.3
- 主要语言：中文（界面 / 错误提示均中文）
- 计划文件：`C:\Users\kaiyu\.claude\plans\e-eric-nutshell-4-v2-2-xlsm-a-woolly-pnueli.md`（已由用户审批通过）

---

## 2. 技术栈

| 层 | 库 | 版本 | 用途 |
|---|---|---|---|
| GUI | PySide6 | 6.11.0 | Qt6 绑定，无边框圆角窗口 |
| HTTP | httpx[http2] | 0.28.1 | 异步 HTTP，替代 VBA 的 WinHttp |
| HTML 解析 | selectolax | 0.4.7 | lxml 后端，比 BS4 快 5-10× |
| A 股数据 | akshare | 1.18.59 | 包装新浪 / 东方财富 / Tushare |
| 韩股数据 | OpenDartReader | 0.2.3 | DART OpenAPI 包装器 |
| Excel | openpyxl | 3.1.x | 读写 .xlsx |
| 数据校验 | pydantic | 2.13.3 | 模型 + 配置校验 |
| 重试 | tenacity | 9.1.4 | 指数退避 |
| 缓存 | diskcache | 5.6.3 | 持久化代码↔名称映射 |
| 日志 | loguru | 0.7.3 | 结构化日志 |
| 数据处理 | pandas | (akshare 依赖带入) | DataFrame |
| PDF | pypdf | 4.x | (预留，目前未使用) |
| 并发 | QThreadPool + tenacity | (内置) | 非阻塞 UI |
| 打包 | PyInstaller | 6.20.0 | one-folder 模式，~340 MB 产物 |

`pyproject.toml` 与 `requirements.txt` 已固化。

---

## 3. 项目结构

```
FS Capture/
├── pyproject.toml              # 依赖与打包元数据
├── requirements.txt            # pip 安装列表
├── config.toml                 # 用户配置（API Key、路径、并发、限速）
├── fs_capture.spec             # PyInstaller spec
├── README.md                   # 中文用户手册
├── run.bat                     # 源码启动（开发）
├── build.bat                   # 一键打包 EXE
│
├── app/
│   ├── __init__.py             # __version__
│   ├── main.py                 # QApplication 入口 + 顶层异常网 + crash.log
│   │
│   ├── core/                   # 与 GUI / 插件解耦的核心逻辑
│   │   ├── models.py           # Exchange/PeriodType/Period/Ticker/Company/StatementType/FinancialStatement/ReportFile
│   │   ├── settings.py         # config.toml 读写（pydantic + tomllib + tomli_w）
│   │   ├── ratelimit.py        # TokenBucket（同步+异步）+ RateLimiterRegistry
│   │   ├── http.py             # httpx Client 工厂 + tenacity 重试装饰过的 get_json/post_json/stream_to_file
│   │   ├── cache.py            # diskcache 单例
│   │   ├── job.py              # Job/TaskResult/TaskStatus dataclasses
│   │   └── orchestrator.py     # QThreadPool 调度器，跨 plugin 编排
│   │
│   ├── ui/                     # 全部 PySide6 部件
│   │   ├── main_window.py      # 无边框圆角窗口 + 自定义标题栏 + 边缘 hit-test 拖动/缩放
│   │   ├── main_view.py        # 内容主视图，组装所有 section 并监听 orchestrator 信号
│   │   ├── exchange_selector.py# 顶部 4 个 chip（A/HK/US/KR）
│   │   ├── exchange_panel.py   # 单交易所 section（内嵌 TickerRow 列表）
│   │   ├── ticker_row.py       # 单行：代码输入 + 确认 + 状态徽章 + 名称 + 删除
│   │   ├── period_selector.py  # 起止年份 + 期间类型 checkbox
│   │   ├── output_card.py      # 输出文件夹选择
│   │   ├── progress_dock.py    # 任务进度面板（per-task pill）
│   │   ├── settings_dialog.py  # DART Key、并发数、主题、SEC UA
│   │   └── styles/
│   │       ├── palette.py      # light/dark 调色板 + per-exchange accent
│   │       ├── qss_loader.py   # Template 替换
│   │       └── app.qss         # 280+ 行样式
│   │
│   └── exporters/
│       ├── excel_writer.py     # 8 sheet 工作簿
│       └── templates/          # 预留瑞华底稿模板位置（目前为占位）
│
├── plugins/                    # 各市场插件
│   ├── base.py                 # ExchangePlugin ABC
│   ├── __init__.py             # get_plugin(exchange) 路由
│   │
│   ├── ashare/                 # 已端到端验证（贵州茅台 600519）
│   │   ├── name_resolver.py    # akshare stock_info_a_code_name + cninfo orgId
│   │   ├── reports.py          # cninfo hisAnnouncement → PDF
│   │   └── financials.py       # akshare stock_*_em (yearly/by_report)
│   │
│   ├── hk/                     # 已验证（腾讯 00700）
│   │   ├── name_resolver.py    # 东方财富 stock/get 单股查询（替代不稳定的 stock_hk_spot_em）
│   │   ├── reports.py          # HKEXnews titlesearch.xhtml HTML 解析
│   │   └── financials.py       # akshare stock_financial_hk_report_em
│   │
│   ├── us/                     # 已验证（Apple AAPL）
│   │   ├── name_resolver.py    # SEC company_tickers.json
│   │   ├── reports.py          # SEC submissions API + 分页 files 兜底
│   │   └── financials.py       # SEC XBRL companyfacts + GAAP→中文映射（约 30 字段）
│   │
│   └── kr/                     # 已编码，未验证（缺 DART Key）
│       ├── name_resolver.py    # OpenDartReader.corp_codes
│       ├── reports.py          # DART list + document
│       └── financials.py       # DART finstate_all
│
├── output/                     # 抓取结果（运行时填充）
├── cache/                      # diskcache 持久化
├── logs/                       # loguru 日志
└── dist/FS Capture/            # PyInstaller 产物（27 MB EXE + 314 MB _internal/）
```

---

## 4. 关键架构模式

### 4.1 Plugin-per-Exchange

每个市场 = `plugins/{ashare,hk,us,kr}/` 模块，实现 `plugins/base.py` 的 `ExchangePlugin` ABC：

```python
class ExchangePlugin(ABC):
    exchange: Exchange

    @abstractmethod
    def resolve_name(self, code: str) -> Ticker: ...

    @abstractmethod
    def fetch_company(self, ticker: Ticker) -> Company: ...

    @abstractmethod
    def download_reports(self, ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]: ...

    @abstractmethod
    def fetch_financials(self, ticker: Ticker, period: Period) -> list[FinancialStatement]: ...
```

`plugins/__init__.py::get_plugin(Exchange)` 做懒加载。新增市场只需要新建一个目录。

### 4.2 编排与并发

`Orchestrator` (`app/core/orchestrator.py`)：
- 使用 Qt `QThreadPool` 跑任务，最大并发 = `settings.concurrency.max_workers`（默认 4）
- 每个 `TaskResult` 对应一个 `(Ticker, Period)` 二元组
- `_TaskRunnable` 在工作线程中按序调用：`resolve_name` → `download_reports` → `fetch_financials`
- 通过 `OrchestratorSignals` 把 task 状态推送给 UI（task_started / task_progress / task_finished / job_finished）

### 4.3 限速

`app/core/ratelimit.py::TokenBucket`：
- 每个数据源（cninfo / hkexnews / sec / dart / akshare / eastmoney）独立 bucket
- 同步与异步双 API
- 限速参数从 `config.toml` 读取（默认值见配置文件）

### 4.4 HTTP

`app/core/http.py`：
- `default_client(source=...)` 返回带默认头、HTTP/2、cookie jar 的 `httpx.Client`
- SEC 自动注入用户邮箱作为 User-Agent（合规）
- 三个核心函数：`get_json` / `post_json` / `stream_to_file`，全部 tenacity 装饰（3 次指数退避重试）

### 4.5 GUI 渲染

- `MainWindow` 是 `FramelessWindowHint + WA_TranslucentBackground` 的无边框透明底
- 内层 `#WindowRoot` 应用 `border-radius: 14px` 和 `QGraphicsDropShadowEffect`
- 自定义标题栏（logo + 副标题 + min/max/close 按钮）支持拖动
- 边缘 6px hit-test 实现 8 方向缩放（`mousePressEvent` / `mouseMoveEvent` / `mouseReleaseEvent`）
- 双击标题栏切换最大化；最大化时去掉圆角
- 全部样式在 `app/ui/styles/app.qss` 集中管理，token 替换由 `qss_loader.load_qss(palette)` 完成

---

## 5. 已端到端验证的链路

| 市场 | 测试代码 | 名称解析 | 报告下载 | 财务数据 |
|---|---|---|---|---|
| A 股 | 600519 | 贵州茅台 ✓ | annual_report.pdf 3.5 MB ✓ | 营收 ¥150.6B、净利润 ¥74.7B ✓ (匹配公开数据) |
| 港股 | 00700 | 腾讯控股 ✓ | annual_report.pdf 7.3 MB ✓ | （未在端到端测试中详细对账） |
| 美股 | AAPL | Apple Inc. ✓ | 10-K FY2023 + FY2024 (1.5 MB each) ✓ | 营业收入 $365.8B 等 30 字段 ✓ |
| 韩股 | — | 仅静态导入测试 | — | — |
| Excel 写出 | — | 8 sheet 工作簿，2 ticker × 1 期间 → 19.9 KB | — | — |
| EXE 启动 | — | 已确认窗口正常打开（修复 stderr / stdout 两次 PyInstaller windowed 坑后） | — | — |

---

## 6. 关键实现细节

### 6.1 A 股（最完整、最稳定）

**name_resolver.py**：
- 全市场代码↔名称映射通过 `akshare.stock_info_a_code_name()` 拉取，每天缓存一次
- cninfo orgId 通过 `topSearch/detailOfQuery` POST 获取（按需懒加载，每个 ticker 单独缓存 30 天）
- orgId 格式如 `gssh0600519` (沪市) / `gssz0000001` (深市)，少数公司返回 `GD165627` 等异常格式 — 直接当字符串透传

**reports.py**：
- 使用 cninfo `hisAnnouncement/query` POST API
- 期间类型 → category 映射：年报=`category_ndbg_szsh`、Q1=`category_yjdbg_szsh`、Q2=`category_bndbg_szsh`、Q3=`category_sjdbg_szsh`
- 公告时间窗口宽放（Q1 找 4-7 月公告，年报找次年 1-9 月）
- 选最终公告：标题须含 "年度报告" 与年份字符串、剔除 "摘要"/"更正"/"已取消"，按 `announcementTime` 倒序取最新
- PDF URL = `http://static.cninfo.com.cn/` + `adjunctUrl`

**financials.py**：
- 年报走 `stock_balance_sheet_by_yearly_em` 等 yearly 端点；季报走 `by_report_em`
- 列名兼容 `REPORT_DATE` / `报告期` / `报告日期` 三种 akshare 版本差异
- **目前 lines 字典 key 是英文 EM 字段名（如 `TOTAL_ASSETS`、`TOTAL_OPERATE_INCOME`）**，不是中文。这是 v1 已知简化，需要中英映射表才能完美兼容瑞华底稿

### 6.2 港股

**name_resolver.py**：
- 起初使用 `akshare.stock_hk_spot_em()` 拉全表 → **不稳定，多次 RemoteDisconnected**
- 改成单股查询：`https://push2delay.eastmoney.com/api/qt/stock/get?secid=116.{code}&fields=f57,f58`，可靠
- 代码归一化：去 `.HK` 后缀、去 `HK`/前导零，最后 `zfill(5)`

**reports.py**：
- HKEXnews 没有官方 API；使用 `https://www1.hkexnews.hk/search/titlesearch.xhtml`（GET）+ selectolax HTML 解析
- 期间类型 → t1code/t2code 映射（年报 40000/40100、中期 40200、Q3 40300、Q1 40400）
- 期间筛选靠标题关键词："年報"/"年度報告"/"中期"/"半年"/"Annual Report" 等
- PDF 直链来自 `<a href>`，文件名结尾 `_c.pdf`(中文) 或 `_e.pdf`(英文)
- ⚠️ **目前选 main filing 的逻辑较粗糙**：仅按"标题是否含年份"筛选；某些公司年报标题年份语义可能错位（截至 v1，验证案例腾讯返回了 `2023/1222/2023122200423_c.pdf`，标题"年報 2023"，但发布时间是 2023-12-22，该 PDF 实际对应的财年并未严格校验）

**financials.py**：
- akshare `stock_financial_hk_report_em(stock=, symbol=, indicator=)`
- `indicator` 取 `年度`（年报）或 `报告期`（半年/季度）
- 同样有英文 EM 字段名问题

### 6.3 美股

**name_resolver.py**：
- 来源 `https://www.sec.gov/files/company_tickers.json`，每天缓存
- `external_id` 存 10 位补零的 CIK
- 支持 `BRK.B` → `BRK-B` 变体

**reports.py**：
- SEC submissions API：`https://data.sec.gov/submissions/CIK{cik}.json`
- 表单类型映射：年报=10-K/10-K/A/20-F/20-F/A/40-F；Q1/Q2/Q3=10-Q + Q2 加 6-K
- ⚠️ **重要陷阱**（已修复）：SEC 字段是 `reportDate` 而非文档普遍提到的 `periodOfReport`
- ⚠️ **分页 files 数组兜底**：`recent.*` 只存最近 ~1000 条；老 ticker（如 AAPL 高频 8-K）历史数据需要走 `filings.files[].name` 分页加载
- 优先非 amendment（`/A` 排后），按 filingDate 升序取最早

**financials.py**：
- SEC XBRL companyfacts API：`https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json`
- `_GAAP_MAPPING` 是手工维护的 (statement, 中文名) → [GAAP tags] 映射表，目前约 30 项核心字段
- 期间选取：根据 `fp` (FY/Q1/Q2/Q3) 和 `fy` 匹配
- 单位优先 USD，否则取首个

### 6.4 韩股（未端到端验证）

**name_resolver.py**：
- `OpenDartReader.corp_codes` 返回全 DataFrame，按 `stock_code` 索引为 `external_id=corp_code`
- 缓存 7 天
- ⚠️ DART Key 必填，否则 `_dart()` 直接抛 ValueError

**reports.py**：
- `dart.list(corp, start, end, kind='A')` 拉公告列表
- 通过 `OpenDartReader.document(rcept_no)` 下载（库默认存当前工作目录，我们用 `os.chdir` 切换 CWD —— ⚠️ **线程不安全**，详见 §8）
- 年度期间额外尝试拉 `감사보고서`（独立审计报告）

**financials.py**：
- `dart.finstate_all(corp, bsns_year, reprt_code)`
- `sj_div`：BS / IS / CIS / CF / SCE → 我们的 `StatementType`
- account_nm 是韩文，账目名直接用韩文存

### 6.5 GUI

**main_window.py**：
- 无边框 + 自定义标题栏 + 8 方向边缘缩放
- 已知问题：丢失 Win11 aero-snap（拖到屏幕边缘自动半屏的功能）
- 图形效果用 `QGraphicsDropShadowEffect`，性能影响轻微

**main_view.py**：
- 单一 `MainView` 持有 4 个 `ExchangePanel`，根据 `ExchangeSelector` 的勾选状态显示/隐藏
- 默认勾选 A 股
- "开始抓取" 走完三道校验：tickers 非空、有期间勾选、KR 时检查 DART Key
- job_finished 信号回调里调用 `excel_writer.write_workbook` 同步生成工作簿

**ticker_row.py**：
- "确认" 按钮通过 `_ResolveSignals` + `_ResolveRunnable` 在工作线程做名称解析（避免阻塞 UI）
- 修改输入框内容会自动让 ticker 失效，要求重新确认
- 状态徽章 (`StatusPill`) 通过 `dynamic property + style().unpolish/polish` 切换 QSS 样式

**progress_dock.py**：
- 接收 orchestrator 发出的 `task_started/progress/finished`，按任务渲染一行
- 使用 `QScrollArea` 包装，任务多时可滚动
- 进度条按完成数 / 总数填充

### 6.6 Excel 写出

8 个 sheet 顺序：样本池 / 上市公司基本资料 / 资产负债表 / 利润表 / 现金流量表 / 指标表 / 报告下载汇总 / 瑞华底稿（占位）

- 三大报表的列布局：第 1 列是项目名，2..N 列每列是一个 (ticker × period) 组合
- 项目行是所有 ticker 出现过的字段并集；缺失值留空
- 指标表手算 ROE / 毛利率 / 资产负债率 / 营收(亿) / 净利润(亿) / 总资产(亿) / 总负债(亿)
- 字段名查找时同时尝试中文（"营业收入"）、英文 EM key（`TOTAL_OPERATE_INCOME`）、GAAP（`Revenues`）
- 样式：标题行紫色 (#4338CA) 反白、数字 `#,##0.00;(#,##0.00);-`、freeze panes
- ⚠️ **未真正克隆瑞华底稿模板**——最后一个 sheet 是占位说明文字。计划中需要从原 .xlsm 转 .xlsx 后置入 `app/exporters/templates/ruihua_v2.xlsx`

### 6.7 PyInstaller 打包

`fs_capture.spec`：
- one-folder 模式（一文件模式启动慢）
- `collect_all("akshare")` + `collect_all("OpenDartReader")` + `collect_data_files("pykrx")`
- `excludes` 砍掉 matplotlib/scipy/torch/QtWebEngine 等大头
- 最终 dist/FS Capture/ 目录 ~341 MB，FS Capture.exe 是 27 MB 启动器

⚠️ **PyInstaller windowed 模式踩了两个坑**：
1. `sys.stderr is None` → loguru 加 sink 时 `TypeError`（已加 None 守护）
2. `sys.stdout is None` → akshare 内部 tqdm 进度条 `.write()` 崩溃（已在 main.py 顶部把 None 流重定向到 `os.devnull`）

---

## 7. 用户配置 / 运行时

`config.toml`（项目根，从 settings.py 读写）：

```toml
[paths]
output_dir = "output"
cache_dir = "cache"
log_dir = "logs"

[concurrency]
max_workers = 4

[rate_limits]
cninfo = 5
hkexnews = 3
sec = 8
dart = 5
akshare = 4

[sec]
user_agent = "FS Capture (kaiyu199602@gmail.com)"

[dart]
api_key = ""

[ui]
theme = "light"
language = "zh_CN"
window_width = 1280
window_height = 820
```

PyInstaller 冻结后，`config.toml` 与 `output/` 等路径均相对 `sys.executable` 解析（见 `settings.py::project_root`）。

---

## 8. 已知问题与未覆盖路径（请审核重点）

### 8.1 高优先级

1. **KR 插件完全未运行时验证** — DART API Key 缺失导致整条链路从未实跑。审核请重点看：
   - `plugins/kr/reports.py::_download_filing` 用 `os.chdir` 切换 CWD —— **多线程并发下不安全**！QThreadPool 默认 4 worker，多个 KR 任务并行会互相覆盖 CWD
   - DataFrame 列名假设（`rcept_no`、`report_nm`、`rcept_dt`）是否真的是 OpenDartReader 当前版本的列名
   - `OpenDartReader.document()` 实际产物（PDF? ZIP? 单文件还是多文件？）与我们的 `target_dir.glob(f"{rcept_no}*")` 是否对得上

2. **HK 报告选取逻辑过于宽松**
   - `_select_main` 仅按"标题含年份字符串"筛选 → 同一年份的多份关联文件（补充公告、ESG 报告等）可能误选
   - 财年 vs 公告年份混淆：港股可有非 12 月财年（如汇丰、阿里巴巴），当前代码对腾讯返回了 `2023/1222/...` 即 2023-12-22 公布的"年報 2023"，可疑
   - 建议：补充对 PDF 标题做更严格的正则；或把所有候选都下载并附 metadata 让用户自己判断

3. **akshare 字段名是英文 EM 命名**
   - 三大报表中 line 名称如 `TOTAL_ASSETS` / `TOTAL_OPERATE_INCOME` / `PARENT_NETPROFIT` 直接进 Excel
   - 不符合用户瑞华底稿（中文行）的工作流
   - 需要补一份英→中映射表，建议放在 `plugins/ashare/translations.py` 与 `plugins/hk/translations.py`，在写出前 normalize

4. **SEC submissions 的分页 files 兜底未在生产数据上验证**
   - 改完 `reportDate` 后只在 AAPL FY2023/2024 上跑过（这两个都在 recent.* 里）
   - 老年份（FY2019 之前）会落到 `filings.files[]` 分页路径上 — 这条路径未实测
   - 单测建议：用 `MSFT` 跑 FY2018，强制走分页

### 8.2 中优先级

5. **orchestrator 信号断连时机**
   - 修复了 `submit_job` 跨 job 信号叠加 bug
   - 但 `cancel_requested` 信号在 `progress_dock` 中定义却**没有任何接收者** —— 用户点不了取消
   - 同时 `_TaskRunnable.run` 没有 cancellation token，跑到一半的下载无法中断

6. **httpx.stream 超时**
   - `stream_to_file` 用全局 `timeout=30s`，但港股年报 PDF 7+ MB、有些美股 10-K HTML 文档更大，慢网下可能 timeout
   - 建议：分两段超时（连接 30s + 流读 0=无超时）

7. **资源句柄**
   - `app/core/cache.py::get_cache()` 是 `lru_cache(1)` 单例，从未关闭 → 退出时可能留下 lock 文件
   - 建议：在 `QApplication.aboutToQuit` 信号挂上 `cache.close()`

8. **GUI 文本拷贝/选中**
   - `name_label` 是 `QLabel`，不能选中文本拷贝。审计场景下用户经常需要复制公司全名 → 建议改为 `QLineEdit` 只读 或 `QLabel.setTextInteractionFlags(Qt.TextSelectableByMouse)`

9. **首次启动 onboarding**
   - 用户第一次打开 EXE 时没有任何引导（没有"输入第一个股票代码试试"的提示）
   - DART Key 缺失 + 用户勾选了韩股 + 还没注册 — 当前会用 QMessageBox 弹"是否打开设置"，但没解释如何注册

### 8.3 低优先级

10. **bjse 北交所代码识别不全**：`reports.py::_column_for` 用 `bse` 但实际 cninfo 用 `bj`。需 cross-check
11. **韩股 Q3 vs Q1 撞车**：DART 季度报告类型相同（`분기보고서`），代码里 Q3 keywords 没区分；当年内出 2 份分季报告时可能误选
12. **Excel `_safe_div` 把 None 当 0** — 不一定符合财务语义
13. **frameless window 边缘 hit-test** 在多显示器/不同 DPI 下未测试
14. **README 提到的 playwright 兜底**实际未在 HK 插件中触发
15. **`OpenDartReader` 在 PyInstaller 下的资源文件**——`collect_all("OpenDartReader")` 是否覆盖了它内部的 `corpCode.xml` ZIP 加载逻辑，需要冷启动测试

---

## 9. 给 Codex 的具体审核请求

请按以下顺序审核，每一类都给出**具体文件:行号引用**和**修复建议**：

### 9.1 正确性 / Bug

- **§8.1 全部 4 个高优先级项**逐一过一遍代码
- 重点：`plugins/kr/reports.py::_download_filing` 的 CWD 线程安全
- `plugins/us/reports.py::_filter_table` 的 ANNUAL 期间筛选——`reportDate` startswith 年份字符串是否会匹配跨年的 10-K（如 Apple FY2023 报告日 2023-09-30，但封装期是 2022-09-25 至 2023-09-30，"年份"语义模糊）
- `app/core/orchestrator.py` 的 task_finished 信号与 main_view._on_job_finished 的连接生命周期（特别是用户关闭窗口时是否正确清理）
- `app/ui/main_window.py::_perform_resize` 在 RTL 输入法/拖到屏幕边缘的边界条件

### 9.2 安全

- HTTP 请求的 SSL 证书校验是否禁用（确认全部用 default `verify=True`）
- 是否有 SSRF 风险（用户输入的 ticker code 是否会被拼进 URL —— 全局 grep `f"{...code...}"` 模式）
- DART API Key 存 `config.toml` 明文 → 是否需要平台 keyring？v1 接受这个权衡，但请提示
- 用户邮箱硬编码到 SEC User-Agent 是合理（SEC 政策要求），但代码中是否有其他地方泄露邮箱

### 9.3 健壮性 / 错误处理

- 网络错误时 GUI 是否正确回到可交互状态（progress dock 显示失败、run_btn 重新可点）
- 部分 task 失败、其他 task 成功时 Excel 是否正确生成（已写但未端到端测）
- Excel 写出失败时是否会冒泡到用户（main_view 已 try/except，但只展示给主对话框）
- 缓存损坏（手动删 cache 目录、cache 文件被锁）时的兜底
- DART API Key 错误时给用户的中文提示是否清晰

### 9.4 性能

- 1 ticker × 12 期间（3 年 × 4 季度）的并发是否合理 —— akshare 限速 4 rps × 3 库 = 每个 ticker ~9 秒，12 期间 ~108 秒
- `_load_name_map` 在多线程下首次调用是否有惊群效应（多个 worker 同时见到 cache miss）—— 当前没有锁
- `app/exporters/excel_writer.py` 在 100+ ticker × 12 期间 × 200+ 字段下的内存占用与生成速度

### 9.5 可维护性 / 代码质量

- 类型标注是否一致（部分文件用 `from __future__ import annotations` + 字符串类型，部分直接用具体类型）
- `Exchange` enum 用法在 `_FORMS_BY_PERIOD` 等映射里直接拿 PeriodType 当 key，跨模块复用是否清晰
- `plugins/{ashare,hk,us,kr}/__init__.py` 4 个文件几乎是 boilerplate 复制 —— 是否值得抽到 base？
- 命名一致性：`fetch_company` vs `download_reports` vs `fetch_financials` 动词不统一（fetch/download）

### 9.6 用户体验（中文金融用户视角）

- 错误提示中文是否符合金融行业用语（避免 "解析失败"，更专业的是 "代码未识别" / "暂未收录"）
- 进度面板的 status pill 文本是否清晰
- 默认期间（最近 3 年年报）是否符合用户预期
- 输出工作簿命名 `FS_Capture_{timestamp}.xlsx` 是否符合用户习惯（建议改 `底稿_{tickers首项}_{period}.xlsx`）

### 9.7 PyInstaller / 部署

- `fs_capture.spec` 里 `excludes` 是否过激（漏排了什么必要包？）
- 打包后冷启动一次 EXE，确认 4 个 plugin 都能完成首次 cache 加载（A 股已验证；HK 单股查询已验证；US 已验证；**KR 未验证**）
- `config.toml` 在 frozen 模式下的写入路径（`Path(sys.executable).parent`）是否真的可写（受限的 Program Files 路径会失败）
- crash.log 写入路径与 _show_fatal MessageBox 流程在打包后是否真能弹出

---

## 10. 已修复但应留心的回归风险

| 修复 | 文件 | 描述 |
|---|---|---|
| sys.stderr is None | `app/main.py:16-25` | windowed 模式 loguru sink 守护 |
| sys.stdout is None | `app/main.py:8-14` | tqdm 等流写入兜底 → `os.devnull` |
| SEC reportDate vs periodOfReport | `plugins/us/reports.py` | 字段名错误导致全部美股 10-K 找不到 |
| HKEXnews titleSearchServlet vs titlesearch.xhtml | `plugins/hk/reports.py:30` | 之前的 servlet URL 已 404 |
| akshare HK bulk 接口不稳 | `plugins/hk/name_resolver.py` | 改成单股 stock/get 端点 |
| Orchestrator 跨 job 信号叠加 | `app/core/orchestrator.py::submit_job` | 新版本带 disconnect + 任务 ID 过滤 |

---

## 11. 文件 LOC 速览

```
app/main.py                              ~100
app/core/models.py                        ~90
app/core/settings.py                      ~95
app/core/orchestrator.py                  ~95
app/core/http.py                          ~80
app/core/ratelimit.py                     ~70
app/core/job.py                           ~35
app/core/cache.py                         ~10
app/ui/main_window.py                    ~200
app/ui/main_view.py                      ~190
app/ui/ticker_row.py                     ~125
app/ui/exchange_panel.py                 ~110
app/ui/exchange_selector.py              ~110
app/ui/period_selector.py                 ~95
app/ui/output_card.py                     ~70
app/ui/progress_dock.py                  ~140
app/ui/settings_dialog.py                 ~55
app/ui/styles/palette.py                  ~70
app/ui/styles/qss_loader.py               ~30
app/ui/styles/app.qss                    ~280
plugins/base.py                           ~50
plugins/__init__.py                       ~25
plugins/ashare/{name_resolver,reports,financials}.py ~330
plugins/hk/{name_resolver,reports,financials}.py     ~290
plugins/us/{name_resolver,reports,financials}.py     ~330
plugins/kr/{name_resolver,reports,financials}.py     ~290
app/exporters/excel_writer.py            ~365
fs_capture.spec                           ~70
=== 合计 ~3900 行 Python + ~280 行 QSS ===
```

---

## 12. 审核交付期望

请输出：
1. **Critical issues**（必修）— 必须修复才能上生产
2. **High** — 用户使用前几次内大概率会撞到的问题
3. **Medium / Low** — 改善代码质量和长期可维护性
4. **格式**：每条带文件:行号、问题描述、根因、建议补丁（diff 或伪代码）
5. **不需要重写整个项目**：定向修复优于推倒重来
6. **如发现需求理解偏差**（基于 §1 项目背景），请明确指出哪一处与用户原意不符

— 委托方：Claude（开发执行者） / 项目所有者：Eric (kaiyu199602@gmail.com)

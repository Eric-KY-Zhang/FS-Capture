# Filings Atlas / 全球披露图谱 — 架构

> Last updated: 2026-05-24 (v1.0 — 8 市场 + SGX + 性能优化 + 首发 release)

本文档描述 Filings Atlas / 全球披露图谱（原 FS Capture）的 Python/PySide6 桌面工具架构，与同仓 `VBA Captor/` 子项目（旧 Excel/VBA 工具，已 v1.0 release）并存。本文档面向后续维护者和 Codex/Reviewer 交接使用，不替代 `development/DEVELOPMENT_BRIEF.md` 的开发委托书细节。

## 1. 项目目标

本工具是给财务、审计和卖方研究专业人士使用的本地 EXE 桌面工具。

核心目标（v0.2 重定位后）：

- 一键下载 A 股 / 港股 / 美股 / 韩股 / 台股 / 日股 / 英股 / 新加坡上市公司的**官方披露文件**：年报、审计报告、季报、半年报、IPO 招股书。
- 单 EXE 双击运行，配置（DART Key、并发、限速）走 `config.toml`。
- 输出扁平：所有 PDF 平铺在 `output/` 下，文件名内嵌交易所 / 代码 / 年份 / 期间 / 报告类型元信息。
- 无服务端依赖，所有 HTTP / 缓存 / 限速都在本地。

**明确不做**（重定位边界）：

- 不抓三大报表数字（资产负债 / 利润 / 现金流量）。
- 不算财务指标（ROE / 毛利率等）。
- 不生成 Excel 底稿（任何 .xlsx 输出都不在范围）。
- 不做跨市场对标。
- 不是 Web 服务、不部署到云、不做证券行情终端、不抓实时价格、不写入外部数据库、不做交易建议。

设计取舍：

- GUI 用 PySide6 + 自定义 QSS（无边框圆角窗口，扁平化设计），不用 Tkinter / Electron。
- 每市场一个 plugin 模块，新增市场只需要新建目录。
- HTTP 走 `httpx` + `tenacity`，不再用 VBA 的 WinHttp + 正则。
- 韩股 DART API Key 变为可选加速器；未配置时走 DART 公网披露页 fallback。
- 日股 EDINET 支持双模式，但 v1.0 **强烈推荐配置 EDINET Subscription-Key**（免费注册 `disclosure2.edinet-fsa.go.jp` / `api.edinet-fsa.go.jp`）。未配置时仅尝试 EDINET 公网页 fallback，官方 API 会返回 invalid subscription key。
- 英股 UK/FCA NSM 不需要 Key。
- 新加坡 SGXNet 不需要 Key，走公网 JSON API + SGX 附件 HTML 二跳解析，不需要 Playwright。
- 台股 MOPS 服务端证书缺失 Subject Key Identifier 扩展，`app/core/http.default_client` 仅对 `source=twse` 关闭证书校验（用户授权），其他七市场保留 certifi 严格校验。
- 抓数据 / Excel 装填的需求请用 `VBA Captor/`（已 v1.0 release，4 市场已稳定）；FS Capture 只解决「批量拿到原始 PDF」一个问题。

跟 `VBA Captor/` 的关系：

- VBA Captor 仍是 Excel-only 用户的快速路径（4 市场 + 跨市场指标表 + 实时 B6 toggle）。
- Filings Atlas 提供「报告 PDF 下载」+「不依赖 Excel COM」+「现代 GUI」。
- 两者数据源有大量重叠（A 股新浪/akshare、美股 SEC、港股雪球/HKEXnews），但定位互补：VBA Captor = 数据，FS Capture = 文件。

## 2. 顶层结构

```text
项目根/
├── README.md                # 给最终用户看的使用说明
├── ARCHITECTURE.md          # 本文档
├── PROJECT_RETROSPECTIVE.md # 复盘
├── roadmap/archive/         # v0.1→v0.5 sprint 规划归档 + 历史 planner 日志
├── config.example.toml      # 用户配置模板（不含密钥）
├── VBA Captor/              # 旧 Excel/VBA 工具（v1.0 release，仅作历史参考）
└── development/             # Python EXE 源码 + 构建脚本 + 测试 + 开发委托书
    ├── DEVELOPMENT_BRIEF.md # Codex 代码审核委托书
    ├── pyproject.toml
    ├── requirements.txt
    ├── filings_atlas.spec   # PyInstaller spec（one-folder 模式）
    ├── run.bat              # 源码启动（开发）
    ├── build.bat            # 一键打包 EXE
    ├── app/                 # GUI + 核心逻辑
    ├── plugins/             # 8 市场插件
    └── tests/               # unittest 用例
```

发布版（项目根直接 ship）会把 `development/` 折叠成隐藏目录，根上只留 `Filings Atlas.exe` + `_internal/` + `output/` + `cache/` + `config.toml`。

## 3. 模块依赖图（Python）

```text
GUI 入口
└─ app/main.py
   └─ QApplication + 顶层异常网 + crash.log

UI 层
└─ app/ui/main_window.py        无边框圆角窗口 + 自定义标题栏 + 8 方向 hit-test
   └─ app/ui/main_view.py       主视图，组装所有 section 并监听 orchestrator
      ├─ exchange_selector.py   顶部 8 个 chip (A/HK/US/KR/TW/JP/UK/SG)
      ├─ exchange_panel.py      单交易所 section（内嵌 ticker_row 列表）
      ├─ ticker_row.py          单行：代码输入 + 确认 + 状态徽章 + 名称 + 删除
      ├─ period_selector.py     起止年份 + 期间类型 checkbox（年报 / 季报 / 半年报 / IPO 招股书）
      ├─ output_card.py         输出文件夹选择
      ├─ progress_dock.py       任务进度面板（per-task pill）
      ├─ settings_dialog.py     DART Key / 并发数 / 主题 / SEC UA
      └─ styles/
         ├─ palette.py          light/dark 调色板 + per-exchange accent
         ├─ qss_loader.py       Token 替换
         └─ app.qss             280+ 行集中样式

核心层（与 GUI / 插件解耦）
└─ app/core/
   ├─ models.py             Exchange / PeriodType / Period / Ticker / Company / ReportFile
   ├─ settings.py           config.toml 读写（pydantic + tomllib + tomli_w）
   ├─ http.py               httpx Client 工厂 + tenacity 装饰过的 get_json / post_json / stream_to_file
   ├─ ratelimit.py          TokenBucket（同步+异步） + RateLimiterRegistry
   ├─ cache.py              diskcache 单例 + cached_or_load 单飞缓存
   ├─ output_paths.py       safe_filename + report_output_path + stale .part 清理
   ├─ pdf_renderer.py       Playwright 线程本地 browser/context pool + URL/PDF 渲染
   ├─ job.py                Job / TaskResult / TaskStatus dataclass
   └─ orchestrator.py       QThreadPool 调度器 + OrchestratorSignals

插件层
└─ plugins/
   ├─ base.py               ExchangePlugin ABC（3 方法）
   ├─ __init__.py           get_plugin(Exchange) 懒加载路由
   ├─ ashare/               akshare + cninfo
   ├─ hk/                   东方财富 + HKEXnews
   ├─ us/                   SEC submissions
   ├─ kr/                   OpenDartReader (DART) + DART public web fallback
   ├─ tw/                   TWSE ISIN + MOPS
   ├─ jp/                   EDINET API v2 + EDINET public web fallback
   ├─ uk/                   FCA NSM public search + artefact download/render
   └─ sg/                   SGXNet public API + attachment download
```

依赖方向：

- UI 只依赖 `app/core/` 与 `plugins/__init__.py::get_plugin`，**不直接 import** 具体 plugin。
- `plugins/*/` 只依赖 `app/core/{http,cache,ratelimit,models}`，**不反向依赖 UI**。
- `orchestrator` 用 `from plugins import get_plugin` 在 worker 线程内懒加载，避免主线程 import 阻塞。
- 不存在 `app/exporters/`（v0.1 → v0.2 重定位时删除）。

## 4. 插件契约（ExchangePlugin ABC）

每个市场 = `plugins/{ashare,hk,us,kr,tw,jp,uk}/` 模块，实现 `plugins/base.py` 的 `ExchangePlugin` ABC：

```python
class ExchangePlugin(ABC):
    exchange: Exchange

    @abstractmethod
    def resolve_name(self, code: str) -> Ticker: ...
        # 失败必须 raise ValueError（不能返回 None）

    @abstractmethod
    def fetch_company(self, ticker: Ticker) -> Company: ...
        # 行业 / 上市日期 / 币种等基础资料

    @abstractmethod
    def download_reports(
        self, ticker: Ticker, period: Period, output_root: Path
    ) -> list[ReportFile]: ...
        # 返回 [] = 该期间无披露（非异常）
```

各市场实现概览：

| 市场 | 名称解析 | 报告下载 | 端到端验证 |
|---|---|---|---|
| A 股 | akshare `stock_info_a_code_name` + cninfo orgId（30d 缓存） | cninfo `hisAnnouncement/query` POST → 静态 PDF | ✓ 600519 贵州茅台 |
| 港股 | 东方财富 `stock/get` 单股查询 | HKEXnews `titlesearch.xhtml` HTML 解析 | ✓ 00700 腾讯 |
| 美股 | SEC `company_tickers.json`（24h 缓存） | SEC `submissions API` + 分页 `files` 兜底 | ✓ AAPL 10-K FY2023/2024 |
| 韩股 | OpenDartReader `corp_codes` 或 DART 公网按需搜索（无 TTL 缓存） | DART OpenAPI `list` 或公网 HTML 搜索 + 公告 PDF/渲染下载 | ✓ 005930 三星电子（有 Key / 无 Key） |
| 台股 | TWSE ISIN 服务（30d 缓存，覆盖上市 + 上柜） | MOPS `doc.twse.com.tw/server-java/t57sb01`（POST step=9 → 解析临时 `/pdf/` 链接 → GET） | ✓ 2330 台积电 年报 / 半年报 / IPO |
| 日股 | EDINET 文档列表 API 或 EDINET 公网页 fallback | EDINET document API / PDF 展示链路 | API Key 路径单测覆盖；无 Key 官方 API 阻断 |
| 英股 | 预置主流 LSE ticker + FCA NSM issuer metadata | FCA NSM search API + artifact PDF/HTML 渲染 | ✓ ULVR / HSBA / AZN |

## 5. 数据流

### 5.1 公司名称识别（UI 阻塞确认）

```text
1. 用户在 ticker_row 输入代码 → 点"确认"
2. _ResolveRunnable 在工作线程调 plugin.resolve_name(code)
3. 成功：Ticker.name 显示在行内、status pill 变绿
4. 失败：raise ValueError → status pill 红 + tooltip 错误信息
5. 用户改 code 输入 → 状态自动失效，必须重新确认
```

确认是阻塞性必须：未确认的行不能进抓取队列（main_view 校验）。

### 5.2 一键抓取（并发 PDF 下载）

```text
 1. main_view 收集所有已确认 Ticker × 期间 → 构造 Job
 2. orchestrator.submit_job(Job)
 3. QThreadPool 按 settings.concurrency.max_workers (默认 4) 并发起 _TaskRunnable
 4. 每个 task 顺序：
       resolve_name (兜底，已确认行会跳过)
       fetch_company (失败仅 warning，不阻断)
       download_reports → 文件流式写盘
 5. 每步推 task_progress 信号给 progress_dock
 6. task_finished 信号 → progress_dock 行更新 + 计数累加
 7. 所有 task 完成 → job_finished 信号 → main_view 提示 "已下载 N 个 PDF，输出在 {output_dir}"
```

并发隔离：`Orchestrator.submit_job` 用 task_id set 隔离不同 job 的回调（见 `orchestrator.py`），防止跨 job 信号叠加。

### 5.3 输出文件名

`output_paths.py::report_output_path`：

```text
{exchange}_{code}_{name}_{year}_{kind_zh}.pdf
例：A_600519_贵州茅台_2024_年报.pdf
   US_AAPL_Apple Inc._2024_年报.pdf
   UK_ULVR_Unilever PLC_2024_年报.pdf
```

扁平铺在 `output_root` 下，**不再按公司或市场嵌套**（test_output_layout.py 锁定）。`safe_filename` 会替换 `< > : " / \ | ? *` 与 `\x00-\x1f` 为下划线。

`output/` 下**只有 PDF**，不会有 .xlsx 或其他衍生产物。

## 6. 关键 Invariants

这些不变量是后续修改必须维持的边界。

### 6.1 Plugin 接口

- `ExchangePlugin` ABC 3 个方法签名固定（`plugins/base.py`）。
- `resolve_name` 失败必须 `raise ValueError`，不能返回 `None`。
- `download_reports` 返回 `[]` 表示该期间无披露（非异常）。
- 新增市场只需加 `plugins/{exchange}/` 目录 + 在 `plugins/__init__.py::get_plugin` 注册。

### 6.2 数据模型（`app/core/models.py`）

- `Ticker.code`：validator 已 `strip().upper()`。
- `Period.year`：1990–2100。
- `PeriodType` 枚举：`ANNUAL` / `Q1` / `Q2`（半年报）/ `Q3` / `IPO_PROSPECTUS`（如已加入）。
- `Ticker.external_id`：cninfo orgId / SEC CIK / DART corp_code / HKEX stock_id 之一，按市场。
- `ReportFile.kind`：`"annual_report"` / `"audit_report"` / `"q1_report"` / `"q3_report"` / `"interim_report"` / `"ipo_prospectus"` 之一。

### 6.3 编排（`app/core/orchestrator.py`）

- 状态机：`PENDING → RESOLVING → DOWNLOADING → DONE/FAILED`（v0.2 删除 `SCRAPING`）。
- 每 job 内部 task_id set 隔离，跨 job 信号必须 disconnect。
- 工作线程不能直接动 GUI 部件，所有交互走 `OrchestratorSignals`。
- `task_finished` 永远在 `finally` 块内 emit，保证失败也通知 progress_dock。

### 6.4 HTTP 与限速（`app/core/http.py` / `ratelimit.py`）

- 所有 HTTP 走 `default_client(source=...)` 工厂，不允许各 plugin 自行 `httpx.Client(...)`。
- SEC 自动注入 `settings.sec.user_agent`（含用户邮箱，SEC 政策强制）。
- 每数据源独立 TokenBucket（`config.toml` `[rate_limits]`）：cninfo / hkexnews / sec / dart / akshare / eastmoney。
- `verify=True` 永远不能改（不禁用 SSL 校验）。
- `get_json` / `post_json` / `stream_to_file` 全部 tenacity 装饰，3 次指数退避。
- `stream_to_file` 写入 `{dest}.part`，支持 `Range` 断点续传；服务器忽略 Range 返回 200 时从 0 重写，416 时删除 stale `.part` 后让 tenacity 重试，瞬断失败时保留 `.part` 供下次恢复。
- 失败响应（4xx）不写入 cache。

### 6.5 输出

- `safe_filename` 输入路径必经过滤（防注入）。
- 输出目录扁平，无公司/市场嵌套（`tests/test_output_layout.py` 锁定）。
- `output/` 下**只产出 PDF**，sidecar 元数据落在 `data/cache/sidecars/`，不写入任何 Excel / JSON 衍生产物。
- 文件名规则不可变（`{exchange}_{code}_{name}_{year}_{kind_zh}.pdf`）。

### 6.6 GUI

- `MainWindow` 是 `FramelessWindowHint + WA_TranslucentBackground`；圆角通过 `#WindowRoot` QSS `border-radius: 14px` + `QGraphicsDropShadowEffect`。
- 边缘 6px hit-test 实现 8 方向缩放；最大化时去掉圆角（避免裁切）。
- 全部样式集中在 `app/ui/styles/app.qss`，不能在 Python 代码里 setStyleSheet。
- 顶层 UI 文案集中在 `app/ui/strings.py`；运行时语言切换走 `LanguageManager`，保留 widget 状态并原地 `setText()`。
- 状态徽章 `StatusPill` 通过 `dynamic property + style().unpolish/polish` 切换，不重建 widget。

### 6.7 PyInstaller 兼容

- `app/main.py:8-25` 必须保留 `sys.stdout` / `sys.stderr` None 守护（windowed 模式两次踩坑教训）。
- `filings_atlas.spec` 的 `collect_all("akshare")` / `collect_all("OpenDartReader")` / `collect_data_files("pykrx")` 不能去掉。
- `excludes` 列表只能加不能减（已剔除 matplotlib/scipy/torch/QtWebEngine）。
- 冻结模式所有路径走 `Path(sys.executable).parent`（`settings.py::project_root`），不能用 `__file__`。
- `app/core/pdf_renderer.py` 持有线程本地 Playwright browser/context pool：Playwright Sync API 不跨线程共享，每次渲染结束后清理 cookies/permissions；`shutdown_renderer()` 随 `QApplication.aboutToQuit` 做 best-effort 关闭。

## 7. 关键文件路径

```text
development/
├── DEVELOPMENT_BRIEF.md     给 Codex 的代码审核委托书
├── pyproject.toml           依赖与打包元数据
├── requirements.txt         pip 安装列表
├── filings_atlas.spec       PyInstaller spec（one-folder）
├── run.bat                  源码启动（开发）
├── build.bat                一键打包 EXE
│
├── app/
│   ├── main.py              QApplication 入口 + 顶层异常网 + crash.log
│   ├── core/                与 GUI / 插件解耦的核心逻辑
│   │   ├── models.py
│   │   ├── settings.py
│   │   ├── http.py
│   │   ├── ratelimit.py
│   │   ├── cache.py
│   │   ├── output_paths.py
│   │   ├── pdf_renderer.py
│   │   ├── job.py
│   │   └── orchestrator.py
│   ├── ui/                  全部 PySide6 部件（10 文件 + styles/）
│   └── assets/              filings_atlas.ico / filings_atlas_logo.png
│
├── plugins/
│   ├── base.py              ExchangePlugin ABC（3 方法）
│   ├── __init__.py          get_plugin(exchange) 路由
│   └── {ashare,hk,us,kr,tw,jp,uk}/
│       ├── name_resolver.py
│       └── reports.py
│
└── tests/
    ├── test_output_layout.py        PDF 文件名规则（扁平 + safe_filename）
    └── test_selection_logic.py      US 10-K/10-Q + KR 季报选片
```

`config.toml` 在源码模式下落在 `development/` 下；冻结模式下落在 `Path(sys.executable).parent`。

## 8. 数据源声明

| 数据范围 | 数据源 | 入口 | 鉴权 |
|---|---|---|---|
| A 股名称 | akshare `stock_info_a_code_name` | `plugins/ashare/name_resolver.py` | 不需要 |
| A 股 orgId | cninfo `topSearch/detailOfQuery` POST | 同上 | 不需要 |
| A 股报告 | cninfo `hisAnnouncement/query` POST → `static.cninfo.com.cn` PDF | `plugins/ashare/reports.py` | 不需要 |
| 港股名称 | 东方财富 `push2delay.eastmoney.com/api/qt/stock/get` | `plugins/hk/name_resolver.py` | 不需要 |
| 港股报告 | HKEXnews `www1.hkexnews.hk/search/titlesearch.xhtml` HTML | `plugins/hk/reports.py` | 不需要 |
| 美股名称 | SEC `www.sec.gov/files/company_tickers.json` | `plugins/us/name_resolver.py` | 不需要 |
| 美股报告 | SEC `data.sec.gov/submissions/CIK{cik}.json` + 分页 files | `plugins/us/reports.py` | 不需要 |
| 韩股名称 | OpenDartReader `corp_codes`；无 Key 时 DART 公网 `detailSearch.ax` | `plugins/kr/name_resolver.py` / `plugins/kr/dart_web.py` | Key 可选 |
| 韩股报告 | DART `list`；无 Key 时 DART 公网 HTML 搜索 + 公告 PDF/渲染下载 | `plugins/kr/reports.py` / `plugins/kr/dart_web.py` | Key 可选 |
| 台股名称 | TWSE ISIN `isin.twse.com.tw/isin/C_public.jsp?strMode=2/4` 解析 HTML | `plugins/tw/name_resolver.py` | 不需要 |
| 台股报告 | MOPS `doc.twse.com.tw/server-java/t57sb01` 两步下载（POST step=9 → GET 临时 PDF URL） | `plugins/tw/reports.py` | 不需要 |

HTTP 请求原则：

- 只使用公开 GET / POST 请求。
- 不绕过登录或权限控制。
- 不抓取私人数据。
- SEC 请求带最低限速（TokenBucket sec=8 rps）+ 邮箱 UA。
- DART Key 缺失时 KR 自动走公网 fallback；填 Key 时优先走 OpenAPI。
- 用户输入的 ticker code 全程 normalize（strip+upper），不直接拼 URL（防 SSRF）。
- **TWSE 例外**：MOPS 服务端证书缺失 Subject Key Identifier 扩展（Python 3.12+ 严格 OpenSSL 拒绝），`default_client` 仅对 `source=twse` 设 `verify=False`，其他六市场保持 certifi 严格校验。这是 TWSE 服务端的证书 hygiene 问题，非我们可控。

| 层 | 实现 | TTL / 策略 |
|---|---|---|
| name_resolver 全表 | diskcache | 24 小时（A 股代码表 / SEC ticker 表）/ 30 天（TWSE ISIN 上市+上柜） |
| name_resolver 单股 | diskcache | 30 天（cninfo orgId）/ 永久（KR `kr:corp:{code}` corp_code） |
| HTTP 层 | tenacity 指数退避 | 重试 3 次，base 1s，cap 30s |
| 限速 | TokenBucket（同步+异步） | 每数据源独立桶，参数从 config.toml 读 |
| 4xx 响应 | — | 不缓存，直接抛 |

`diskcache` 单例（`app/core/cache.py`）随 `QApplication.aboutToQuit` 调用 `close_cache()` 关闭。名称解析大表或按票 corp_code 使用 `cached_or_load()` 双重检查锁，避免 QThreadPool 并发 cache miss 时重复拉取同一份远端数据。

## 10. 打包与部署

- PyInstaller **one-folder 模式**（一文件模式启动慢 5-10×）。
- 大头依赖：PySide6 ~140 MB / akshare 全家桶 ~120 MB / OpenDartReader+pykrx ~50 MB。
- `excludes`：matplotlib / scipy / torch / QtWebEngine / tkinter / IPython（节省 ~200 MB）。
- 产物：`dist/Filings Atlas/` ~340 MB；入口 `Filings Atlas.exe` ~27 MB（启动器）+ `_internal/` ~314 MB。
- 冻结模式路径：`config.toml` / `output/` / `cache/` / `logs/` 均相对 `Path(sys.executable).parent`。
- 已修复 windowed 模式两个坑：
  - `sys.stderr is None` → loguru `logger.add(...)` `TypeError`（`app/main.py:16-25` 守护）
  - `sys.stdout is None` → akshare 内部 tqdm `.write()` crash（`app/main.py:8-14` 重定向 `os.devnull`）

发布版目录布局（项目根直接 ship）：

```text
Filings Atlas.exe              启动器，27 MB，双击即用
_internal/                     PyInstaller bundled，~314 MB，请勿删除
config.toml                    本机配置（隐藏）
config.example.toml            模板（隐藏）
output/                        抓取结果（PDF 平铺，无 .xlsx）
cache/                         diskcache 持久化
logs/                          loguru 日志
data/                          预留运行时数据
docs_cache/                    预留文档缓存
development/                   源码 + 构建脚本（隐藏，开发用）
```

## 11. 如何加新市场（v1.0 起 8 市场，扩展模板）

新增市场时只扩展“官方披露 PDF 下载”这一条链路，不引入财务数据抽取、指标计算或 Excel 输出。推荐先复制 SG/UK 的结构做最小闭环，再逐步补充 IPO、季度和边界场景。

### 11.1 Exchange 枚举

在 `development/app/core/models.py::Exchange` 增加新成员，并同步 `display_name`：

```python
class Exchange(str, Enum):
    UK = "UK"
    SG = "SG"
    NEW = "NEW"

    @property
    def display_name(self) -> str:
        return {
            "UK": "英股",
            "SG": "新股",
            "NEW": "新市场",
        }[self.value]
```

枚举值会进入输出文件名、sidecar exchange 字段、增量更新匹配和进度 UI。不要改已有枚举值，否则会破坏历史 sidecar 与已下载文件识别。

### 11.2 Plugin 三件套

每个市场至少包含：

```text
development/plugins/{mk}/
├── __init__.py          # {Market}Share(ExchangePlugin)，懒 import 下游模块
├── name_resolver.py    # 用户输入 code -> Ticker(name + external_id)
└── reports.py          # Period -> ReportFile[]，只落地 PDF
```

`__init__.py` 只做契约适配，保持和 SG 插件类似：

```python
class SGShare(ExchangePlugin):
    exchange = Exchange.SG

    def resolve_name(self, code: str) -> Ticker:
        from .name_resolver import resolve as _resolve
        return _resolve(code)

    def fetch_company(self, ticker: Ticker) -> Company:
        from .name_resolver import fetch_company as _fetch_company
        return _fetch_company(ticker)

    def download_reports(self, ticker: Ticker, period: Period, output_root: Path) -> list[ReportFile]:
        from .reports import download as _download
        return _download(ticker, period, output_root)
```

如果数据源有 Key 与无 Key 两条路，拆成 `{mk}_api.py` 与 `{mk}_web.py`，在 `name_resolver.py` / `reports.py` 里统一选择分支。JP 是 Key/API 双模式模板：

- `edinet_api.py`：EDINET API v2，使用 `Subscription-Key`，走 `default_client(source="edinet")` 与 `rate_limits.edinet`。
- `edinet_web.py`：公网 fallback 入口，选择器集中在 `_SELECTORS`。
- `name_resolver.py`：规范化 `7203` / `7203.T` / `JP7203`，缓存 EDINET code 与 filer name。
- `reports.py`：按 `PeriodType` 映射 EDINET `docTypeCode`，下载或转换为 PDF，返回 `ReportFile`。

SG 是无 Key、公网 JSON + HTML 附件二跳模板：

- `sgxnet_web.py`：读取 SGX config/CMS token，走 `default_client(source="sgxnet")` 与 `rate_limits.sgxnet`。
- `name_resolver.py`：规范化 `D05` / `D05.SI`，公告搜索失败时可回退 IPO 列表 id。
- `reports.py`：年报走 `financialreports/v1.0`，半年报走 `announcements/v1.1/securitycode`，IPO 走 `ipoprospectus/v1.0/`；API `url` 先到 HTML，再解析真实 PDF。

v1.0 的 Key 口径：

- **Japan / EDINET**：强烈推荐配置 EDINET Subscription-Key。官方 API v2 的 documents list 与 document download 都要求 `Subscription-Key`；用户可免费注册 `disclosure2.edinet-fsa.go.jp` / `api.edinet-fsa.go.jp` 后在设置里填写。无 Key 时只尝试公网 fallback，稳定性不等同于 API。
- **United Kingdom / FCA NSM**：不需要 Key。UK 插件走 FCA NSM 公网 search endpoint 与 `data.fca.org.uk/artefacts/` 文件下载/HTML 渲染。
- **Singapore / SGXNet**：不需要 Key。SG 插件用公开 config + CMS token 访问 SGXNet API，PDF 下载仍经 `stream_to_file`；SG 页面本身不作为 Playwright 依赖。

### 11.3 注册

在 `development/plugins/__init__.py::get_plugin` 加显式分支。当前项目不使用自动发现，便于 PyInstaller hidden import 与 Reviewer 审查。

```python
if exchange is Exchange.SG:
    from .sg import SGShare
    return SGShare()
```

新增市场后同步检查 `development/filings_atlas.spec` 的 `hiddenimports`，确保 frozen EXE 能加载新插件。

### 11.4 UI 接入

最小 UI 接线清单：

- `app/ui/main_view.py`：exchange tuple 加新 `Exchange`。
- `app/ui/exchange_selector.py`：chip 名称、meta 与顺序。
- `app/ui/exchange_panel.py`：panel title/subtitle。
- `app/ui/ticker_row.py`：输入 placeholder。
- `app/ui/batch_import_dialog.py`：批量导入正则、normalize 和示例。
- `app/ui/progress_dock.py`：市场名称映射。
- `app/ui/styles/palette.py`：新增 accent 色。
- `app/ui/strings.py`：至少补 `ES_NAME_*`、`ES_META_*`、`EP_TITLE_*`、`EP_SUBTITLE_*`、`TR_PLACEHOLDER_*` 的 zh/en 双语 key。

所有用户可见文字必须进 `strings.py`，并保证 `STRINGS["zh"]` 和 `STRINGS["en"]` key 完全一致，英文 value 不含 CJK。

### 11.5 Settings 可选 Key

只有数据源确实需要或显著受益于 Key 时才加设置项。结构参照 EDINET：

```python
class EDINETCfg(BaseModel):
    api_key: str = ""

class Settings(BaseModel):
    edinet: EDINETCfg = Field(default_factory=EDINETCfg)
```

如果 Key 被缓存到客户端或 resolver，必须提供 cache invalidation hook，并在 `settings_dialog.py` 保存 Key 变化后调用。EDINET 当前用 `invalidate_edinet_client_cache()` 清理 `plugins.jp.name_resolver` 里的缓存。

不要为 UK/NSM 这类免 Key 数据源加空设置项，避免误导用户。

### 11.6 测试要求

新增市场至少补两组测试：

- `tests/test_{mk}_name_resolver.py`：代码 normalize、resolve/fetch_company、Key 分支或公网搜索 shape。
- `tests/test_{mk}_reports.py`：list_filings、select_filing、PDF 下载/HTML 渲染、无披露返回 `[]`。

最低要求不少于 5 个用例，并同步：

- `tests/integration/test_plugins.py`：`get_plugin(Exchange.X)` 注册测试。
- `tests/test_batch_import.py`：批量导入 normalize 测试。
- `tests/test_ui_strings.py` / `tests/test_language_switch.py`：双语 key 和 UI 扫描保持绿。

最终验证矩阵：

```powershell
cd development
python -m pytest -m "not e2e" -v -p no:cacheprovider
python -m ruff check . --no-cache
```

如果数据源允许公网访问，还应做 1-3 票真实 smoke，并确认落地文件以 `%PDF` 开头。若需要 Key 且本机未配置，必须在自检报告中明确说明阻断原因。

## 12. 性能优化注意事项

- 并发数和限流必须分开调：`concurrency.max_workers` 控制本地 worker 数，`settings.rate_limits.*` 控制单 source 请求速率。发现 429/403 时优先单独降低对应 source rate，不要全局回滚并发。
- `stream_to_file` 默认 chunk size 为 256 KB。大 PDF 下载吞吐更好，但 UI 进度信号粒度会比 64 KB 更粗；如未来加实时进度条，需要重新评估这个 tradeoff。
- Playwright Sync API 对线程敏感，不能把同一个 browser/context 跨 QThreadPool worker 共享。当前实现是线程本地 browser/context pool，并在每次渲染后执行 `clear_cookies()` / `clear_permissions()`，避免 HTML 渲染任务之间泄漏登录态或站点权限。
- name resolver 预热只适合 A 股、台股、美股这类全量映射表；HK/KR/JP/UK/SG 等单票远端查询型 resolver 不应在启动期批量预热。
- sidecar 写入用同目录临时文件再 `replace()`，避免下载任务中断时留下半截 JSON。

## 13. 验证矩阵

当前测试集（`development/tests/`）：

```bat
cd development
python -m pytest -m "not e2e" -v
python -m pytest tests/e2e/test_smoke.py::test_kr_005930_2024 -v
python -m ruff check . --no-cache
```

覆盖：

- `test_output_layout`：PDF 报告路径扁平化 + `safe_filename` 行为。
- `test_selection_logic`：US 非日历财年 10-Q 季度归类；KR Q1/Q3 按 `rcept_dt` 月份选片。
- `test_cached_or_load`：name_resolver 缓存 single-flight、防 None 污染。
- `test_stream_resume`：`.part` 断点续传、Range 忽略、416 清理、失败保留。
- `test_playwright_pool`：Playwright 线程本地 context 复用、cookie/permission 清理与幂等关闭。
- `test_ui_strings`：顶层 UI 模块无 CJK 字符串字面量、strings 常量命名唯一。

v1.0 本地验收：非 e2e `170 passed / 22 deselected`，benchmark opt-in `11 passed`，ruff 全绿。

端到端验证（手动）：

| 市场 | Ticker | 名称 | 报告 PDF |
|---|---|---|---|
| A 股 | 600519 | ✓ 贵州茅台 | ✓ annual_report.pdf 3.5 MB |
| 港股 | 00700 | ✓ 腾讯控股 | ✓ annual_report.pdf 7.3 MB |
| 美股 | AAPL | ✓ Apple Inc. | ✓ 10-K FY2023 + FY2024 (1.5 MB each) |
| 韩股 | 005930 | ✓ 三星电子 | ✓ annual_report.pdf（OpenAPI + 无 Key 公网 fallback） |
| 台股 | 2330 | ✓ 台積電 | ✓ annual_report 9.99 MB / 半年报 7.2 MB / 60+ IPO 公开说明书 |
| 日股 | 7203 | ✓ Toyota | EDINET API Key 推荐；无 Key 环境不列入 release smoke |
| 英股 | ULVR / HSBA / AZN | ✓ | ✓ FCA NSM PDF/HTML 渲染路径 |
| 新加坡 | D05 / U11 / Z74 / 3407 | ✓ | ✓ 年报 / 半年报 / IPO 招股书 |

未覆盖的回归路径（v0.1→v0.5 规划已归档于 `roadmap/archive/`），重点：

- KR 公网 fallback 已跑通；后续重点是 DART HTML 改版时维护 `plugins/kr/dart_web.py::_SELECTORS`。
- US `submissions.files[]` 分页兜底（老 ticker，AAPL FY2018 之前会落到这条路径）。
- HK 选片只匹配年份字符串，同年份多份补充公告 / ESG 报告可能误选（Sprint 03 解决）。
- KR `_download_filing` 用 `os.chdir`，QThreadPool 4 worker 并发不安全（Sprint 02 解决）。

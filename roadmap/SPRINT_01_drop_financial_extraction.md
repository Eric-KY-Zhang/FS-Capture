# Sprint 01 — 工具重定位：移除抓数据功能

> **目标**：把 FS Capture 从「PDF + 三大报表数字 + Excel 底稿」混合工具简化为**纯 PDF 抓取工具**。删除所有 `fetch_financials` 调用链路、`FinancialStatement` 模型、`excel_writer` 输出层及相关测试。
>
> **预估工作量**：5-7 h
> **依赖**：无
> **风险**：中（涉及 model + ABC + 4 plugin + orchestrator + UI + 测试，影响面广但每处改动相对机械）

## 1. 背景

工具原 v0.1 设计是「下 PDF + 抓三大报表数字 + 写 8-sheet Excel 底稿」三合一。用户重新定位为**纯 PDF 抓取**——只解决"批量拿到原始 PDF"这个问题，三大报表数字与底稿装填都不做。

理由：

- 数据准确性兜底很难（akshare 英文 EM key / SEC GAAP / 港股中文混编，跨市场字段对齐困难）
- 用户瑞华底稿是中文行，但要保完整匹配需要持续维护映射表
- VBA Captor v1.0 已经覆盖了财务数字抓取场景（专做 Excel-only 的快速路径）
- FS Capture 的差异化价值是「下载 PDF」，不是再做一个数字工具

本 sprint 把这部分代码与文档全部移除，让代码和定位一致。

## 2. 涉及范围

### 待删除文件

```text
development/plugins/ashare/financials.py
development/plugins/hk/financials.py
development/plugins/us/financials.py
development/plugins/kr/financials.py
development/app/exporters/excel_writer.py
development/app/exporters/templates/sina_v22_schema.json
```

`development/app/exporters/templates/.gitkeep` 保留（templates 目录留空，将来若要做 PDF sidecar 模板可复用）。

`development/app/exporters/__init__.py` 清空成 placeholder（或整个目录 + `__init__.py` 一并删掉，看 Codex 判断是否需要保留 namespace）。

### 待修改文件

```text
development/plugins/base.py                  # ExchangePlugin ABC 删除 fetch_financials 方法
development/app/core/models.py               # 删除 FinancialStatement、StatementType
development/app/core/job.py                  # TaskResult 删除 statements 字段；TaskStatus 删除 SCRAPING
development/app/core/orchestrator.py         # _TaskRunnable.run 删除 SCRAPING 阶段调用
development/plugins/ashare/__init__.py       # 去掉对 financials 的引用（如有）
development/plugins/hk/__init__.py           # 同上
development/plugins/us/__init__.py           # 同上
development/plugins/kr/__init__.py           # 同上
development/app/ui/main_view.py              # 去掉 excel_writer.write_workbook 调用 + 完成提示文案改"下载完成 N 个 PDF"
development/app/ui/progress_dock.py          # 去掉 SCRAPING 状态文案
development/tests/test_output_layout.py      # 删除 test_workbook_uses_visible_wide_sheets_without_ruihua_tab；保留 test_report_path_is_flat_under_output_root
development/tests/test_selection_logic.py    # 删除 HKFinancialParsingTests 与 USFinancialSelectionTests 整两个类；保留 USReportSelectionTests 与 KRReportSelectionTests
development/DEVELOPMENT_BRIEF.md             # 已由 Planner 同步更新（本 sprint 不再动）
```

### 严禁动

- `development/plugins/base.py` 的 `resolve_name` / `fetch_company` / `download_reports` 三个方法签名（保留）
- `development/app/core/models.py` 的 `Exchange` / `PeriodType` / `Period` / `Ticker` / `Company` / `ReportFile`（保留）
- `development/app/core/{http,cache,ratelimit,settings,output_paths}.py`（HTTP / 限速 / cache / 配置 / 文件名规则）
- `development/app/main.py:8-25`（PyInstaller stdio None 守护）
- `development/app/ui/main_window.py`（无边框窗口/标题栏/拖拽）+ 其他 UI 视觉文件
- `development/app/ui/styles/`（QSS 样式）
- `development/fs_capture.spec`（PyInstaller spec）
- 所有 `plugins/{ashare,hk,us,kr}/{name_resolver,reports}.py`（PDF 抓取主路径，本 sprint 不动）

## 3. 实施步骤

### Step 1 — 删除 4 个 financials.py 与 excel_writer 相关文件

```bat
del development\plugins\ashare\financials.py
del development\plugins\hk\financials.py
del development\plugins\us\financials.py
del development\plugins\kr\financials.py
del development\app\exporters\excel_writer.py
del development\app\exporters\templates\sina_v22_schema.json
```

`development/app/exporters/__init__.py` 内容改为空（保 namespace 给将来 PDF sidecar）：

```python
# Reserved for future per-PDF sidecar metadata writers.
```

### Step 2 — 修改 `plugins/base.py`：精简 ABC

删除 `fetch_financials` 方法（连同 docstring 与 `FinancialStatement` import）。修改后：

```python
from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from app.core.models import Company, Exchange, Period, ReportFile, Ticker


class ExchangePlugin(ABC):
    """Per-market data adapter. Subclass once per exchange.

    The orchestrator calls these in sequence per (ticker, period) task:
      1. resolve_name(code)        -> Ticker (with name + external_id)
      2. fetch_company(ticker)     -> Company metadata (industry, currency, ...)
      3. download_reports(t, p)    -> ReportFile[] streamed to disk
    """

    exchange: Exchange

    @abstractmethod
    def resolve_name(self, code: str) -> Ticker: ...

    @abstractmethod
    def fetch_company(self, ticker: Ticker) -> Company: ...

    @abstractmethod
    def download_reports(
        self, ticker: Ticker, period: Period, output_root: Path
    ) -> list[ReportFile]: ...
```

### Step 3 — 修改 `app/core/models.py`：删除财务数据 model

删除：
- `StatementType` 枚举（整个）
- `FinancialStatement` 类（整个）

保留：
- `Exchange` / `PeriodType` / `Period` / `Ticker` / `Company` / `ReportFile`

### Step 4 — 修改 `app/core/job.py`：精简 TaskResult / TaskStatus

`TaskStatus` 删除 `SCRAPING` 枚举值（保留 `PENDING` / `RESOLVING` / `DOWNLOADING` / `DONE` / `FAILED` / `CANCELLED`）。

`TaskResult` 删除 `statements: list[FinancialStatement]` 字段（如有）。保留 `reports: list[ReportFile]`。

### Step 5 — 修改 `app/core/orchestrator.py`：精简任务流程

`_TaskRunnable.run` 删除 SCRAPING 阶段调用：

```python
# Before
r.status = TaskStatus.SCRAPING
self.signals.task_progress.emit(r, f"抓取{r.period.label()}财务数据")
statements = plugin.fetch_financials(r.ticker, r.period)
r.statements = statements
r.status = TaskStatus.DONE

# After
r.status = TaskStatus.DONE
```

### Step 6 — 修改 4 个 plugin 包

`plugins/ashare/__init__.py`、`plugins/hk/__init__.py`、`plugins/us/__init__.py`、`plugins/kr/__init__.py` 内若有 `from .financials import ...` 之类引用，全部删除。这 4 个文件应该是 boilerplate，主要 import + 实例化。

每个 plugin 的实现类（`AsharePlugin` / `HKPlugin` / `USPlugin` / `KRPlugin`）若实现了 `fetch_financials`，把整个方法删掉。

### Step 7 — 修改 UI

**`app/ui/main_view.py`**：

- 删除 `from app.exporters.excel_writer import write_workbook` 等 import
- 删除 `_on_job_finished` 回调里调 `write_workbook(...)` 的整段
- 调整完成提示，改成 "已下载 N 个 PDF，输出在 {output_dir}"（具体文案 Codex 自定）
- 不再向用户提示 .xlsx 路径

**`app/ui/progress_dock.py`**：

- `_STATUS_LABEL` 字典（如存在）删除 `SCRAPING` 行
- 去掉任何"抓取财务数据"文案

### Step 8 — 修改测试

**`tests/test_output_layout.py`**：

保留 `test_report_path_is_flat_under_output_root`（PDF 文件名规则）。

删除 `test_workbook_uses_visible_wide_sheets_without_ruihua_tab` 整个测试方法（连同相关 import：`load_workbook`、`write_workbook`、`FinancialStatement`、`StatementType` 等）。

**`tests/test_selection_logic.py`**：

保留 `USReportSelectionTests`（10-K / 10-Q 选片）与 `KRReportSelectionTests`（DART 季报选片）。

删除 `HKFinancialParsingTests`（HK long-table pivot 是财报 normalize 逻辑）与 `USFinancialSelectionTests`（companyfacts 选片是财报取数逻辑）整两个类，连同它们的 import：`hk_row_to_lines`、`select_hk_period_row`、`select_us_period_value`。

文件末尾 import 段简化后只保留：

```python
from plugins.kr.reports import _select_filing as select_kr_filing
from plugins.us.reports import _filter_table as filter_us_table
```

### Step 9 — 验证 Job 完整链路

Codex 自检：在删完后跑测试：

```bat
cd development
python -m unittest tests.test_output_layout tests.test_selection_logic
```

预期全 pass。

如能联网跑端到端：

```bat
cd development
python -m app.main
```

操作：

1. 勾选 A 股、输入 600519，点确认
2. 期间选 2024 年报，点开始抓取
3. 期望：
   - progress_dock 不再显示"抓取财务数据"步骤
   - 完成后弹"已下载 N 个 PDF"
   - `output/` 下平铺得到 `A_600519_2024_annual_annual_report.pdf` 等 PDF 文件
   - **不应有任何 .xlsx 文件产出**

## 4. 验证

### 4.1 单元测试（Codex 必跑）

```bat
cd development
python -m unittest tests.test_output_layout tests.test_selection_logic
```

预期：保留的 4 个测试方法全 pass（`test_report_path_is_flat_under_output_root` + `USReportSelectionTests.test_non_calendar_fiscal_quarters_are_mapped_by_annual_cycle` + `KRReportSelectionTests.test_quarterly_reports_choose_q1_and_q3_by_receipt_month`）。

### 4.2 端到端冒烟（Reviewer 必跑）

```bat
cd development
python -m app.main
```

操作：

1. 勾选 A 股 + 港股 + 美股
2. 输入 600519 / 00700 / AAPL，点确认
3. 期间选 2024 年报
4. 点开始抓取
5. 完成后检查：
   - **`output/` 下只有 PDF，没有 .xlsx**
   - PDF 文件名符合 `{exchange}_{code}_{year}_{period_type}_{kind}.pdf` 规则
   - progress_dock 在每个 task 上的状态文本不出现"抓取财务数据"

### 4.3 frozen 回归（Reviewer 必跑）

- EXE 启动不退化（无新崩溃）
- `safe_filename` 行为不变（`tests.test_output_layout::test_report_path_is_flat_under_output_root` pass）
- Plugin ABC 签名简化但保留 3 方法（`from plugins import get_plugin` 不报错）
- HTTP / cache / ratelimit / settings 模块未被改动（`git diff` 0 行）
- `app/main.py:8-25` 的 stdio 守护未动

### 4.4 grep 兜底检查（Reviewer 跑）

```bat
cd development
findstr /S /M "FinancialStatement" *.py
findstr /S /M "fetch_financials" *.py
findstr /S /M "StatementType" *.py
findstr /S /M "write_workbook" *.py
findstr /S /M "_LOOKUP_ALIASES" *.py
findstr /S /M "_LINE_LABELS" *.py
```

期望全部 0 命中（除注释 / 文档字符串提到历史背景的情况，Codex 应当一并清理）。

## 5. 联系 Planner 触发条件

Codex 在以下情况下停下，把状态写到 `current.md` 等 Planner 决策：

1. 任一 frozen 测试退化（`test_report_path_is_flat_under_output_root` 不通过）
2. UI 改动后双击启动崩溃，且短时间无法定位
3. 发现某个 plugin 的 `__init__.py` 在实例化时还引用 `fetch_financials`，但删除后影响其他模块
4. `app/exporters/__init__.py` 删除后有外部 import 残留（如某 UI 文件 `from app.exporters import ...`）
5. 端到端冒烟（如能联网）下载 PDF 链路退化（HK 选片 / US submissions / cninfo 任一失败）

## 6. Codex 完成回报格式

完成后在 `current.md` 末尾「## Codex 回报」段写：

```markdown
## Codex 回报

**commit**: <hash>

**删除文件**:
- plugins/ashare/financials.py (LOC -130)
- plugins/hk/financials.py (LOC -100)
- plugins/us/financials.py (LOC -120)
- plugins/kr/financials.py (LOC -80)
- app/exporters/excel_writer.py (LOC -365)
- app/exporters/templates/sina_v22_schema.json (LOC -X)

**修改文件**（行数 delta）:
- plugins/base.py (-15)
- app/core/models.py (-30)
- app/core/job.py (-2)
- app/core/orchestrator.py (-8)
- plugins/{ashare,hk,us,kr}/__init__.py (-N each)
- app/ui/main_view.py (-12)
- app/ui/progress_dock.py (-3)
- tests/test_output_layout.py (-50)
- tests/test_selection_logic.py (-60)

**清理后总 LOC**:
- Python: v0.1 ~3,900 → v0.2 ~3,000 行（约 -900 行）

**单元测试结果**: 3 passed, 0 failed
**端到端冒烟**: 已跑 / 沙箱无网络，请 Reviewer 跑 / 跑了但 X 市场失败
**grep 兜底**: 全部 0 命中 / 残留 X 处（说明位置）

**意外问题 / 决策点**:
- (空 / 列出 1-3 项)
```

## 7. Reviewer 验收清单

- [ ] 4 个 `financials.py` 文件不再存在
- [ ] `excel_writer.py` 与 `sina_v22_schema.json` 不再存在
- [ ] `plugins/base.py::ExchangePlugin` 只剩 3 个 abstract method
- [ ] `app/core/models.py` 不含 `FinancialStatement` / `StatementType`
- [ ] `app/core/job.py::TaskStatus` 不含 `SCRAPING`
- [ ] `python -m unittest tests` 保留的测试全 pass
- [ ] 端到端 600519 / 00700 / AAPL 跑完，`output/` 只有 PDF 没有 .xlsx
- [ ] `findstr /S "FinancialStatement"` 0 命中（除有意保留的历史注释）
- [ ] PyInstaller 启动无退化（双击 EXE 窗口正常打开）
- [ ] HTTP / cache / ratelimit / settings / output_paths 未被改动

验收通过 → 在 `current.md` 末尾「## Reviewer 验收」段写 `✅ Sprint 01 通过` + Sprint 02 起号。

验收不通过 → 在 `current.md` 末尾「## Reviewer 验收」段写 follow-up 任务清单，Codex 修复后再次验收。

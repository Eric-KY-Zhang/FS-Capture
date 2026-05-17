# Codex Bundle Run — Sprint 01~05 端到端执行 Spec

> **任务模式**：一次性按顺序跑完 Sprint 01 → 02 → 03 → 04 → 05，每 sprint 独立 commit，所有 sprint 完成后整体交付给 Reviewer。
>
> **协作链路**：Planner（Claude Code）↔ `current.md` ↔ Codex（你）。详细 plan 文件在 `roadmap/`。
>
> **当前定位**：FS Capture v0.2 = 纯 PDF 抓取工具（年报 / 审计报告 / 季报 / 半年报 / IPO 招股书），不抓三大报表数字、不生成 Excel 底稿。

## 0. 你是谁，做什么

你是 Codex generator。本次任务是端到端实施 FS Capture 项目的 5 个 sprint 优化路线图，把工具从 v0.1（混合工具）推进到 v0.5（生产可用的纯 PDF 抓取工具）。Planner 不做中间 review，等你 5 个 sprint 全跑完后再统一验收。

---

## 1. 项目 anchor

- **项目根**：`E:\Claude+CODEX Project\FS Capture`
- **源码**：`development/`（PySide6 GUI + 4 市场 plugin + httpx + diskcache + tenacity）
- **关键文档**：
  - [`ARCHITECTURE.md`](../ARCHITECTURE.md) — 架构 + invariants
  - [`PROJECT_RETROSPECTIVE.md`](../PROJECT_RETROSPECTIVE.md) — v0.1→v0.2 重定位背景
  - [`development/DEVELOPMENT_BRIEF.md`](../development/DEVELOPMENT_BRIEF.md) — 实现细节 + 已知问题
  - [`roadmap/ROADMAP.md`](ROADMAP.md) — 5 sprint 整体路线图
  - [`roadmap/SPRINT_01_drop_financial_extraction.md`](SPRINT_01_drop_financial_extraction.md) — Sprint 01 详细计划

**首要原则**：本工具**只做 PDF 批量下载**。任何 sprint 不能加回数据抓取 / Excel 输出 / 财务指标计算 / 跨市场对标功能。

---

## 2. 必读文档顺序

开始执行前**必须**按顺序阅读：

1. `roadmap/ROADMAP.md`（整体 sprint 路线 + Out of Scope）
2. `ARCHITECTURE.md`（架构 + 6 大类 Invariants）
3. `development/DEVELOPMENT_BRIEF.md`（§1 项目背景、§4 架构模式、§6 实现细节、§8 已知问题、§10 已修复风险）
4. `roadmap/SPRINT_01_drop_financial_extraction.md`（Sprint 01 完整 spec）
5. 本文件 §6-§9（Sprint 02-05 inline spec）

---

## 3. 执行规则

### 3.1 顺序与提交

- 严格按 Sprint 01 → 02 → 03 → 04 → 05 顺序执行
- 每个 sprint 完成后 **1 个独立 commit**，commit message 格式：
  ```
  Sprint NN: <主题>
  
  - <要点 1>
  - <要点 2>
  ...
  ```
- Sprint 01 完成 + 所有自检 pass 后才能进 Sprint 02。前 sprint 失败不能进下一 sprint。

### 3.2 每 sprint 自检

每 sprint 收尾时跑：

```bat
cd development
python -m unittest tests.test_output_layout tests.test_selection_logic
```

预期全 pass。Sprint 03 跑 + `tests.test_hk_selection`；Sprint 05 改 pytest 后跑 `pytest tests/`。

如沙箱可联网，跑端到端冒烟：

```bat
python -m app.main
```

操作：勾选 A/HK/US，输入 600519/00700/AAPL，期间选 2024 年报，点开始抓取。期望：`output/` 平铺得到 PDF，无 .xlsx。

### 3.3 回报格式

每 sprint 完成后，在 `current.md` 末尾追加：

```markdown
## Sprint NN 回报

**commit**: <hash>
**主题**: <sprint 名>

**新增文件**:
- path/to/file.py (LOC: X)

**删除文件**:
- path/to/file.py (LOC -X)

**修改文件**（行数 delta）:
- path/to/file.py (+X / -Y)

**单元测试结果**: N passed, M failed
**端到端冒烟**: 已跑 / 沙箱无网络 / 跑了但 X 失败
**意外问题 / 决策点**:
- (空 / 列出 1-3 项)
```

### 3.4 遇阻处理

下列情况立刻停下，在 `current.md` 末尾追加"## Codex 阻塞"段写明状态，**不要继续后续 sprint**：

- 任一 frozen 测试退化（不在该 sprint 计划改动范围内的 unittest 失败）
- PDF 抓取主路径退化（A/HK/US 任一市场冒烟失败）
- 删除文件后有外部 import 残留无法解决
- 实现某个特性需要改 frozen 区（见 §4）
- 第三方库行为与文档不符（如 OpenDartReader.document 参数与预期不同）
- 任一 sprint 的具体 spec 跟 ARCHITECTURE 不变量冲突，且 Codex 无法判断哪个优先

---

## 4. 整体严禁动（跨 5 sprint 都不能动）

以下文件 / 区域在所有 sprint 中**都不能修改**（除非该 sprint 的 plan 明确允许）：

- `app/main.py:8-25`（PyInstaller windowed 模式 stdio None 守护）
- `app/core/http.py`（httpx Client 工厂 + tenacity 装饰器）的核心 API 签名（`get_json` / `post_json` / `stream_to_file`），但 `stream_to_file` 内部超时配置允许 Sprint 02 改
- `app/core/ratelimit.py`（TokenBucket）
- `app/core/output_paths.py`（`safe_filename` + `report_output_path` 文件名规则）
- `fs_capture.spec`（PyInstaller spec）
- `app/ui/main_window.py`（无边框窗口 + 标题栏 + hit-test）— 视觉相关
- `app/ui/styles/`（QSS）— 视觉相关
- `VBA Captor/` 整目录（独立 v1.0 子项目）
- `roadmap/ROADMAP.md` / `ARCHITECTURE.md` / `PROJECT_RETROSPECTIVE.md` / `development/DEVELOPMENT_BRIEF.md` / `README.md`（这些文档由 Planner 维护，Codex 不应修改）

ABC 与数据模型在 Sprint 01 后**冻结**：

- `plugins/base.py::ExchangePlugin`：`resolve_name` / `fetch_company` / `download_reports` 三方法签名，Sprint 02-05 不得增删
- `app/core/models.py`：`Exchange` / `PeriodType` / `Period` / `Ticker` / `Company` / `ReportFile` 字段，Sprint 02-05 只能加新字段不能删（除非 ROADMAP 明确允许，如 Sprint 04 加 `IPO_PROSPECTUS` 到 `PeriodType`）

---

## 5. Sprint 01 — 工具重定位（已有详细 plan）

**详见**：[`roadmap/SPRINT_01_drop_financial_extraction.md`](SPRINT_01_drop_financial_extraction.md)

**简述**：删除 4 个 `financials.py` + `excel_writer.py` + `sina_v22_schema.json`；从 ABC 删除 `fetch_financials`；从 models 删除 `FinancialStatement` / `StatementType`；orchestrator 移除 SCRAPING 阶段；UI 改完成提示文案；测试删两个失效类。

**自检**：
- 单元测试 3 pass（`test_report_path_is_flat_under_output_root` + 2 个 report selection 测试）
- grep 兜底：`findstr /S "FinancialStatement" *.py` 等 6 个 grep 全 0 命中
- 端到端冒烟：`output/` 只产 PDF 不产 .xlsx

**完成后进 Sprint 02。**

---

## 6. Sprint 02 — 健壮性（3-4 h）

### 6.1 子任务

#### A. KR `_download_filing` 移除 `os.chdir`

**当前问题**：`plugins/kr/reports.py` 用 `os.chdir(target_dir)` 临时切到下载目录调 `dart.document(rcept_no)`，下载完切回。但 `os.chdir` 是进程级状态，QThreadPool 4 worker 并发跑 KR task 会互相覆盖 cwd → 下载到错乱目录或 race condition。

**修复方案**（按优先级选）：

1. **首选**：用 `threading.Lock` 在调用 `dart.document` 时全局加锁。代价：KR 下载串行，但实际 KR 数据源也限速 5rps，影响有限。
   ```python
   _DART_DOWNLOAD_LOCK = threading.Lock()
   
   def _download_filing(self, rcept_no: str, target_dir: Path) -> Path | None:
       with _DART_DOWNLOAD_LOCK:
           original_cwd = os.getcwd()
           try:
               os.chdir(target_dir)
               dart.document(rcept_no)
               # find downloaded file
           finally:
               os.chdir(original_cwd)
   ```

2. **次选**（如 OpenDartReader 库源码可改）：monkey-patch `OpenDartReader.document` 接受 `output_dir` 参数。需要验证库版本兼容性。

3. **如发现库内部用 `requests.get()` + `open(filename, 'wb')`**，可以重写整个下载逻辑，绕过库：直接用 httpx 拉 DART filing URL。

**Codex 选哪个方案**：默认走方案 1（最小改动 + 线程安全）；如果发现 OpenDartReader 内部实现允许更优解，写在回报里。

#### B. Cancel 按钮接通

**当前问题**：`app/ui/progress_dock.py` 定义了 `cancel_requested = Signal()` 但没 receiver；`_TaskRunnable.run` 没 cancel token，跑到一半的下载无法中断。

**修复**：

1. `app/core/orchestrator.py::Orchestrator` 加 `_cancel_event: threading.Event`：
   ```python
   def __init__(self, settings, parent=None):
       ...
       self._cancel_event = threading.Event()
   
   def request_cancel(self) -> None:
       self._cancel_event.set()
   
   def submit_job(self, job: Job) -> None:
       self._cancel_event.clear()
       ...
   ```

2. 把 event 传给每个 `_TaskRunnable`：
   ```python
   for r in tasks:
       self.pool.start(_TaskRunnable(r, self.signals, output_root, self._cancel_event))
   ```

3. `_TaskRunnable.run` 在每个 step 开始前 check：
   ```python
   def _check_cancel(self):
       if self._cancel_event.is_set():
           raise _Cancelled()
   
   def run(self):
       try:
           self._check_cancel()
           self.signals.task_started.emit(r)
           ...
           self._check_cancel()
           # resolve_name
           ...
           self._check_cancel()
           # download_reports
           ...
       except _Cancelled:
           r.status = TaskStatus.CANCELLED
       except Exception as exc:
           r.status = TaskStatus.FAILED
           ...
       finally:
           self.signals.task_finished.emit(r)
   ```

4. `app/ui/main_view.py` 把 `progress_dock.cancel_requested` connect 到 `orchestrator.request_cancel`：
   ```python
   self.progress_dock.cancel_requested.connect(self.orchestrator.request_cancel)
   ```

5. 确认 `TaskStatus` 已有 `CANCELLED` 枚举（v0.2 应该保留这个）；progress_dock 状态文案"已取消"。

**注意**：cancel 不能中断 httpx 流式下载到一半的请求（除非用 timeout），但可以防止跨 task 启动新下载——这是合理的折中。

#### C. diskcache close on exit

**当前问题**：`app/core/cache.py::get_cache()` 用 `lru_cache(1)` 单例，从未关闭 → 进程退出时可能留 lock 文件 / 数据未刷盘。

**修复**：

1. `app/core/cache.py` 加：
   ```python
   def close_cache() -> None:
       """Close the cache singleton if it exists. Safe to call multiple times."""
       cache = get_cache.__wrapped__() if get_cache.cache_info().currsize > 0 else None
       # 或者维护一个模块级 _cache_singleton 变量更直接
   ```
   
   更简洁实现：
   ```python
   _cache: Cache | None = None
   
   def get_cache() -> Cache:
       global _cache
       if _cache is None:
           _cache = Cache(...)
       return _cache
   
   def close_cache() -> None:
       global _cache
       if _cache is not None:
           _cache.close()
           _cache = None
   ```

2. `app/main.py`：在 `QApplication` 创建后挂 `aboutToQuit`：
   ```python
   from app.core.cache import close_cache
   ...
   app.aboutToQuit.connect(close_cache)
   ```

#### D. httpx.stream 超时拆分

**当前问题**：`app/core/http.py::stream_to_file` 用全局 `timeout=30s`。港股年报 PDF 7+ MB、慢网下读到一半就 timeout。

**修复**：

1. 改 `stream_to_file` 内部用 `httpx.Timeout`：
   ```python
   from httpx import Timeout
   
   def stream_to_file(client, url, dest, ...):
       timeout = Timeout(connect=30.0, read=None, write=30.0, pool=30.0)
       with client.stream("GET", url, timeout=timeout) as resp:
           ...
   ```

2. `read=None` 表示读不超时（流式下载只要数据在动就不算超时）。`connect=30.0` 仍守住连不上的情况。

3. **不动** `get_json` / `post_json`（API 调用应该有读超时，避免挂死）。

#### E. ticker_row name_label 可选

**当前**：`app/ui/ticker_row.py` 的 `name_label` 是 `QLabel`，鼠标无法选中拷贝。

**修复**：

```python
from PySide6.QtCore import Qt
...
self.name_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
```

确认 cursor hover 时变成 IBeam 光标（QLabel 默认行为，不需额外配）。

### 6.2 验证

- 单元测试：补一个 `tests/test_orchestrator.py`：
  - `test_cancel_event_stops_task`：mock plugin 让 download_reports sleep 1s，submit job，立即 cancel，断言 task status 是 CANCELLED
  - `test_cache_close_idempotent`：调 2 次 close_cache 不报错
- 端到端冒烟：跑 600519，途中点取消，应该能停在当前 task；name_label 双击/拖选能选中
- frozen 回归：`test_output_layout` + `test_selection_logic` 全 pass

### 6.3 严禁动（Sprint 02 专属）

- 不能改 ABC 签名（Sprint 01 后已冻结）
- 不能动 ratelimit / output_paths
- KR 修复方案选 lock 路径就**不要**改 OpenDartReader 库本身
- httpx 改动只在 `stream_to_file`，不能 leak 到其他 HTTP API

---

## 7. Sprint 03 — HK 报告选片精确化（5-7 h）

### 7.1 子任务

#### A. PDF 内嵌文本验证

**目标**：HK `_select_main` 当多于 1 个候选时，对 top-N 候选下载首页+目录页做内嵌文本验证。

**实施**：

1. 加 `plugins/hk/_pdf_verify.py`：
   ```python
   import io
   import httpx
   from pypdf import PdfReader
   from app.core.http import default_client
   
   def verify_pdf_year_and_kind(url: str, target_year: int, target_kind: str) -> bool:
       """Download first 3 pages of PDF, parse text, check for target year + kind keywords.
       
       target_kind: "annual" -> ["年報", "年度報告", "Annual Report"]
                    "interim" -> ["中期報告", "半年報", "Interim Report"]
       """
       try:
           with default_client(source="hkexnews") as client:
               resp = client.get(url, timeout=httpx.Timeout(connect=30.0, read=60.0))
               resp.raise_for_status()
               pdf_bytes = resp.content[:5 * 1024 * 1024]   # cap at 5 MB to avoid huge downloads
           reader = PdfReader(io.BytesIO(pdf_bytes))
           text = "".join(reader.pages[i].extract_text() for i in range(min(3, len(reader.pages))))
       except Exception:
           return False
       has_year = str(target_year) in text
       keywords = {"annual": ["年報", "年度報告", "Annual Report"], "interim": ["中期", "半年", "Interim"]}.get(target_kind, [])
       has_kind = any(kw in text for kw in keywords)
       return has_year and has_kind
   ```

2. `plugins/hk/reports.py::_select_main` 在多候选时调用：
   ```python
   if len(candidates) > 1:
       scored = []
       for cand in candidates[:5]:   # cap at top 5 to avoid wasting bandwidth
           score = _base_score(cand, period)
           if verify_pdf_year_and_kind(cand.url, period.year, _period_kind(period.type)):
               score += 10
           scored.append((score, cand))
       scored.sort(key=lambda x: -x[0])
       return scored[0][1]
   ```

#### B. 财年识别（非 12 月财年）

**问题**：港股有非 12 月财年公司，如汇丰（HSBC, 0005）12 月、阿里巴巴（9988）3 月、太古股份 12 月、长江基建（1038）12 月。

**实施**：

1. 加 `plugins/hk/fiscal_year.py`：
   ```python
   # Hong Kong listed companies with non-December fiscal year-end.
   # Format: hk_code (5-digit zero-padded) -> fiscal year end month (1-12).
   NON_DEC_FISCAL_YEAR: dict[str, int] = {
       "09988": 3,   # 阿里巴巴 BABA-SW
       "00788": 12,  # (example: 中国铁塔 12 月 -- 不在 non-dec 列表，删)
       # 添加 5-10 家常见非 12 月财年的港股
       # 数据来源：HKEX 公司年报披露 / 雪球公司资料
   }
   
   def fiscal_year_end_month(hk_code: str) -> int:
       """Return fiscal year-end month for HK ticker. Default 12."""
       return NON_DEC_FISCAL_YEAR.get(hk_code, 12)
   ```

2. `plugins/hk/reports.py` 选片时基于财年映射的 filing 时间窗口判断：
   ```python
   fye_month = fiscal_year_end_month(ticker.code)
   # 财年 N 的年报 filing 通常在财年结束后 3-5 个月内
   # 例：FYE = 3 月，FY2024 年报应在 2024-04-01 至 2024-09-30 公布
   expected_start = date(period.year, fye_month, 1) + relativedelta(months=1)
   expected_end = expected_start + relativedelta(months=4)
   ```

3. 给 `Period` 不需要加 `fiscal_year` 字段——通过 ticker code 查映射即可。

#### C. 多候选打分排序

**评分规则**（在 `_select_main` 内合并）：

```python
def _score_candidate(cand, period, ticker) -> int:
    score = 0
    title = cand.title or ""
    
    # 标题精确匹配
    if str(period.year) in title and any(kw in title for kw in ["年報", "Annual Report", "年度報告"]):
        score += 20
    
    # PDF 内嵌文本验证（仅多候选时调用，单候选跳过节省带宽）
    if multi_candidate_mode and verify_pdf_year_and_kind(cand.url, period.year, kind):
        score += 10
    
    # filingDate 在期望范围内
    fye = fiscal_year_end_month(ticker.code)
    if _filing_in_expected_window(cand.filing_date, period.year, fye):
        score += 5
    
    # 文件大小 > 1 MB
    if cand.file_size and cand.file_size > 1_000_000:
        score += 3
    
    # 排除补充公告
    if any(kw in title for kw in ["補充", "更正", "勘誤", "Supplementary"]):
        score -= 30
    
    return score
```

ties 用 filingDate 最早。

#### D. 加 `tests/test_hk_selection.py`

5 个 fixture 场景（每个有 3-5 个候选 dict，期望选中正确的）：

```python
TENCENT_2023 = {
    "ticker_code": "00700",
    "period_year": 2023,
    "candidates": [
        {"title": "年報 2023", "filing_date": "2024-03-22", "file_size": 8_500_000, "url": "https://..."},
        {"title": "ESG 報告 2023", "filing_date": "2024-04-15", "file_size": 3_200_000, "url": "https://..."},
        {"title": "補充公告：股東週年大會通告", "filing_date": "2024-05-01", "file_size": 200_000, "url": "..."},
    ],
    "expected_idx": 0,
}
```

测试 mock 掉 `verify_pdf_year_and_kind`（避免实际下载），断言选中第 0 个。

5 个场景：
- 腾讯 00700（普通 12 月财年）
- 阿里巴巴 9988（3 月财年）
- 汇丰 0005（12 月财年，多份子公司报告）
- 友邦保险 1299（12 月财年，IFRS）
- 中国移动 0941（12 月财年，标题中文/英文混合）

#### E. 不能让 600519/AAPL 退化

- 跑 `test_selection_logic.py`，预期保留的测试全 pass
- 端到端冒烟时 A 股 + 美股 + 港股各跑一个 ticker

### 7.2 验证

- 单元测试：`tests/test_hk_selection.py` 5 场景 pass
- 现有 `test_selection_logic.py` + `test_output_layout.py` 不退化
- 端到端：跑 00700 2023 年报，应该选中 `年報 2023`

### 7.3 严禁动（Sprint 03 专属）

- **不动 A 股 / 美股 / 韩股 plugin**（PDF 验证只对 HK）
- 不改 `output_paths.py`
- 不改 ABC

---

## 8. Sprint 04 — UX 打磨（4-5 h）

### 8.1 子任务

#### A. 首次启动 onboarding

**触发**：检测 `config.toml` 不存在（首次双击 EXE）。

**实施**：

1. `app/main.py` 启动后判断：
   ```python
   from app.core.settings import settings_path
   
   if not settings_path().exists():
       _show_onboarding(window)
   ```

2. `app/ui/onboarding_dialog.py`（新建）：QDialog，内容：
   - 欢迎语：「FS Capture 帮你一键下载 4 市场上市公司官方披露 PDF」
   - 「输入第一个股票代码试试」提示
   - 韩股 DART Key 注册引导：链接到 https://opendart.fss.or.kr/
   - 按钮：「稍后再说」/「现在就配 DART」（后者打开 settings_dialog）

3. 关闭对话框后写默认 `config.toml` 到 `Path(sys.executable).parent`（或源码模式 `development/`）。

#### B. 北交所代码兼容

**当前**：`plugins/ashare/reports.py::_column_for(code)` 用 `bse` 前缀给 cninfo。

**验证**：cninfo 实际 column 参数是 `bj`（北交所），不是 `bse`。Codex 自行用 `httpx` 直接打一次 cninfo 接口（带 `column=bj` 与 `column=bse` 各一次），看哪个返回有效结果。

**修复**（结合验证结果）：
- 如果 `bj` 有效：把 `bse` 改成 `bj`
- 如果 cninfo 同时支持两个：保留兼容
- 写一个北交所 ticker 测试（如 `430047.BJ` 类似）到 `tests/test_selection_logic.py`

#### C. KR Q1 vs Q3 关键词区分

**当前**：`plugins/kr/reports.py::_select_filing` 通过 `rcept_dt` 月份区分（已在 `tests/test_selection_logic.py::KRReportSelectionTests` 锁定）。

**任务**：确认实现与测试一致；如发现 fall-through bug（如某月 Q1/Q3 同时命中），加显式 month range：
- Q1：`rcept_dt` 月份在 4-6 月
- Q3：`rcept_dt` 月份在 10-12 月

如已经如此实现，跳过本子任务。

#### D. IPO 招股书完整支持

**前提**：`README.md` 提到「IPO 招股书」是支持的期间类型，但 `models.py::PeriodType` 是否包含？检查代码状态。

**实施**（按状态分两情况）：

**情况 1：`PeriodType` 没有 IPO**
1. `app/core/models.py::PeriodType` 加：
   ```python
   IPO_PROSPECTUS = "ipo_prospectus"
   ```
   并在 `display_name` 字典加 `"IPO 招股书"`。

2. 4 plugin 各加 IPO 路径：
   - **A 股**：cninfo `category=category_zsbg_szsh`（招股书）；时间窗口宽放（用上市日期 ±2 年）
   - **港股**：HKEXnews `t1code=20000`（IPO 类）+ `t2code=20100`（招股说明书）
   - **美股**：SEC submissions API form `S-1` / `S-1/A` / `F-1` / `F-1/A`
   - **韩股**：DART `kind="B"`（증권신고서/上市申请书）+ 关键词 `"증권신고서"`

3. `app/ui/period_selector.py` 加「IPO 招股书」checkbox。

**情况 2：`PeriodType` 已有 IPO**
- 验证 4 plugin 都实现，没实现的补上
- 测试每个市场至少一个 IPO ticker（如 A 股新股；US 用最近 IPO）

#### E. PDF sidecar 元数据

**目标**：每个下载的 PDF 旁写一个 `.meta.json`，含元信息便于用户后续查验。

**实施**：

1. 加 `app/core/sidecar.py`：
   ```python
   import hashlib
   import json
   from datetime import datetime, timezone
   from pathlib import Path
   from app.core.models import ReportFile
   
   def write_sidecar(report: ReportFile) -> Path:
       pdf_path = Path(report.local_path)
       sidecar = pdf_path.with_suffix(pdf_path.suffix + ".meta.json")
       
       file_size = pdf_path.stat().st_size
       sha256 = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
       
       meta = {
           "exchange": report.ticker.exchange.value,
           "ticker_code": report.ticker.code,
           "ticker_name": report.ticker.name,
           "period_year": report.period.year,
           "period_type": report.period.type.value,
           "kind": report.kind,
           "title": report.title,
           "source_url": report.source_url,
           "downloaded_at": datetime.now(timezone.utc).isoformat(),
           "file_size_bytes": file_size,
           "sha256": sha256,
       }
       sidecar.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
       return sidecar
   ```

2. 4 plugin 的 `download_reports` 在每个 PDF 下载完后调用 `write_sidecar(report)`。或者在 orchestrator 层统一处理：`_TaskRunnable.run` 的 download_reports 完成后遍历 `r.reports` 写 sidecar。后者更集中，更好。

3. 测试：`tests/test_sidecar.py`，构造 ReportFile，调 write_sidecar，断言生成 .meta.json 含正确字段。

#### F. 顺手清理

- README 中 playwright 兜底未实现的描述（如还在）
- `pyproject.toml` 移除 `openpyxl` 依赖（v0.2 已不需要）

### 8.2 验证

- 单元测试：`test_sidecar.py` + 北交所 ticker 测试
- 现有所有 unittest 不退化
- 端到端：跑 600519 年报，确认 `output/` 同时有 `*.pdf` 和 `*.pdf.meta.json`
- 首次启动：删 `config.toml` 后双击 EXE，确认 onboarding 弹出

### 8.3 严禁动（Sprint 04 专属）

- 不动 ABC 签名
- 不动 GUI 视觉（main_window / styles）
- onboarding dialog 用 QSS 已有 token，不引入新颜色

---

## 9. Sprint 05 — 测试基建升级（4-6 h）

### 9.1 子任务

#### A. pytest 迁移

1. `pyproject.toml` 加：
   ```toml
   [tool.pytest.ini_options]
   testpaths = ["tests"]
   pythonpath = ["."]
   addopts = "-v --strict-markers"
   markers = [
       "e2e: end-to-end smoke tests requiring network",
   ]
   ```

2. `requirements.txt` 加：
   ```
   pytest>=8.0
   pytest-httpx>=0.30
   ```
   （pytest-httpx 用于 mock httpx）

3. 现有 unittest 测试转 pytest：
   - `class XxxTests(unittest.TestCase)` 改成 `class TestXxx:` 或独立函数
   - `self.assertEqual(a, b)` → `assert a == b`
   - `self.assertIn(...)` → `assert ... in ...`
   - 保留覆盖：`test_output_layout` / `test_selection_logic` / `test_hk_selection`（Sprint 03 加的）/ `test_orchestrator`（Sprint 02 加的）/ `test_sidecar`（Sprint 04 加的）

4. 验证：`pytest tests/` 全 pass。

#### B. plugin 集成测试 (mock-based)

`tests/integration/test_plugins.py`（新建），覆盖每市场 plugin × 2 步骤 = 8 测试：

```python
import pytest
import respx
import httpx
from plugins.ashare import name_resolver as ashare_name
from plugins.hk import name_resolver as hk_name
# ...

class TestAShareResolveName:
    @respx.mock
    def test_resolves_known_code(self):
        # mock akshare HTTP
        respx.get("...").mock(return_value=httpx.Response(200, json={...}))
        ticker = ashare_name.resolve_name("600519")
        assert ticker.name == "贵州茅台"
    
    @respx.mock
    def test_handles_404(self):
        respx.get("...").mock(return_value=httpx.Response(404))
        with pytest.raises(ValueError):
            ashare_name.resolve_name("999999")
    
    @respx.mock
    def test_remote_disconnected_retried(self):
        # 第 1 次抛 RemoteDisconnected, 第 2 次成功
        respx.get("...").mock(side_effect=[
            httpx.RemoteProtocolError("disconnect"),
            httpx.Response(200, json={...}),
        ])
        ticker = ashare_name.resolve_name("600519")
        assert ticker.name == "贵州茅台"
```

8 个测试覆盖：

| Plugin | resolve_name | download_reports |
|---|---|---|
| ashare | 4xx + retry | 选片正确 + 4xx |
| hk | 单股查询 + RemoteDisconnected | 多候选选片（用 Sprint 03 fixture） |
| us | 已知 ticker + BRK.B 变体 | recent.* + 分页 fallback |
| kr | corp_codes lookup | 月份选片（Q1/Q3） |

#### C. e2e smoke 脚本

`tests/e2e/test_smoke.py`（新建）：

```python
import pytest
from plugins import get_plugin
from app.core.models import Exchange, Period, PeriodType

@pytest.mark.e2e
class TestEndToEndSmoke:
    def test_ashare_600519_2024(self, tmp_path):
        plugin = get_plugin(Exchange.A_SHARE)
        ticker = plugin.resolve_name("600519")
        assert ticker.name == "贵州茅台"
        
        period = Period(year=2024, type=PeriodType.ANNUAL)
        company = plugin.fetch_company(ticker)
        assert company.industry  # not empty
        
        reports = plugin.download_reports(ticker, period, tmp_path)
        assert any(r.kind == "annual_report" for r in reports)
        for r in reports:
            assert Path(r.local_path).exists()
            assert Path(r.local_path).stat().st_size > 100_000  # > 100 KB
    
    def test_us_aapl_2024(self, tmp_path): ...
    def test_hk_00700_2024(self, tmp_path): ...
    # KR skip without DART key
```

3 测试（A/HK/US），KR 在 Sprint 05 仍跳过（用户未注册 Key）。

CI 默认不跑 e2e（`pytest -m "not e2e"`）。

#### D. ruff lint

1. `requirements.txt` 加 `ruff>=0.6.0`（dev）

2. `pyproject.toml` 加：
   ```toml
   [tool.ruff]
   line-length = 100
   target-version = "py311"
   exclude = [".venv", "build", "dist", "_internal"]
   
   [tool.ruff.lint]
   select = ["E", "F", "I", "B", "UP", "N"]
   ignore = ["E501"]   # line too long handled by formatter
   
   [tool.ruff.lint.per-file-ignores]
   "tests/**" = ["B", "N"]
   ```

3. 跑 `ruff check .` + `ruff format .`，修所有 lint 错误（除非数量太多需 Planner 决策）。

#### E. CI: GitHub Actions

`.github/workflows/ci.yml`（新建）：

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:

jobs:
  lint-and-test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - working-directory: development
        run: |
          pip install -r requirements.txt
          ruff check .
          pytest -m "not e2e"
```

注意：项目用 PySide6，CI Linux 跑 GUI 测试需要 `xvfb-run` 或者 mock Qt。先跑非 GUI 测试（plugin 集成 + 单元），GUI 测试 CI skip。

### 9.2 验证

- `pytest tests/` 全 pass（包括迁移过来的所有测试）
- `pytest -m "not e2e" tests/` 同样 pass（CI 路径）
- `ruff check .` 0 errors
- `ruff format --check .` 0 changes（已格式化）
- `.github/workflows/ci.yml` 文件 yaml 合法（可用 `yamllint` 验）

### 9.3 严禁动（Sprint 05 专属）

- 不改业务逻辑（plugin / orchestrator / UI 内部）
- 测试改写不能删覆盖（已有断言全部保留）
- ruff 改不动 ABC 签名 / 不改 frozen 区代码

---

## 10. 整体验证矩阵（5 sprint 全部完成后）

完成 Sprint 05 后跑：

```bat
cd development

REM 1. 单元 + 集成测试
pytest -m "not e2e"

REM 2. lint
ruff check .

REM 3. grep 兜底（确认 v0.2 重定位 wording 全清）
findstr /S /M "FinancialStatement" *.py
findstr /S /M "fetch_financials" *.py
findstr /S /M "write_workbook" *.py
findstr /S /M "_LOOKUP_ALIASES" *.py

REM 4. 端到端（如可联网）
pytest -m e2e
```

如果 e2e 沙箱跑不了，写在最终回报里让 Reviewer 跑。

期望：所有上述命令 0 错误。

---

## 11. 完成回报格式

5 sprint 全部完成后，在 `current.md` 末尾追加：

```markdown
## 整体回报（Sprint 01-05 全部完成）

**总 commits**: 5（每 sprint 1 个）
- Sprint 01: <hash> - 工具重定位
- Sprint 02: <hash> - 健壮性
- Sprint 03: <hash> - HK 选片精确化
- Sprint 04: <hash> - UX 打磨
- Sprint 05: <hash> - 测试基建

**v0.1 → v0.2 LOC 变化**：
- Python: ~3,900 → ~3,XXX 行
- 新增测试: tests/test_orchestrator + tests/test_hk_selection + tests/test_sidecar + tests/integration/ + tests/e2e/

**最终验证矩阵**：
- pytest -m "not e2e": N passed
- ruff check: 0 errors
- grep 兜底: 全 0 命中
- 端到端冒烟: 已跑 / 沙箱无网络

**意外问题与决策**：
- (列出过程中所有需要 Reviewer 注意的点)

**已知未解决项**：
- KR e2e 验证待用户注册 DART Key
- (其他)
```

之外，在每个 sprint 提交时，提交说明已经包含了细节（commit message + sprint 段在 current.md）。

---

## 12. 给 Reviewer 的提示

Reviewer（Claude Code）会在你交付后做：

1. 读 `current.md` 整体回报 + 各 sprint 段
2. 对照 `roadmap/SPRINT_NN_*.md`（已存在的）+ 本文件 §6-§9 的 spec 逐项确认
3. 跑 §10 整体验证矩阵
4. 抽样 spot-check 关键修改（Sprint 01 grep 兜底；Sprint 02 cancel 路径；Sprint 03 HK 选片打分；Sprint 04 sidecar 内容；Sprint 05 CI yaml）
5. 写 `current.md` 末尾 `## Reviewer 验收` 段

如有阻塞 / 退化 / 需返工的项，Reviewer 写 follow-up 任务清单到 `current.md`，你再次按其修复。

---

**开始执行：先读 §2 必读文档清单按顺序读完，确认理解后从 Sprint 01 开始。Good luck.**

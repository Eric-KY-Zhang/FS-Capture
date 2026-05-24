# SPRINT v0.8 — 性能优化 + UI 字符串集中化 + Lint 锁定

**日期**：2026-05-23
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex
**Reviewer**：Claude Code
**状态**：已实施于 v0.8 internal（2026-05-23，Codex 批次 1-6）
**预计工作量**：2 周
**发布策略**：内部迭代，**不发 GitHub release**（按用户指示，GitHub release 延后到 `SPRINT_v1.0_sg_and_perf.md`）

---

## Context

v0.7 验收通过（11 commit `2dc5fd2 → d07d133`，85/85 tests，KR 无 Key 实跑 4 家 + audit 路径打通）。v0.8 进入性能与清债校准期。

经 v0.8 SPRINT 起草前的代码侦察，**原 ROADMAP v0.8 三项已不成立**：

| 原 ROADMAP 项 | 状态 | 实测证据 |
|---|---|---|
| ~~ruff 0 warning（清 114 个历史 warnings）~~ | ❌ 已完成 | `python -m ruff check development/` 输出 `All checks passed!`。`pyproject.toml` 已 `select = ["E","F","I","B","UP","N"]`，仅 ignore `E501`/`UP042`，且未在 `[tool.ruff]` 排除 `app/` 或 `plugins/`。本期只需把这状态在 pre-commit 锁定 |
| ~~Playwright 移除决断~~ | ❌ 不可行 | v0.7 KR 4 家 smoke 实跑显示 audit 报告 100% 走 `render_url_to_pdf`（DART 不暴露 audit 公告的直接 PDF 下载）；US HTML filings 同样依赖。**Playwright 是必要基础设施，本期改为池化优化** |
| ~~diskcache 批写 `cache.transact()`~~ | ⏳ 挪到后续性能 sprint | grep 显示当前 `cache.set` 都不在热循环里（name_map 是一次性大 dict 写）；ROI 低 |

经用户确认（2026-05-23），**bundle 体积优化挪到后续体积 sprint**：当前 425MB（Playwright 104M + llvmlite 102M + PySide6 93M = 70%），本期不动。

### v0.8 实际范围（5 项，已完成）

| Part | 内容 | ROI |
|---|---|---|
| **A.1** | Playwright 浏览器池化（替换 per-call `sync_playwright()`） | 🔴 高 — 批量 HK/US/KR 渲染省 30-50% 时间 |
| **A.2** | 大 PDF 断点续传（`stream_to_file` 保留 `.part` + `Range` header） | 🟡 中 — TW IPO/HK 大年报断网恢复 |
| **A.3** | `_load_name_map` / `resolve_one` 加锁防惊群（5 plugin） | 🟡 中 — 4 worker 启动省 3 次冗余请求 |
| **B** | UI 中文字符串物理集中化（97 处 / 11 文件 → `app/ui/strings.py`） | 🟢 低（清债，未来 i18n 铺路） |
| **C** | Lint 锁定（pre-commit hook + CI workflow） | 🟢 低 |

### 实施记录

- 批次 1：`cached_or_load` 单飞缓存 + name_resolver 接入 + 缓存并发测试。
- 批次 2：`stream_to_file` `.part` 断点续传 + 过期 part 清理 + 续传测试。
- 批次 3：Playwright singleton browser pool + 应用退出清理 + KR audit e2e smoke。
- 批次 4：UI 中文字符串集中到 `app/ui/strings.py` + tokenize 锁定测试。
- 批次 5：pre-commit、GitHub Actions、ruff isort 配置锁定。
- 批次 6：测试打桩随 `cached_or_load` 更新，文档收尾，全量非 e2e 回归。

---

## Part A.1 — Playwright 浏览器池化（主项）

### A.1.1 现状

`development/app/core/pdf_renderer.py:31-84`：

```python
_RENDER_SEMAPHORE = threading.Semaphore(2)

def render_url_to_pdf(url, dest, *, source="generic", timeout_ms=120_000):
    from playwright.sync_api import sync_playwright
    ...
    with _RENDER_SEMAPHORE:
        with sync_playwright() as p:           # ← 每次调用启 playwright runtime
            browser = _launch_browser(p)       # ← 每次启 chromium 进程（5-10s）
            try:
                context = browser.new_context(...)
                page = context.new_page()
                page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
                ...
                page.pdf(path=str(tmp_dest), ...)
                context.close()
            finally:
                browser.close()                 # ← 渲染完关闭，下次重启
```

**问题**：4 worker 并发跑 KR 5 票 ANNUAL 时，5 份 audit 各跑一次 `_launch_browser` = 25-50 秒纯启动开销，远超单页渲染时间。

### A.1.2 改动要求

改写 `app/core/pdf_renderer.py`，引入 **process-level singleton 浏览器**：

**新增模块级状态**：
```python
import threading
from typing import Optional

_PW_LOCK = threading.Lock()
_PW_RUNTIME = None        # sync_playwright().__enter__() 的返回值
_PW_BROWSER = None        # 单例 Chromium
_PW_CONTEXT_SEMAPHORE = threading.Semaphore(4)  # 并发 context 数（>= max_workers）
```

**lazy init**：
```python
def _ensure_browser():
    global _PW_RUNTIME, _PW_BROWSER
    with _PW_LOCK:
        if _PW_BROWSER is not None:
            return _PW_BROWSER
        from playwright.sync_api import sync_playwright
        _PW_RUNTIME = sync_playwright().start()
        _PW_BROWSER = _launch_browser(_PW_RUNTIME)
        return _PW_BROWSER
```

**render_url_to_pdf 改造**：
- 不再 `with sync_playwright()` 也不再 `browser.close()`
- 每次只 `browser.new_context()` → `new_page()` → `page.pdf()` → `context.close()`
- 用 `_PW_CONTEXT_SEMAPHORE` 限制并发 context 数（避免 Chromium OOM）

**清理**：新增 `shutdown_renderer()`，在 `app/main.py` 退出时调用：
```python
def shutdown_renderer() -> None:
    global _PW_RUNTIME, _PW_BROWSER
    with _PW_LOCK:
        if _PW_BROWSER is not None:
            try:
                _PW_BROWSER.close()
            except Exception as exc:
                logger.warning(f"playwright browser close failed: {exc}")
            _PW_BROWSER = None
        if _PW_RUNTIME is not None:
            try:
                _PW_RUNTIME.stop()
            except Exception as exc:
                logger.warning(f"playwright runtime stop failed: {exc}")
            _PW_RUNTIME = None
```

`app/main.py` 现有 `atexit` 链中接入 `shutdown_renderer()`（与 `close_cache()` 并列）。

### A.1.3 测试

新建 `development/tests/test_pdf_renderer_pool.py`：

1. **`test_render_url_to_pdf_reuses_browser_across_calls`**
   - monkeypatch `_launch_browser` 计数调用次数
   - 连续调用 `render_url_to_pdf` 3 次（用 `data:text/html,<html><body>x</body></html>` 或 file:// URL）
   - 断言 `_launch_browser` 仅被调用 1 次（singleton 复用）
   - 测试结束后调用 `shutdown_renderer()` 清理

2. **`test_shutdown_renderer_is_idempotent`**
   - 连续调用 `shutdown_renderer()` 两次，不抛错
   - 第二次调用时 `_PW_BROWSER is None`

3. **`test_render_url_to_pdf_serializes_via_context_semaphore`**（可选）
   - 用 mock browser + 计时器证明并发 context 数 ≤ `_PW_CONTEXT_SEMAPHORE._value`

### A.1.4 风险

| 风险 | 缓解 |
|---|---|
| Chromium 子进程长寿命累积内存 | `_PW_CONTEXT_SEMAPHORE = 4` 限并发；每次 `context.close()` 清理页面状态；每渲染 50 次自动 `browser.close()` 重启（可选，本期不做） |
| QThreadPool 4 worker + sync_playwright 是否线程安全 | sync_playwright 文档说 "sync API can only be used from the same thread that created it"；新设计中 browser 由首个 worker 创建后所有 worker 共享 `browser.new_context()`，需实测验证（Chromium 内部对此是支持的，但 sync_playwright 包装是否暴露 thread issue 需测） |
| PyInstaller windowed 模式下 stdio = None 与 Playwright | v0.6 已经测过 KR 走 Playwright 不会闪退，本期改动不引入新 stdout 使用 |

**降级路径**：如果 singleton 复用在某 worker 抛 `playwright._impl._api_types.Error`，捕获后调用 `shutdown_renderer()` + `_ensure_browser()` 重启 browser，重试 1 次。

---

## Part A.2 — 大 PDF 断点续传

### A.2.1 现状

`development/app/core/http.py:101-134`：

```python
def stream_to_file(client, url, dest, *, source, rate, chunk_size=65536, read_timeout=None):
    limiter(source, rate).acquire_blocking()
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")
    n = 0
    try:
        ...
        with client.stream("GET", url, timeout=stream_timeout) as r:
            r.raise_for_status()
            with tmp_dest.open("wb") as f:          # ← 总是 wb，不读已有 .part
                for chunk in r.iter_bytes(chunk_size):
                    f.write(chunk)
                    n += len(chunk)
        tmp_dest.replace(dest)
    except Exception:
        tmp_dest.unlink(missing_ok=True)            # ← 失败必删，无法续传
        raise
```

**问题**：HK 大 PDF（HSBC 年报 30+ MB）、TW IPO（单 PDF 10-50 MB），慢网或瞬断会导致从 0 开始重下；4 worker 并发时一个失败拖整轮。

### A.2.2 改动要求

修改 `stream_to_file`：

```python
def stream_to_file(client, url, dest, *, source, rate, chunk_size=65536, read_timeout=None):
    limiter(source, rate).acquire_blocking()
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_dest = dest.with_name(f"{dest.name}.part")

    # 断点续传：如果 .part 已存在且非空，发 Range 请求续传
    resume_from = tmp_dest.stat().st_size if tmp_dest.exists() else 0
    headers = {}
    open_mode = "wb"
    if resume_from > 0:
        headers["Range"] = f"bytes={resume_from}-"
        open_mode = "ab"
        logger.info(f"resuming {url} from byte {resume_from}")

    n = resume_from
    try:
        stream_timeout = httpx.Timeout(connect=30.0, read=read_timeout, write=30.0, pool=30.0)
        with client.stream("GET", url, headers=headers, timeout=stream_timeout) as r:
            # 服务器不支持 Range（返回 200 而非 206）→ 从头来
            if resume_from > 0 and r.status_code == 200:
                logger.warning(f"server ignored Range header for {url}, restarting from 0")
                resume_from = 0
                n = 0
                open_mode = "wb"
            r.raise_for_status()
            with tmp_dest.open(open_mode) as f:
                for chunk in r.iter_bytes(chunk_size):
                    f.write(chunk)
                    n += len(chunk)
        tmp_dest.replace(dest)
    except Exception:
        # 关键改变：保留 .part 文件供下次续传，不再 unlink
        logger.warning(f"stream_to_file failed for {url}, kept .part ({n} bytes) for resume")
        raise
    logger.debug(f"downloaded {url} -> {dest} ({n} bytes)")
    return n
```

**配套清理**：v0.8 引入 `.part` 永驻可能堆积。新增 `app/core/output_paths.py::cleanup_stale_parts(output_root, max_age_days=7)`：
- 扫 `output_root/**/*.part`，删除超过 7 天的孤儿
- 在 `Orchestrator.start_job` 入口调一次（非阻塞，best-effort）

### A.2.3 测试

新建 `development/tests/test_stream_resume.py`：

1. **`test_stream_to_file_resumes_from_existing_part`**
   - 预写 `dest.part` 内容 `b"AAA"` (3 bytes)
   - mock httpx client：断言 GET 请求头包含 `Range: bytes=3-`，返回 206 + body `b"BBB"`
   - 调用 `stream_to_file`，断言最终 `dest` 内容 = `b"AAABBB"`，返回 6

2. **`test_stream_to_file_restarts_when_server_ignores_range`**
   - 预写 `dest.part` 内容 `b"AAA"`
   - mock httpx client：GET 带 Range header，但服务器仍返回 200 + body `b"FULL"`
   - 调用 `stream_to_file`，断言 `dest` 内容 = `b"FULL"`（不是 `b"AAAFULL"`）

3. **`test_stream_to_file_keeps_part_on_failure_for_resume`**
   - mock httpx client：写入 2 chunks 后抛 `httpx.ReadError`
   - 调用 `stream_to_file`，断言抛异常 + `dest.part` 存在 + 大小 = 已写字节数

4. **`test_cleanup_stale_parts_drops_old_orphans`**
   - 在 tmp_path 下造 2 个 `.part` 文件（一新一旧 mtime）
   - 调用 `cleanup_stale_parts(tmp_path, max_age_days=1)`
   - 断言旧的被删、新的保留

### A.2.4 风险

| 风险 | 缓解 |
|---|---|
| `.part` 永驻撑大磁盘 | `cleanup_stale_parts` 7 天兜底；用户 SPRINT 完成后实测一轮 5 市场 smoke 确认无残留 |
| 服务器返回 416 (Range Not Satisfiable) — 已下载文件长度等于服务器总长 | tenacity 第 2 次重试时已不再发 Range（因 .part 被 unlink 在 replace 时），但本设计中失败保留 .part，需在 416 时显式删 .part + 重下。**建议**：捕获 `httpx.HTTPStatusError` 中 `r.status_code == 416` 时 `tmp_dest.unlink()` + 重抛让 tenacity 重试 |
| `read_timeout=None` 在 US HTML 渲染路径下 | 现状已使用，本期不动，确认改动后仍兼容 |

---

## Part A.3 — `_load_name_map` / `resolve_one` 加锁防惊群

### A.3.1 现状

5 个 plugin 的 name_resolver 都是无锁 read-modify-write：

| Plugin | 函数 | Race 文件:行 |
|---|---|---|
| A 股 | `_load_name_map` | `plugins/ashare/name_resolver.py:59-72` |
| HK | (单股查询，无 race) | `plugins/hk/name_resolver.py:46-52` |
| US | `_load_map` | `plugins/us/name_resolver.py:16-37` |
| KR | `resolve_one` | `plugins/kr/name_resolver.py:42-72`（v0.7 新增） |
| TW | `_load_map` | `plugins/tw/name_resolver.py:107-134` |

**问题**：4 worker 并发提交 4 个 A 股 ticker 时，4 worker 都 cache miss → 4 次调 `ak.stock_info_a_code_name()`（每次拉 5000+ 行）。最终 cache.set 是 last-write-wins，无数据损坏，但浪费 3 次冗余请求 + 5-10 秒。KR `resolve_one` 同理：4 个 KR ticker 各查 DART 公网一次反而被限速。

### A.3.2 改动要求

新增 `app/core/cache.py` helper：

```python
import threading
from typing import Callable, TypeVar

T = TypeVar("T")

# Per-key lock dict to avoid 5 plugin's separate fetches stomping each other.
_PER_KEY_LOCKS: dict[str, threading.Lock] = {}
_LOCKS_MUTEX = threading.Lock()


def _get_key_lock(key: str) -> threading.Lock:
    with _LOCKS_MUTEX:
        lock = _PER_KEY_LOCKS.get(key)
        if lock is None:
            lock = threading.Lock()
            _PER_KEY_LOCKS[key] = lock
        return lock


def cached_or_load(key: str, loader: Callable[[], T], *, expire: float | None) -> T:
    """Cache lookup with single-flight semantics — at most one loader call per key.

    - First reader: hit cache → return; miss → take per-key lock → re-check cache → call loader → set cache.
    - Concurrent readers: block on the same lock; after release, hit the populated cache.
    """
    cache = get_cache()
    cached = cache.get(key)
    if cached is not None:
        return cached
    lock = _get_key_lock(key)
    with lock:
        cached = cache.get(key)  # double-check: another thread may have populated
        if cached is not None:
            return cached
        value = loader()
        if value is not None:
            cache.set(key, value, expire=expire) if expire is not None else cache.set(key, value)
        return value
```

### A.3.3 改造各 plugin

每个 plugin 改写要点：

**`plugins/ashare/name_resolver.py:59-72`** — 替换 `_load_name_map`：

```python
def _load_name_map() -> dict[str, str]:
    def _fetch() -> dict[str, str]:
        import akshare as ak
        logger.info("Loading A-share code↔name map from akshare ...")
        df = ak.stock_info_a_code_name()
        return _strict_dict(df["code"].astype(str), df["name"].astype(str), context="code/name")
    return cached_or_load(_CACHE_KEY_NAME_MAP, _fetch, expire=_NAME_MAP_TTL)
```

**`plugins/us/name_resolver.py:16-37`** — 同 pattern
**`plugins/tw/name_resolver.py:107-134`** — 同 pattern
**`plugins/kr/name_resolver.py:42-72`** — `resolve_one` 内嵌锁逻辑（因为含分支：OpenAPI / 公网 fallback），不能直接套 `cached_or_load`。改为：

```python
def resolve_one(stock_code: str) -> dict | None:
    cache_key = f"kr:corp:{stock_code}"

    def _fetch():
        dart = _dart()
        if dart is not None:
            try:
                df = dart.corp_codes
                df = df[df["stock_code"].astype(str).str.zfill(6) == stock_code]
                if not df.empty:
                    row = df.iloc[0]
                    return {"corp_code": str(row["corp_code"]), "corp_name": str(row["corp_name"])}
            except Exception as exc:
                logger.warning(f"DART OpenAPI corp lookup failed for {stock_code}: {exc}")
        from .dart_web import resolve_corp
        logger.info(f"DART OpenAPI not configured; falling back to public crawler for {stock_code}")
        return resolve_corp(stock_code)

    return cached_or_load(cache_key, _fetch, expire=None)
```

注意：`cached_or_load` 接受 `expire=None` 表示永久缓存（diskcache 默认行为），KR corp_code 应永久缓存。

### A.3.4 测试

新建 `development/tests/test_cached_or_load.py`：

1. **`test_cached_or_load_single_flight_under_concurrency`**
   - 用 `concurrent.futures.ThreadPoolExecutor(max_workers=4)` 同时 4 线程调 `cached_or_load("test:race", loader, expire=None)`
   - `loader` 内部 `time.sleep(0.5)` + 全局计数 += 1，返回 `"value"`
   - 断言：4 个 future 全部返回 `"value"`；计数 == 1（loader 仅被调一次）

2. **`test_cached_or_load_returns_cached_immediately_on_hit`**
   - 预先 `get_cache().set("test:hit", "X")`
   - 调用 `cached_or_load("test:hit", loader=lambda: pytest.fail("loader should not run"), expire=None)`
   - 断言返回 `"X"`，loader 未被调

3. **`test_cached_or_load_does_not_cache_none_result`**
   - `loader` 返回 None
   - 调用 `cached_or_load("test:none", loader, expire=None)`
   - 断言返回 None；后续 `get_cache().get("test:none") is None`（不污染缓存）

4. **回归测试**：`tests/integration/test_plugins.py` 中现有 5 个 KR + A股 + US 测试继续通过（monkeypatch 入口未变）

---

## Part B — UI 中文字符串物理集中化

### B.1 现状

经 tokenize 准确扫描，**97 处 CJK 字符串 / 11 个 UI 文件**：

| 文件 | 字符串数 |
|---|---|
| `exchange_panel.py` | 16 |
| `main_view.py` | 15 |
| `ticker_row.py` | 14 |
| `period_selector.py` | 10 |
| `progress_dock.py` | 9 |
| `exchange_selector.py` | 8 |
| `onboarding_dialog.py` | 8 |
| `settings_dialog.py` | 7 |
| `batch_import_dialog.py` | 5 |
| `output_card.py` | 4 |
| `main_window.py` | 1 |

### B.2 改动要求

新建 `development/app/ui/strings.py`：

```python
"""UI string constants — centralized for future i18n.

Convention:
- Module-level constants, ALL_CAPS_SNAKE_CASE, grouped by source widget.
- Values are plain Chinese strings. **No tr() wrapper yet** — physical centralization only.
- Format strings keep f-string semantics: use `STR.format(...)` at use site, not embedded {var}.
"""

from __future__ import annotations

# ── main_window ─────────────────────────────────────────────────────────
WINDOW_TITLE = "FS Capture — 跨市场上市公司披露 PDF 下载工具"

# ── main_view ───────────────────────────────────────────────────────────
MV_RUN_BUTTON = "开始下载"
MV_SETTINGS_BUTTON = "设置"
MV_NO_VALID_TICKERS_TITLE = "无可用股票"
MV_NO_VALID_TICKERS_BODY = "请先在左侧输入并确认至少一只股票"
# ... (95 more)

# ── period_selector ────────────────────────────────────────────────────
PS_ANNUAL = "年报"
PS_Q1 = "一季报"
PS_Q2 = "半年报"
PS_Q3 = "三季报"
PS_IPO_PROSPECTUS = "IPO 招股说明书"
# ...

# (organize by source file, comment headers between sections)
```

### B.3 各文件机械替换

Codex 用 grep+rewrite 把每个 UI 文件中所有中文字符串替换为 `from app.ui import strings as S; ... S.PS_ANNUAL`。

**约束**：
- **不改任何业务逻辑**，纯字符串引用替换
- **不引入 `tr()` 或 `QCoreApplication.translate`**：这是物理集中，i18n 留待后续版本
- f-string `f"已选 {n} 只股票"` 改为 `S.MV_SELECTED_COUNT.format(n=n)`
- 短字符串如 `"是"` / `"否"` 可保留内联，不强制集中（避免过度抽象）
- `QMessageBox.question(self, "标题", "正文", ...)` 这种位置型参数照常调用 `S.XXX_TITLE`, `S.XXX_BODY`

### B.4 测试

新建 `development/tests/test_ui_strings.py`：

1. **`test_no_cjk_string_literals_remain_in_ui_modules`**
   - tokenize 扫描 `app/ui/*.py`（排除 `strings.py` 自身）
   - 断言 CJK string literal 数 == 0（或一个白名单 ≤ 5）
   - 这个测试**锁定**未来回归

2. **`test_strings_module_has_no_duplicates`**
   - import `app.ui.strings`，断言所有常量值唯一（避免无意义重复定义）

3. **不改 GUI 集成测试** — 现有 `test_ticker_row.py` / `test_batch_import.py` 不应受影响（它们断言行为，不断言字面字符串）

### B.5 风险

| 风险 | 缓解 |
|---|---|
| QSS 中嵌入中文 | grep `.qss` 文件确认无 CJK；如有，独立处理 |
| f-string 转 `.format()` 写错变量名 | tokenize 测试会抓到漏改的；smoke 实跑 EXE 一轮检查 UI 显示 |
| Codex 替换错位（用错 constant） | Reviewer 抽查 3-5 处对照 |

---

## Part C — Lint 锁定

### C.1 现状

`pyproject.toml` ruff 配置已严格（`select = ["E","F","I","B","UP","N"]`），实际运行 0 warning。**本期只需把这状态锁定**避免回归。

### C.2 改动要求

#### C.2.1 新建 `.pre-commit-config.yaml`

```yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.6.9
    hooks:
      - id: ruff
        args: [--fix]
        files: ^development/
      - id: ruff-format
        files: ^development/
```

（如果用户不用 pre-commit，本步可跳；但配文件先写好）

#### C.2.2 新建 `.github/workflows/ci.yml`（如已有则修改）

```yaml
name: CI
on: [push, pull_request]
jobs:
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: astral-sh/setup-uv@v3
      - name: Install ruff
        run: uv pip install --system ruff
      - name: Lint
        run: ruff check development/
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'
      - name: Install
        run: |
          pip install -e development/
      - name: Test (non-e2e)
        run: cd development && pytest -m "not e2e" -v
```

**注意**：v0.7 引入 bs4 + lxml，CI 上 pip install 时这俩需要 build wheel 可能慢。先观察一轮，必要时改 `uv pip install`。

#### C.2.3 现有 ruff 配置加一条 isort 锁定

`pyproject.toml` 已有 `"I"` selected。本期确认 `[tool.ruff.lint.isort]` 段有合理 `known-first-party`：

```toml
[tool.ruff.lint.isort]
known-first-party = ["app", "plugins"]
```

如不存在则新增。

### C.3 测试

- Codex 本地跑 `ruff check development/` 输出 `All checks passed!`
- 如有 pre-commit 装好的环境，跑 `pre-commit run --all-files` 全绿
- 不需要新增测试

---

## 实施顺序（Codex 分批 commit）

| 批次 | 内容 | 验证 |
|---|---|---|
| **1** | Part A.3 `cached_or_load` helper + 5 plugin 改造 + `tests/test_cached_or_load.py` 4 个用例 | `pytest tests/test_cached_or_load.py tests/integration/test_plugins.py -v` 全绿 |
| **2** | Part A.2 `stream_to_file` 续传逻辑 + `cleanup_stale_parts` + `tests/test_stream_resume.py` 4 个用例 + Orchestrator 接入 cleanup | `pytest tests/test_stream_resume.py -v` 绿 |
| **3** | Part A.1 Playwright pool 改造 `pdf_renderer.py` + `tests/test_pdf_renderer_pool.py` + `app/main.py` 接入 `shutdown_renderer()` atexit | `pytest tests/test_pdf_renderer_pool.py -v` 绿；手动 smoke 一轮 KR 005930 audit 渲染确认未崩溃 |
| **4** | Part B 新建 `app/ui/strings.py` + 11 UI 文件批量替换 + `tests/test_ui_strings.py` 2-3 个用例 | `pytest tests/test_ui_strings.py -v` + 启动 EXE 肉眼检查 UI 显示无错乱 |
| **5** | Part C `.pre-commit-config.yaml` + `.github/workflows/ci.yml` + `pyproject.toml` isort 段 | `ruff check development/` 全绿；如 CI 可推则推一份观察一轮 |
| **6** | 文档收尾：`PROJECT_RETROSPECTIVE.md §12 v0.8 postscript`；`ARCHITECTURE.md` `pdf_renderer` 段加"singleton 池化"；`CLAUDE.md` + `AGENTS.md` §10 更新版本进度（双文件同步） | 全量回归 `pytest -m "not e2e" -v` 全绿 |

每批结束后 Codex 提 commit，commit message 格式：`v0.8: <一句话变更>（批次 N）`。

---

## 测试矩阵

### 单元测试（必绿）

```bash
cd development
pytest -m "not e2e" -v
```

预期：v0.7 基线 85 passed + 本期新增约 13 个用例 → **98 passed**：
- `test_cached_or_load.py`：4
- `test_stream_resume.py`：4
- `test_pdf_renderer_pool.py`：3
- `test_ui_strings.py`：2

### e2e Smoke（用户实测）

```powershell
$env:FS_CAPTURE_RUN_E2E="1"
cd development
pytest tests/e2e -v
```

预期：5 市场全过；**特别关注 KR 005930 audit_report**（Playwright 池化后这条路径仍走 render，必须保活）。

### 手动 smoke（用户实测时建议跑一轮）

- **性能对照**：v0.7 KR 5 票（005930/013890/000660/035420 + 1 IPO）总耗时 vs v0.8（理论上 audit 渲染 5×节省启动时间 = ~25 秒省）
- **断点续传**：刻意网络掐断 TW 2330 IPO 下载到一半，重启抓任务，确认从上次 `.part` 字节继续
- **UI 视觉回归**：5 市场各点开一个面板，确认所有按钮 / 提示 / dialog 文字正常

---

## Reviewer Checklist

### Part A.1 — Playwright 池化
- [ ] `_PW_BROWSER` 是 module-level singleton（不在函数内定义）
- [ ] `_ensure_browser()` 用 `_PW_LOCK` 保护初始化
- [ ] `render_url_to_pdf` 不再 `with sync_playwright()` 也不再 `browser.close()`，只 `new_context()` + `context.close()`
- [ ] `shutdown_renderer()` 在 `app/main.py` 退出链中注册（atexit 或显式 finally）
- [ ] `_PW_CONTEXT_SEMAPHORE` 与 settings `max_workers` 对齐（>= 4）
- [ ] `_RENDER_SEMAPHORE` 旧的 `Semaphore(2)` 已移除或显式替换
- [ ] `tests/test_pdf_renderer_pool.py` 3 个用例齐备

### Part A.2 — 大 PDF 续传
- [ ] `stream_to_file` 失败时**不再 unlink `.part`**
- [ ] 续传时发 `Range: bytes={size}-` header
- [ ] 服务器返回 200（忽略 Range）时正确从 0 重写
- [ ] 服务器返回 416 时 unlink `.part` 触发 tenacity 重试
- [ ] `cleanup_stale_parts` 在 `Orchestrator.start_job` 入口被调用（非阻塞）
- [ ] `tests/test_stream_resume.py` 4 个用例齐备

### Part A.3 — 单线程拉取
- [ ] `app/core/cache.py::cached_or_load` 实现双重检查锁
- [ ] 5 个 plugin 的 `_load_name_map` / `resolve_one` 已用 `cached_or_load` 重写
- [ ] `cached_or_load` 在 `loader()` 返回 None 时**不缓存**（不污染缓存）
- [ ] `tests/test_cached_or_load.py` 含 single-flight 并发用例（ThreadPoolExecutor=4）
- [ ] 现有 `tests/integration/test_plugins.py` 5 个 plugin 集成测试**未改逻辑**且继续通过

### Part B — UI 字符串集中化
- [ ] `app/ui/strings.py` 存在，所有 97 个字符串都有对应常量
- [ ] 11 个 UI 文件中**无残留 CJK 字符串字面量**（白名单 ≤ 5）
- [ ] f-string 改造正确（用 `S.X.format(...)` 不丢变量）
- [ ] **不引入 `tr()` 或 `QCoreApplication.translate`**
- [ ] `tests/test_ui_strings.py` 含 "no_cjk_literals_remain" 锁定测试

### Part C — Lint 锁定
- [ ] `.pre-commit-config.yaml` 存在
- [ ] `.github/workflows/ci.yml` 含 lint + test job
- [ ] `pyproject.toml` 有 `[tool.ruff.lint.isort] known-first-party = ["app", "plugins"]`
- [ ] `ruff check development/` 输出 `All checks passed!`

### 质量门禁
- [ ] `pytest -m "not e2e" -v` 全绿，**总数 ≥ 98**（v0.7 基线 85 + 本期约 13）
- [ ] e2e（用户实测）：5 市场全绿，**特别确认 KR audit_report 路径仍走 Playwright 渲染**
- [ ] `app/core/orchestrator.py`：仅新增 `cleanup_stale_parts` 调用，无其他改动
- [ ] `app/core/http.py`：仅修 `stream_to_file`，`default_client` / `get_json` / `post_json` 不动
- [ ] 输出文件命名 schema 未变（`test_output_layout.py` 全绿）

### 文档
- [ ] `PROJECT_RETROSPECTIVE.md §12 v0.8 postscript`
- [ ] `ARCHITECTURE.md` `pdf_renderer` 一段加 "Singleton browser pool, lazy init at first render"
- [ ] **`CLAUDE.md` + `AGENTS.md` 同步更新**（§10 版本进度表 v0.8 标 ✅）
- [ ] **不发 GitHub release，不更新 README 顶部版本号**

---

## 关键文件路径（Codex 速查清单）

### 新建
- `development/app/ui/strings.py`
- `development/tests/test_pdf_renderer_pool.py`
- `development/tests/test_stream_resume.py`
- `development/tests/test_cached_or_load.py`
- `development/tests/test_ui_strings.py`
- `.pre-commit-config.yaml`（仓库根）
- `.github/workflows/ci.yml`（如已有则改）

### 修改
- `development/app/core/pdf_renderer.py`（Playwright 池化）
- `development/app/core/http.py`（`stream_to_file` 续传）
- `development/app/core/cache.py`（新增 `cached_or_load` + `_PER_KEY_LOCKS`）
- `development/app/core/output_paths.py`（新增 `cleanup_stale_parts`）
- `development/app/core/orchestrator.py`（在 start_job 入口调 `cleanup_stale_parts`）
- `development/app/main.py`（atexit 注册 `shutdown_renderer`）
- `development/plugins/ashare/name_resolver.py:59-72`
- `development/plugins/us/name_resolver.py:16-37`
- `development/plugins/tw/name_resolver.py:107-134`
- `development/plugins/kr/name_resolver.py:42-72`
- `development/app/ui/*.py`（11 个文件，UI 字符串替换）
- `development/pyproject.toml`（isort 段）
- `PROJECT_RETROSPECTIVE.md`（§12 postscript）
- `ARCHITECTURE.md`（pdf_renderer 段）
- `CLAUDE.md` + `AGENTS.md`（§10 版本进度同步）

### 不动
- `development/app/core/http.py::default_client` / `get_json` / `post_json`
- `development/app/core/ratelimit.py`
- `development/app/core/models.py` / `output_paths.py::report_output_path` / `sidecar.py`
- `development/plugins/*/reports.py`（KR/HK/TW/A股/US 所有 reports 模块）
- `development/plugins/kr/dart_web.py`（v0.7 新增，本期不动）
- `development/plugins/hk/name_resolver.py`（单股查询无 race，无需改）
- `development/plugins/hk/fiscal_year.py` / `_pdf_verify.py`
- 所有 e2e 测试（除非 Playwright 池化破坏 KR audit 路径需要回归补充）

---

## 风险与缓解

| 风险 | 影响 | 缓解 |
|---|---|---|
| Playwright singleton browser 在 QThreadPool 4 worker 跨线程 new_context 失败 | KR audit + US HTML 路径挂 | 实测一轮 KR 5 票 audit + US AAPL HTML；不行则降级 `_PW_CONTEXT_SEMAPHORE = 1`（虽然失去并发收益，但保正确性） |
| 续传 `.part` 文件被 PDF 写完前另一个 worker 抢读 | 文件损坏 | 文件命名按完整 dest 路径派生（`{dest}.part`），单文件一个 worker，不会跨 worker 抢；smoke 实测确认 |
| UI 字符串机械替换错位 | UI 显示乱码或错文字 | tokenize 测试 + smoke 启动 EXE 肉眼检查 |
| `cached_or_load` 与 lru_cache 装饰器混用导致缓存层级混乱 | 缓存 miss 率升 | KR `_dart_for_key` 仍用 `lru_cache(maxsize=4)`，与 disk 缓存层独立，不冲突 |
| CI workflow 触发后失败（bs4/lxml wheel 构建慢） | PR 卡 CI | 用 uv 加速；如仍超时可暂时只跑 lint job，test job 留本地 |

---

## 后续衔接

v0.8 验收通过后：
- Planner 起草 `roadmap/SPRINT_v1.0_sg_and_perf.md`，涵盖：
  - 日本 EDINET（新市场，6 番）
  - 伦敦 LSE/NSM（新市场，7 番）
  - 增量更新检测（基于 sidecar）
  - IPO 路径统一（5 plugin 收敛到 `report_output_path_for_filing`）
  - 模块边界文档化"如何加新市场"
  - **可选**：bundle 体积优化（llvmlite 剥离 / Chromium 按需下载）

`roadmap/SPRINT_v1.0_sg_and_perf.md` 是 GitHub release 首发版，需要：
- README 改版（新增市场列表、更新截图、增量模式介绍）
- CHANGELOG 从 v0.6.1 累积写完
- 打 release tag
- 跑完整 7 市场 e2e

---

## 致 Codex 的提示

1. **不动 v0.7 KR 公网爬虫**：`plugins/kr/dart_web.py` 在 v0.7 验收通过，本期不要改。
2. **A.1 优先级最高**：Playwright pool 是 v0.8 性能核心，建议先做 + 单独 commit + 单独 smoke 验证再做其他。
3. **B 是机械活但易出错**：建议每个 UI 文件改完跑一轮 `pytest -m "not e2e"`（虽然 UI 字符串改动 unit tests 不会抓到，但保证没误删 import）。改完最后**启动 EXE 肉眼扫一遍**。
4. **C 是收尾**：lint 锁定在所有改动完成后做，避免反复触发 ruff format 改动 diff。
5. **Reviewer 会跑**：`ruff check development/` 必须 `All checks passed!`；`pytest -m "not e2e" -v` 必须 ≥ 98 passed。

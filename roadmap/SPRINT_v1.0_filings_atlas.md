# SPRINT v1.0 — Filings Atlas / 全球披露图谱（首发 release）

**日期**：2026-05-23
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex
**Reviewer**：Claude Code
**状态**：实施中（批次 3 strings.py 双语拆分已完成，下一步批次 4 语言开关与 live re-translate）
**预计工作量**：4-6 周
**发布策略**：**首个 GitHub release**（v1.0.0），含 PyInstaller bundle artifact + 双语 README + CHANGELOG

---

## Context

v0.8 验收通过（6 批次，100/100 tests，Playwright pool + 断点续传 + UI strings 集中 + lint 锁定）。v1.0 是首个 GitHub release 版本，合并 **3 项用户新增需求** + **5 项 ROADMAP 原计划**：

| # | 范围 | 来源 |
|---|---|---|
| 1 | 品牌重命名 "FS Capture" → **"Filings Atlas / 全球披露图谱"** | 用户 2026-05-23 |
| 2 | 中英双语 UI（运行时切换） + README 双语版 | 用户 2026-05-23 |
| 3 | Sidecar JSON 从 `output/` 迁移到 `data/cache/sidecars/` | 用户 2026-05-23 |
| 4 | 日本市场 plugin（EDINET，双模式） | ROADMAP |
| 5 | 英国市场 plugin（NSM/FCA，公网为主） | ROADMAP |
| 6 | 增量更新检测（基于 sidecar diff） | ROADMAP |
| 7 | IPO/filing 路径统一（ashare + hk 收敛） | ROADMAP |
| 8 | "如何加新市场" 文档（追加 ARCHITECTURE.md） | ROADMAP |

**依赖主线**：重命名 → 双语脚手架 → 字符串拆分 → 开关接线 → IPO 统一 → sidecar 迁移 → 增量 → JP → UK → 文档 → 发布

**仓库目录名 `FS Capture/` 保留**（节省迁移成本，只调可见品牌）。

---

## 已锁定决策（与用户对齐）

| 决策 | 选定方案 |
|---|---|
| 产品名 | **Filings Atlas / 全球披露图谱** |
| 双语实现 | **Pattern B**：`STRINGS = {"zh": {...}, "en": {...}}` + `LanguageManager(QObject)` 单例 + `language_changed = Signal(str)` |
| Sprint 划分 | 单个 v1.0 mega sprint（11 批次） |
| Sidecar 落位 | `data/cache/sidecars/{exchange}/{stem}.meta.json` （`stem = accession_number or pdf_basename`） |
| 增量更新 UI | 独立按钮"增量更新 / Incremental" 与"抓报告"并列；不用 checkbox / 不放设置 |
| 加新市场文档 | 追加章节到 `ARCHITECTURE.md`（不新建独立 docs） |
| Live re-translate | 每个 widget 加 `_retranslate(self)` + 监听信号 in-place setText()；**禁止重建 widget**（保 TickerRow 用户输入） |
| 不动 | 仓库目录名、`VBA Captor/` 子项目、`http.py`/`ratelimit.py`/`cache.py`/`pdf_renderer.py`、5 现有 plugin（除 IPO helper 拆除） |

---

## Part 1 — 品牌重命名（批次 1）

### 1.1 改动文件

**Python 源码**：
- [development/app/main.py](development/app/main.py)：
  - line 47：`fs_capture.log` → `filings_atlas.log`
  - line 62-63：`setApplicationName("FS Capture")` → `setApplicationName("Filings Atlas")`、display name 同步
  - line 65：`fs_capture.ico` → `filings_atlas.ico`
  - line 75, 104, 110：log + 错误对话框 `"FS Capture 启动失败"` → 走 strings.py 双语 key（批次 4 完成后即生效；本批先硬编码 "Filings Atlas / 全球披露图谱"，批次 4 时改 key）
- [development/app/ui/main_window.py](development/app/ui/main_window.py)：line 36（logo label）、line 89（窗口 title）
- [development/app/core/settings.py:42](development/app/core/settings.py:42)：SEC UA `"FS Capture (kaiyu199602@gmail.com)"` → `"Filings Atlas (kaiyu199602@gmail.com)"`

**Build 配置**：
- `development/fs_capture.spec` → 重命名为 `development/filings_atlas.spec`，内部 `name="FS Capture"` → `name="Filings Atlas"`（line 1, 4, 23, 24, 58, 69, 80）
- [development/pyproject.toml](development/pyproject.toml)：
  - line 2：`name = "fs-capture"` → `name = "filings-atlas"`
  - line 33：console script `fs-capture` → `filings-atlas`
- `development/build.bat`：所有 `FS Capture.exe` / `dist\FS Capture` → `Filings Atlas.exe` / `dist\Filings Atlas`；spec 文件名同步
- `development/run.bat`：注释同步

**UI 资源**：
- `development/app/assets/fs_capture.ico` → `filings_atlas.ico`（`git mv` 二进制不变）
- `development/app/assets/fs_capture_logo.png` → `filings_atlas_logo.png`
- `development/app/ui/styles/app.qss:2`：header 注释品牌

**文档（仅 header pass）**：
- `README.md`、`ARCHITECTURE.md`、`PROJECT_RETROSPECTIVE.md`、`CLAUDE.md`、`AGENTS.md`、`development/DEVELOPMENT_BRIEF.md`
- `roadmap/archive/` 中历史 sprint 文档**不重写**，只在 ROADMAP 主文件加一行"v1.0 起重命名为 Filings Atlas"
- `roadmap/ROADMAP_v0.6.1_to_v1.0.md`、`roadmap/SPRINT_v1.0_filings_atlas.md`（本文件）

**保留不动**：
- `VBA Captor/` 子项目（另一独立产品的所有 "Captor" 引用）
- 仓库目录名 `FS Capture/`

### 1.2 验证

```bash
grep -ri "fs capture" development/app development/plugins
# 仅剩有意保留项；不应有"FS Capture"窗口标题/exe 名/console script

grep -ri "captor" development
# VBA Captor 引用应原封不动
```

启动 EXE 检查窗口标题、logo、about 对话框（如有）显示 "Filings Atlas / 全球披露图谱"。

---

## Part 2 — 双语 UI（批次 2-4）

### 2.1 批次 2 — 双语脚手架（零字符串变更）

#### 新建 `development/app/ui/i18n.py`

```python
from __future__ import annotations
from PySide6.QtCore import QObject, Signal

_VALID_LANGS = {"zh", "en"}

class LanguageManager(QObject):
    """Process-wide singleton for runtime UI language switching."""

    language_changed = Signal(str)  # emits new code

    _instance: "LanguageManager | None" = None

    def __init__(self) -> None:
        super().__init__()
        self._current_language = "zh"
        assert LanguageManager._instance is None, "LanguageManager is a singleton"

    @classmethod
    def instance(cls) -> "LanguageManager":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    @property
    def current_language(self) -> str:
        return self._current_language

    def set_language(self, code: str) -> None:
        if code not in _VALID_LANGS:
            raise ValueError(f"Unsupported language: {code!r}")
        if code == self._current_language:
            return  # idempotent
        self._current_language = code
        self.language_changed.emit(code)
```

#### 修改 [app/core/settings.py:51](development/app/core/settings.py:51)

```python
class UICfg(BaseModel):
    theme: str = "light"
    language: str = "zh"  # was "zh_CN" (vestigial)
    window_width: int = 1280
    window_height: int = 820
```

新增 `model_validator` 自动迁移旧值：

```python
@model_validator(mode="before")
@classmethod
def _migrate_language(cls, data: Any) -> Any:
    if isinstance(data, dict):
        lang = data.get("language", "")
        if lang in {"zh_CN", "zh-cn", "zh-CN", "chinese"}:
            data["language"] = "zh"
        elif lang in {"en_US", "en-us", "en-US", "english"}:
            data["language"] = "en"
    return data
```

#### 修改 [app/main.py](development/app/main.py)

启动时（设 QApplication 后、构造 MainView 前）：

```python
from app.ui.i18n import LanguageManager
LanguageManager.instance().set_language(settings.ui.language)
```

#### 新增 `development/tests/test_i18n.py`

```python
def test_set_language_emits_signal_on_change():
    mgr = LanguageManager.instance()
    received: list[str] = []
    mgr.language_changed.connect(received.append)
    mgr.set_language("en")
    assert received == ["en"]

def test_set_language_idempotent_when_unchanged():
    mgr = LanguageManager.instance()
    mgr.set_language("zh")
    received: list[str] = []
    mgr.language_changed.connect(received.append)
    mgr.set_language("zh")
    assert received == []

def test_set_language_rejects_invalid_code():
    with pytest.raises(ValueError):
        LanguageManager.instance().set_language("ja")

def test_settings_migrates_old_zh_CN():
    s = Settings.model_validate({"ui": {"language": "zh_CN"}})
    assert s.ui.language == "zh"
```

**验证**：所有现有测试继续绿（`strings.py` 未动）。

### 2.2 批次 3 — strings.py 双语拆分

#### 重构 [app/ui/strings.py](development/app/ui/strings.py)

```python
"""UI string constants — bilingual zh/en, keyed by widget."""

from __future__ import annotations
from app.ui.i18n import LanguageManager

STRINGS: dict[str, dict[str, str]] = {
    "zh": {
        "COMMON_CANCEL": "取消",
        "BID_WINDOW_TITLE_FORMAT": "批量添加{exchange_name}股票",
        "BID_TITLE": "粘贴股票代码",
        # ... 97 keys total
    },
    "en": {
        "COMMON_CANCEL": "Cancel",
        "BID_WINDOW_TITLE_FORMAT": "Batch Add {exchange_name} Tickers",
        "BID_TITLE": "Paste Ticker Codes",
        # ... 97 keys total — Codex must produce all 97 EN translations
    },
}


def tr(key: str) -> str:
    return STRINGS[LanguageManager.instance().current_language][key]


def __getattr__(name: str) -> str:
    """Backward-compat: `ui_strings.MV_SECTION` keeps working as lazy lookup."""
    if name in STRINGS["zh"]:
        return tr(name)
    raise AttributeError(f"module 'strings' has no attribute {name!r}")
```

#### 英文翻译 style guide（Codex 翻译时遵循）

| 中文倾向 | 英文统一用 |
|---|---|
| 抓取 / 抓 / 获取 | **Download**（不要 Fetch / Grab） |
| 报告 / 公告 | **Report** / **Filing**（按上下文） |
| 公司名称 | **Company name** |
| 期间 / 时段 | **Period** |
| 一季报 / 半年报 / 三季报 / 年报 | **Q1 / Interim / Q3 / Annual report** |
| 招股说明书 / 招股书 | **IPO prospectus** |
| 审计报告 | **Audit report** |
| 设置 | **Settings** |
| 输出 | **Output** |
| 市场全称 | **A-Share / Hong Kong / United States / Korea / Taiwan / Japan / United Kingdom**（不缩写） |
| 增量更新 | **Incremental update** |
| API 密钥 | **API key** |
| 可选 | **(Optional)** |

#### 升级 `development/tests/test_ui_strings.py`

- 保留 `test_no_cjk_string_literals_remain_in_ui_modules`（v0.8）
- 新增：
  - `test_strings_zh_and_en_have_identical_keys`：`set(STRINGS["zh"].keys()) == set(STRINGS["en"].keys())`
  - `test_strings_en_values_contain_no_cjk`：遍历 `STRINGS["en"].values()` 断言无 CJK 字符
  - `test_strings_zh_values_all_contain_cjk_or_are_format_anchors`：sanity check（允许 format anchors 如 `{count}` 不带 CJK）

#### 验证

```python
# Default zh
from app.ui import strings
print(strings.MV_RUN_BUTTON)  # 抓报告

# Switch
from app.ui.i18n import LanguageManager
LanguageManager.instance().set_language("en")
print(strings.MV_RUN_BUTTON)  # Download Reports
```

### 2.3 批次 4 — 语言开关 + 11 widget live re-translate

#### 设置面板加语言 combo

[app/ui/settings_dialog.py:42](development/app/ui/settings_dialog.py:42) 仿 theme 模式：

```python
self.language = QComboBox()
self.language.addItem("中文", "zh")
self.language.addItem("English", "en")
idx = self.language.findData(settings.ui.language)
self.language.setCurrentIndex(max(0, idx))
form.addRow(ui_strings.SD_LANGUAGE_LABEL, self.language)  # "界面语言" / "UI Language"
```

`_save()` 中：

```python
new_lang = self.language.currentData()
if new_lang != self.settings.ui.language:
    self.settings.ui.language = new_lang
    save_settings(self.settings)
    LanguageManager.instance().set_language(new_lang)
```

#### 11 widget 加 `_retranslate(self)`

为每个 UI 文件加：

```python
def __init__(self, ...) -> None:
    super().__init__(...)
    # ... 原有 init ...
    LanguageManager.instance().language_changed.connect(self._retranslate)

def _retranslate(self, _lang: str = "") -> None:
    """In-place setText for all user-visible labels/buttons."""
    self.run_btn.setText(ui_strings.MV_RUN_BUTTON)
    self.settings_btn.setText(ui_strings.MV_SETTINGS_BUTTON)
    self.section_label.setText(ui_strings.MV_SECTION)
    # ... etc
```

需改的 11 个文件：
- `app/ui/main_view.py`
- `app/ui/main_window.py`
- `app/ui/exchange_selector.py`
- `app/ui/exchange_panel.py`
- `app/ui/output_card.py`
- `app/ui/period_selector.py`
- `app/ui/progress_dock.py`
- `app/ui/settings_dialog.py`（自身也需要响应——设置 dialog 关闭后重开时显示新语言）
- `app/ui/batch_import_dialog.py`
- `app/ui/onboarding_dialog.py`
- `app/ui/ticker_row.py`

**严格规则**：
- 用 `setText()`，**不重建 widget**（TickerRow 用户已输入的 code 必须保留）
- ExchangeSelector 中市场 chip 文案变（"沪深 / A-Share" 等），但 chip 排列顺序、当前选中态保留
- 已发出的 log toast 不回填语言（只有切语言之后的新 toast 用新语言）

#### 新增 `development/tests/test_language_switch.py`

```python
import pytest
from PySide6.QtWidgets import QApplication, QLabel
from app.ui.i18n import LanguageManager
from app.ui.main_view import MainView
from app.core.settings import Settings
from app.core.orchestrator import Orchestrator

@pytest.fixture
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app

def test_run_button_text_flips_on_language_change(qapp):
    LanguageManager.instance().set_language("zh")
    view = MainView(Settings(), Orchestrator(Settings()))
    assert view.run_btn.text() == "抓报告"
    LanguageManager.instance().set_language("en")
    assert view.run_btn.text() == "Download Reports"

def test_widget_tree_has_no_cjk_after_switching_to_en(qapp):
    LanguageManager.instance().set_language("en")
    view = MainView(Settings(), Orchestrator(Settings()))
    cjk_found: list[str] = []
    for w in view.findChildren((QLabel,)):
        text = w.text()
        if any('一' <= c <= '鿿' for c in text):
            cjk_found.append(f"{w.objectName()}: {text}")
    assert not cjk_found, f"Widgets still showing CJK in en mode: {cjk_found}"

def test_ticker_row_input_preserved_across_language_switch(qapp):
    LanguageManager.instance().set_language("zh")
    view = MainView(Settings(), Orchestrator(Settings()))
    # ... 模拟用户在某 ticker_row 输入 "600519" ...
    # (具体调用 Codex 写)
    LanguageManager.instance().set_language("en")
    # 断言该 ticker_row 的输入 == "600519"
```

### 🔴 Reviewer Checkpoint A（批次 4 后）

Codex 完成批次 1-4 后**暂停 commit 批次 5**，Claude 验：
1. **全部英文翻译质量统一审**（97 keys × en value）。Style guide 一致性、市场名全称、动词统一。
2. **11 widget 全部接信号**：`grep "_retranslate\|language_changed" app/ui/*.py` 应每个文件至少 2 处。
3. **widget 树扫描测试**：`test_widget_tree_has_no_cjk_after_switching_to_en` 必绿。
4. **TickerRow 输入保留测试**：必绿。
5. **设置 dialog 自身重开时显示新语言**：Codex 手测一轮。

Claude 通过 Checkpoint A 后，Codex 进入批次 5。

---

## Part 3 — IPO 路径统一（批次 5）

### 3.1 现状

- 5 plugin 中 US / TW / KR 已用 [app/core/output_paths.py::report_output_path_for_filing](development/app/core/output_paths.py:85)
- A 股仍有 [plugins/ashare/reports.py:315](development/plugins/ashare/reports.py:315) `_ipo_output_path` 自己实现
- HK 仍有 [plugins/hk/reports.py:466](development/plugins/hk/reports.py:466) `_filing_output_path` 自己实现

### 3.2 改动

- **删除** `plugins/ashare/reports.py:24-31` 可选 import 块 + 315 `_ipo_output_path`；调用点（line ~489）改调 `report_output_path_for_filing`
- **删除** `plugins/hk/reports.py:24-27` 可选 import 块 + 466 `_filing_output_path`；调用点（line ~723）同样改
- 如发现 ashare/hk 命名细节（如 HK 的 board 后缀）不在 helper 签名内，**给 helper 加 kwarg**而非分叉

### 3.3 验证

- `pytest tests/test_output_layout.py` 全绿
- 实跑 A 股 IPO 601127 + HK IPO 1810：文件名与改动前完全一致

---

## Part 4 — Sidecar 迁移 + 增量更新（批次 6-7）

### 4.1 批次 6 — Sidecar 落位迁移

#### 改写 [app/core/sidecar.py](development/app/core/sidecar.py)

```python
from __future__ import annotations
import json, hashlib
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator

from app.core.models import Exchange, ReportFile
from app.core.settings import load_settings


def _sidecar_root(cache_root: Path | None = None) -> Path:
    root = cache_root if cache_root is not None else load_settings().cache_path()
    return root / "sidecars"


def _stem_for(report: ReportFile) -> str:
    return (report.accession_number or Path(report.local_path).stem).strip()


def sidecar_path(report: ReportFile, cache_root: Path | None = None) -> Path:
    return (
        _sidecar_root(cache_root)
        / report.ticker.exchange.value
        / f"{_stem_for(report)}.meta.json"
    )


def write_sidecar(report: ReportFile, cache_root: Path | None = None) -> Path:
    path = sidecar_path(report, cache_root)
    path.parent.mkdir(parents=True, exist_ok=True)
    pdf = Path(report.local_path)
    sha = hashlib.sha256(pdf.read_bytes()).hexdigest() if pdf.exists() else ""
    payload = {
        "exchange": report.ticker.exchange.value,
        "ticker_code": report.ticker.code,
        "ticker_name": report.ticker.name,
        "period_year": report.period.year if report.period else None,
        "period_type": report.period.type.value if report.period else None,
        "kind": report.kind,
        "title": report.title,
        "source_url": report.source_url,
        "downloaded_at": datetime.now(timezone.utc).isoformat(),
        "file_size_bytes": report.file_size_bytes,
        "sha256": sha,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def read_sidecar(stem: str, exchange: Exchange, cache_root: Path | None = None) -> dict | None:
    path = _sidecar_root(cache_root) / exchange.value / f"{stem}.meta.json"
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def iter_sidecars(exchange: Exchange | None = None, cache_root: Path | None = None) -> Iterator[dict]:
    root = _sidecar_root(cache_root)
    glob_target = root / exchange.value if exchange else root
    if not glob_target.exists():
        return
    for path in glob_target.rglob("*.meta.json"):
        try:
            yield json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue


def migrate_legacy_sidecars(output_root: Path, cache_root: Path | None = None) -> int:
    """One-shot: move {output}/*.meta.json into cache_root/sidecars/."""
    moved = 0
    for old in output_root.rglob("*.meta.json"):
        try:
            payload = json.loads(old.read_text(encoding="utf-8"))
            exchange_code = payload.get("exchange")
            if not exchange_code:
                continue
            new_path = (
                _sidecar_root(cache_root)
                / exchange_code
                / (old.name.removesuffix(".pdf.meta.json") + ".meta.json"
                   if old.name.endswith(".pdf.meta.json") else old.name)
            )
            new_path.parent.mkdir(parents=True, exist_ok=True)
            old.rename(new_path)
            moved += 1
        except (OSError, json.JSONDecodeError) as exc:
            from loguru import logger
            logger.warning(f"sidecar migration failed for {old}: {exc}")
    return moved
```

#### 改 [app/core/orchestrator.py:78-88](development/app/core/orchestrator.py:78)

`write_sidecar(report)` 调用保持不变（默认 `cache_root=None` 走 settings 路径）。

#### 启动迁移：[app/main.py](development/app/main.py)

```python
from app.core.sidecar import migrate_legacy_sidecars
moved = migrate_legacy_sidecars(settings.output_path())
if moved:
    logger.info(f"Migrated {moved} legacy sidecars from output/ to data/cache/sidecars/")
```

#### 测试改造

- `tests/test_sidecar.py` 断言新落位 + 11 字段保留
- 新增 `tests/test_sidecar_migration.py`：tmp_path 下造 `output/` 含 3 个旧 sidecar（A 股 + HK + US），调 `migrate_legacy_sidecars`，断言：
  - 旧文件已不在 `output/`
  - 新文件在 `cache/sidecars/<exchange>/`
  - 内容完整
  - 返回值 = 3
- 边界 case：`output/` 含孤儿 sidecar（payload 无 `exchange` 字段）→ 跳过 + log warning，不抛

#### 验证

```bash
pytest tests/test_sidecar.py tests/test_sidecar_migration.py -v
```

新下载 PDF 旁 `output/` **无 `.meta.json`**；`data/cache/sidecars/<exchange>/` 含新写。

### 4.2 批次 7 — 增量更新功能

#### 新建 `development/app/core/incremental.py`

```python
from __future__ import annotations
from pathlib import Path

from app.core.models import Period, Ticker
from app.core.sidecar import iter_sidecars


def already_downloaded(ticker: Ticker, period: Period, cache_root: Path | None = None) -> bool:
    """True if a sidecar exists for (exchange, code, year, period_type)."""
    for payload in iter_sidecars(ticker.exchange, cache_root):
        if (
            payload.get("ticker_code") == ticker.code
            and payload.get("period_year") == period.year
            and payload.get("period_type") == period.type.value
        ):
            return True
    return False


def compute_incremental_pairs(
    tickers: list[Ticker],
    periods: list[Period],
    cache_root: Path | None = None,
) -> tuple[list[tuple[Ticker, Period]], int]:
    """Return (pairs_to_download, skipped_count)."""
    todo = []
    skipped = 0
    for t in tickers:
        for p in periods:
            if already_downloaded(t, p, cache_root):
                skipped += 1
            else:
                todo.append((t, p))
    return todo, skipped
```

#### UI 加按钮 [app/ui/main_view.py](development/app/ui/main_view.py)

```python
self.run_btn = QPushButton(ui_strings.MV_RUN_BUTTON)
self.run_btn.setProperty("variant", "primary")
self.run_btn.clicked.connect(self._start_job)

self.incremental_btn = QPushButton(ui_strings.MV_INCREMENTAL_BUTTON)
# 次级样式（不设 variant=primary）
self.incremental_btn.clicked.connect(lambda: self._start_job(incremental=True))
```

`_start_job` 改签名：

```python
def _start_job(self, incremental: bool = False) -> None:
    # ... 收集 tickers + periods ...
    if incremental:
        from app.core.incremental import compute_incremental_pairs
        pairs, skipped = compute_incremental_pairs(tickers, periods)
        if skipped:
            self.log(ui_strings.MV_INCREMENTAL_SKIPPED_FORMAT.format(count=skipped))
        if not pairs:
            QMessageBox.information(self, ui_strings.MV_INCREMENTAL_NONE_TITLE, ui_strings.MV_INCREMENTAL_NONE_BODY)
            return
        # 解构 pairs 为 tickers/periods 子集再走原 submit_job
```

#### strings.py 新增 keys（zh + en）

```python
"MV_INCREMENTAL_BUTTON": "增量更新" / "Incremental Update"
"MV_INCREMENTAL_SKIPPED_FORMAT": "增量模式：跳过 {count} 个已下载文件" / "Incremental: skipped {count} already-downloaded files"
"MV_INCREMENTAL_NONE_TITLE": "无新报告" / "Nothing New"
"MV_INCREMENTAL_NONE_BODY": "所选股票/期间均已下载过。" / "All selected ticker/period pairs are already downloaded."
```

#### 测试

新建 `development/tests/test_incremental.py`：

```python
def test_compute_incremental_pairs_skips_existing(tmp_path):
    # tmp_path 充当 cache_root
    # 在 cache/sidecars/A/ 下造 3 个 sidecar payload
    # 调 compute_incremental_pairs(5 tickers × 1 period)
    # 断言 todo == 2, skipped == 3
    ...

def test_compute_incremental_pairs_when_cache_empty(tmp_path):
    # 空 cache → 全部需要下载
    ...

def test_already_downloaded_matches_exchange_specific(tmp_path):
    # 不同 exchange 同 code 不应混淆
    ...
```

### 🔴 Reviewer Checkpoint B（批次 7 后）

Codex 完成批次 5-7 后**暂停**，Claude 验：
1. **Sidecar 迁移边界**：孤儿 PDF / 孤儿 sidecar / 跨 exchange 同 code 冲突全部测试覆盖
2. **增量 diff 正确**：fake cache 测试通过 + 手测一轮（先全量跑 A 股 600519 ANNUAL 2024，再增量跑同任务，应该 0 任务）
3. **IPO 统一无回归**：A 股 IPO + HK IPO 文件名与改动前一致
4. **ReportFile schema 锁定**：JP/UK 接下来不再改 schema
5. **`output/` 100% 干净**：实跑后只剩 PDF

Claude 通过 Checkpoint B 后，Codex 进入批次 8。

---

## Part 5 — 新市场（批次 8-9）

### 5.1 批次 8 — Japan plugin (EDINET，双模式)

#### 5.1.1 第 1 天：API spike（非 commit 代码）

在 `development/tmp_pytest_ok/` 写一次性脚本试 EDINET 端点：
- `https://disclosure2.edinet-fsa.go.jp/api/v2/documents.json?date=2024-06-28&type=2`（按日期列表）
- 拉某文档 PDF：`/api/v2/documents/{docID}?type=2`（PDF）
- 公网搜索 fallback：`https://disclosure2.edinet-fsa.go.jp/weee0010.aspx`

确认：
- API Key 是否真必需（v2 部分免 Key 端点）
- captcha / session 限制
- rate limit（保守 2-3 req/s）
- "有価証券報告書"（年报）的 docTypeCode = `120`
- 公司搜索关键字：tickerSymbol or `JCN`（法人番号）or 名称

发现阻塞性问题立即报 Claude 重评 scope。

#### 5.1.2 实施

**新建 `development/plugins/jp/`**：

- `__init__.py`：`JPShare(ExchangePlugin)` 主类，懒 import name_resolver + reports
- `name_resolver.py`：ticker code（如 `7203`）→ EDINET docID / 法人番号；用 `cached_or_load(_CACHE_KEY, _fetch, expire=24h)`
- `reports.py`：按 period 拉 EDINET 文档列表 + 下载 PDF；调 `report_output_path` / `report_output_path_for_filing`
- `edinet_api.py`：双模式之"有 Key 走 API"分支（`Subscription-Key` header）
- `edinet_web.py`：双模式之"无 Key 公网兜底"分支（仿 KR `dart_web.py`，含 `_SELECTORS` 字典）

**核心契约**：
- 输出 DataFrame 列名对齐：`doc_id` / `doc_type_code` / `submit_date_time` / `period_end`（用作 period_year 推断）
- 双模式分支在 reports.py 用 `_edinet_api_key() or None` 判断（仿 KR `_dart()`）

#### 5.1.3 配套改动

- [app/core/models.py](development/app/core/models.py)：加 `Exchange.JP = "JP"` + `display_name` dict（"日股 / Japan"）
- [plugins/__init__.py](development/plugins/__init__.py)：`get_plugin` 加 JP 分支
- [app/core/settings.py](development/app/core/settings.py)：加 `EDINETCfg(api_key: str = "")` + `RateLimitsCfg.edinet: float = 2.0`
- [app/ui/settings_dialog.py](development/app/ui/settings_dialog.py)：加 EDINET Key 输入框（仿 DART line 31-34）
- UI strings（zh + en 双份）：
  - `ES_NAME_JP = "日股 / Japan"` & `"Japan"`
  - `ES_META_JP = "EDINET · TSE"`
  - `EP_TITLE_JP` / `EP_SUBTITLE_JP`
  - `TR_PLACEHOLDER_JP = "如 7203 / 6758"` & `"e.g. 7203 / 6758"`
- [app/ui/main_view.py:87](development/app/ui/main_view.py:87) tuple 加 `Exchange.JP`
- [app/ui/exchange_selector.py:67-80](development/app/ui/exchange_selector.py:67) `_zh_name` / `_meta` 加 JP 分支

#### 5.1.4 测试

新建 `development/tests/test_jp_name_resolver.py` + `test_jp_reports.py`：
- 至少 5 个测试覆盖 resolve / fetch_company / list_filings / download_pdf / 双模式分支
- 用 mock httpx 响应（fixture HTML / JSON 保存在 `tests/fixtures/edinet_*.json`）

#### 5.1.5 Smoke（用户实测时）

| 票 | 期间 | 验证 |
|---|---|---|
| 7203 Toyota | ANNUAL 2024 | 有価証券報告書 |
| 6758 Sony | ANNUAL 2024 | 同 |
| 9984 SoftBank Group | ANNUAL 2024 | 同 |

文件名：`JP_7203_Toyota_2024_年报.pdf`（中文 kind 后缀保留）

### 5.2 批次 9 — UK plugin (NSM)

#### 5.2.1 第 1 天 API spike

试 NSM 端点：
- `https://data.fca.org.uk/#/nsm/nationalstoragemechanism`（搜索 UI）
- 实际 JSON 接口：`https://api.data.fca.org.uk/services/nsm/search-results`（具体待 Codex 实地确认）
- 文档详情 + PDF URL 提取方式

公司代码格式 normalize：`ULVR.L` → `ULVR` 或两者都接受。

#### 5.2.2 实施

**新建 `development/plugins/uk/`**：
- `__init__.py`：`UKShare(ExchangePlugin)`
- `name_resolver.py`：ticker（如 `ULVR` 或 `ULVR.L`）→ LEI 号或 NSM 内部 ID
- `reports.py`：按 period 拉 NSM 文档；选片关键字 `"Annual Report"` / `"Annual Financial Report"` / `"Half-year Report"` / `"Q1 Trading Update"`
- `nsm_web.py`：公网客户端（含 `_SELECTORS`）

UK 不需要 `*_api.py`（NSM 无 Key），但保留 `nsm_web.py` 命名与 JP/KR 一致。

#### 5.2.3 配套改动（同 JP）

- `Exchange.UK = "UK"`，display name "英股 / United Kingdom"
- `plugins/__init__.py` 加 UK 分支
- `RateLimitsCfg.nsm: float = 2.0`
- UI strings：`ES_META_UK = "LSE · NSM"`、`TR_PLACEHOLDER_UK = "如 ULVR / HSBA"`
- `main_view.py:87` 加 UK
- `exchange_selector.py` 加 UK 分支

#### 5.2.4 测试 + Smoke

| 票 | 期间 | 验证 |
|---|---|---|
| ULVR Unilever | ANNUAL 2024 | Annual Report PDF |
| HSBA HSBC | ANNUAL 2024 | Annual Report PDF |
| AZN AstraZeneca | ANNUAL 2024 | Annual Report PDF |

文件名：`UK_ULVR_Unilever_2024_年报.pdf`

### 🔴 Reviewer Checkpoint C（批次 9 后）

Codex 完成批次 8-9 后**暂停**，Claude 验：
1. **JP/UK 与 KR 双模式 shape 对齐**：reports.py 中 `_dart() / _edinet_api_key()` 模式一致；公网模块 `_SELECTORS` 集中
2. **settings 流程一致**：EDINET Key 输入与 DART 同样 placeholder / save 后重置 client cache
3. **6 票实跑出真 PDF + sidecar**：3 JP + 3 UK，全部 `%PDF` 头
4. **两种语言下 7 市场无 UI 回归**：切 en + 切 zh 各一遍
5. **strings.py 双语 key 对齐**：JP/UK 新加的 6 个 key（每个市场 ES_NAME / ES_META / EP_TITLE）zh+en 完整

Claude 通过 Checkpoint C 后，Codex 进入批次 10。

---

## Part 6 — 文档 + 发布（批次 10-11）

### 6.1 批次 10 — "如何加新市场" 文档

#### 追加章节到 [ARCHITECTURE.md](ARCHITECTURE.md)

新章节 "## 11. 如何加新市场（v1.0 起 7 市场，扩展模板）" 含 6 节：

1. **Exchange 枚举**：`app/core/models.py::Exchange` 加新成员 + `display_name` dict
2. **Plugin 三件套**：
   - 基础：`plugins/{mk}/__init__.py` + `name_resolver.py` + `reports.py`
   - 双模式可选：拆 `{mk}_api.py`（有 Key）+ `{mk}_web.py`（公网）
   - 用 JP plugin 嵌实际代码段
3. **注册**：`plugins/__init__.py::get_plugin` 加 if 分支
4. **UI 接入**：`main_view.py` tuple + `exchange_selector.py` `_zh_name`/`_meta` + `strings.py` 6 个 key（×2 lang）
5. **Settings 可选 Key**：模型加 `{Market}Cfg(api_key: str = "")` + `settings_dialog.py` 加输入框
6. **测试要求**：`tests/test_{mk}_name_resolver.py` + `tests/test_{mk}_reports.py` 至少 5 用例 + fixture 文件

Claude 人工审可读性（不写测试）。

### 6.2 批次 11 — 发布打磨 + v1.0 tag

#### 版本号 + Spec

- [pyproject.toml](development/pyproject.toml)：`version = "1.0.0"`
- `filings_atlas.spec` version 元数据同步

#### README 双语版

完全重写 `README.md`，结构：

```markdown
# Filings Atlas / 全球披露图谱

[EN](#english) | [中文](#chinese)

## English

One-click multi-market disclosure PDF downloader. Covers 7 markets:
- A-Share (CNINFO + akshare)
- Hong Kong (HKEXnews + Eastmoney)
- United States (SEC EDGAR)
- Korea (DART, dual-mode with optional API key)
- Taiwan (TWSE + MOPS)
- Japan (EDINET, dual-mode)
- United Kingdom (NSM/FCA)

### Quick start
... screenshots (English UI) ...

## 中文

跨市场上市公司官方披露 PDF 一键下载工具，覆盖 7 大市场...

### 快速开始
... 截图（中文 UI） ...
```

截图要求：英文 UI 截图 + 中文 UI 截图各 1 张，主视图含 7 个 chip。

#### CHANGELOG.md

新建，含 v0.6 → v1.0 累积变更（按用户视角分组：新增功能 / Bug 修复 / 性能 / 重大变更 "Breaking changes — 重命名"）。

#### LICENSE

复核（应该是 MIT 或其他），新增 v1.0 年份。

#### `.github/workflows/release.yml`

```yaml
name: Release
on:
  push:
    tags: ['v*.*.*']

jobs:
  build-release:
    runs-on: windows-latest
    environment: production  # 手动审批门禁
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'
      - name: Install
        run: |
          cd development
          pip install -e .
          pip install pyinstaller
      - name: Build
        run: |
          cd development
          pyinstaller --noconfirm filings_atlas.spec
      - name: Package
        run: |
          Compress-Archive -Path "Filings Atlas.exe","_internal" -DestinationPath "FilingsAtlas-${{ github.ref_name }}-windows.zip"
      - name: Release
        uses: softprops/action-gh-release@v2
        with:
          files: FilingsAtlas-${{ github.ref_name }}-windows.zip
          body_path: CHANGELOG.md
```

`environment: production` 需 GitHub 仓库设置中手动配 approval reviewer。

#### PyInstaller smoke

```powershell
cd development
pyinstaller --noconfirm filings_atlas.spec
cd ..
.\"Filings Atlas.exe"
# 启动 + 跑 A 股 600519 ANNUAL 2024
# 切 English → 跑 US AAPL 10-K 2024
```

#### Git tag

```bash
git tag v1.0.0
git push origin v1.0.0  # 触发 release workflow
```

---

## 实施顺序（Codex 分批 commit）

| 批次 | 主题 | 验证 | Checkpoint |
|---|---|---|---|
| **1** | 品牌重命名 | `grep` 清单 + 启动 EXE | — |
| **2** | i18n 脚手架 + `LanguageManager` + settings migration | `pytest tests/test_i18n.py -v` 绿 | — |
| **3** | `strings.py` 双语拆分 + key parity 测试 | `pytest tests/test_ui_strings.py -v` 绿 | — |
| **4** | 11 widget _retranslate + 语言开关 + tree scan 测试 | `pytest tests/test_language_switch.py -v` 绿；启动 EXE 切语言肉眼 | 🔴 **A** |
| **5** | IPO 路径统一（ashare + hk） | `pytest tests/test_output_layout.py -v` 绿；A 股 IPO + HK IPO 实跑 | — |
| **6** | Sidecar 迁移到 `data/cache/sidecars/` + 启动 migrate + iter/read helpers | `pytest tests/test_sidecar.py tests/test_sidecar_migration.py -v` 绿 | — |
| **7** | 增量更新 + UI 按钮 + tests | `pytest tests/test_incremental.py -v` 绿；实跑全量后增量应 0 任务 | 🔴 **B** |
| **8** | JP plugin (EDINET 双模式) | API spike + 3 票 smoke + tests | — |
| **9** | UK plugin (NSM) | API spike + 3 票 smoke + tests | 🔴 **C** |
| **10** | "如何加新市场" 文档章节 | Claude 人工审 | — |
| **11** | 版本号 + README 双语 + CHANGELOG + workflow + smoke build + tag | clean build + 双语 EXE 实跑 + GitHub release artifact | — |

commit message 格式：`v1.0: <一句话变更>（批次 N）`

3 个 Reviewer Checkpoint 必须按顺序通过，任一失败回到上一个 commit fix。

---

## 测试矩阵

### 单元 + 集成（每批必绿）

```bash
cd development
pytest -m "not e2e" -v
ruff check .
```

v0.8 基线 **100 passed**，v1.0 预计 **~120 passed**：
- batch 2：+2（test_i18n）
- batch 3：+3（key parity / en no CJK / zh has CJK）
- batch 4：+3（test_language_switch）
- batch 6：+2（test_sidecar_migration + 改 test_sidecar 一致）
- batch 7：+3（test_incremental）
- batch 8：+5（test_jp_name_resolver + test_jp_reports）
- batch 9：+5（test_uk_name_resolver + test_uk_reports）

### e2e Smoke（用户实测）

```powershell
$env:FS_CAPTURE_RUN_E2E="1"
cd development
pytest tests/e2e -v
```

7 市场全绿：A / HK / US / KR / TW（v0.8 基线 5）+ JP + UK（v1.0 新增 2）。

### 双语手动验证

1. 启动 EXE，默认中文 UI
2. 设置 → 切 English → **不重启**，所有 UI 文字翻转
3. TickerRow 已输入的代码保留不丢
4. 抓一票 EDINET PDF，log 提示用英文
5. 重启 EXE 仍是 English（设置持久化）
6. 切回 中文，重启 → 中文

### Sidecar 迁移验证

1. 启动前在 `output/` 手造 1 个旧 `*.meta.json`（payload 含 exchange）
2. 启动 EXE，log 出现 "Migrated 1 sidecar to data/cache/sidecars/"
3. `output/` 中**只剩 PDF**
4. `data/cache/sidecars/<exchange>/` 含新落位

### 增量按钮验证

1. 全量跑 A 股 600519 ANNUAL 2024（落地 1 PDF + 1 sidecar）
2. 点"增量更新" + 同参数 → 弹窗"无新报告"
3. 加一只 600520 ANNUAL 2024 → 增量跑 → 只下 600520

### 发布包验证

```powershell
cd development
pyinstaller --noconfirm filings_atlas.spec
cd ..
.\"Filings Atlas.exe"
```

- 7 chips 渲染（A / HK / US / KR / TW / JP / UK）
- 中英切换流畅
- 7 市场各跑 1 票真实抓取

---

## Reviewer Checklist（最终验收）

### 重命名
- [ ] `grep -ri "fs capture" development/app development/plugins` 仅剩有意保留
- [ ] EXE 启动窗口标题 = "Filings Atlas / 全球披露图谱"
- [ ] `pyproject.toml` `name = "filings-atlas"` + console script 同步
- [ ] `filings_atlas.spec` 替换 `fs_capture.spec`，build.bat 同步
- [ ] 图标资源已 git mv 重命名
- [ ] log file 写到 `filings_atlas.log`
- [ ] `VBA Captor/` 引用零改动

### 双语
- [ ] `STRINGS["zh"]` 和 `STRINGS["en"]` key 完全对齐
- [ ] `STRINGS["en"]` value 无 CJK
- [ ] 11 widget 全部 `_retranslate` 接信号
- [ ] 切语言 widget 树扫描无 CJK
- [ ] TickerRow 已输入 code 在切语言后保留
- [ ] 设置 dialog 自身重开也响应语言
- [ ] 设置持久化（重启 EXE 保持上次语言）

### Sidecar 迁移
- [ ] 新下载 PDF 旁 `output/` 无 `.meta.json`
- [ ] sidecar 落位 `data/cache/sidecars/<exchange>/<stem>.meta.json`
- [ ] 旧 sidecar 启动时自动迁移 + log
- [ ] `test_sidecar_migration` 覆盖孤儿 + 跨 exchange 冲突
- [ ] 11 字段完整保留

### 增量
- [ ] "增量更新 / Incremental" 按钮在主视图，次级样式
- [ ] 增量跑应跳过已 sidecar 的 (ticker, period_year, period_type)
- [ ] 跨 exchange 同 code 不混淆
- [ ] 全增量场景弹"无新报告"对话框

### IPO 统一
- [ ] `ashare/reports.py:_ipo_output_path` 已删
- [ ] `hk/reports.py:_filing_output_path` 已删
- [ ] 两市场 IPO 文件名与改动前一致
- [ ] `test_output_layout.py` 全绿

### 新市场
- [ ] `Exchange.JP` / `Exchange.UK` 已加，`display_name` 同步
- [ ] `plugins/__init__.py::get_plugin` 含 JP / UK 分支
- [ ] JP 双模式：有 Key 走 EDINET API，无 Key 走公网
- [ ] UK 公网模式：NSM 端点稳定
- [ ] JP 3 票实跑（7203 / 6758 / 9984）+ UK 3 票实跑（ULVR / HSBA / AZN）
- [ ] `_SELECTORS` 字典集中在 `edinet_web.py` / `nsm_web.py`
- [ ] settings_dialog EDINET Key 输入框（仿 DART）

### 文档
- [ ] `ARCHITECTURE.md` 含"如何加新市场"章节 + JP plugin 工作示例
- [ ] `PROJECT_RETROSPECTIVE.md §13 v1.0 postscript`
- [ ] `README.md` 双语版（EN + 中文段落 + 7 市场表 + 双语截图）
- [ ] `CHANGELOG.md` 累积 v0.6 → v1.0
- [ ] `CLAUDE.md` ≡ `AGENTS.md` v1.0 进度同步

### 发布
- [ ] `pyproject.toml` version = "1.0.0"
- [ ] `.github/workflows/release.yml` 配 production environment
- [ ] PyInstaller clean build 通过 + EXE 实跑 7 市场各 1 票
- [ ] `git tag v1.0.0` push 触发 release workflow
- [ ] GitHub release artifact = `FilingsAtlas-v1.0.0-windows.zip` 含 exe + `_internal/`
- [ ] release body = CHANGELOG.md

### 质量门禁
- [ ] `pytest -m "not e2e" -v` ≥ **120 passed**
- [ ] `ruff check development/` `All checks passed!`
- [ ] CI workflow 通过（lint + test job）
- [ ] e2e 7 市场全绿

---

## 风险与缓解

| 风险 | 影响 | 缓解 |
|---|---|---|
| **97 keys × en 翻译质量** | UI 文案不专业 | Checkpoint A Claude 一次性审 + style guide（"Download" 不要 "Fetch" / "Grab"，市场名全称不缩写） |
| **EDINET / NSM 端点逆向超时** | 批次 8/9 卡住 | 各批首日 API spike（throwaway code），发现阻塞立即报 Claude 重评 scope |
| **Sidecar 迁移 bug 毁用户已下载状态** | 用户体验差 | `test_sidecar_migration` 覆盖孤儿 PDF / 孤儿 sidecar / 跨 exchange 冲突；首启逐条 log |
| **Live re-translate 漏 widget** | 半中半英 | `test_widget_tree_has_no_cjk_after_switching_to_en` 走 `findChildren(QLabel)` 全扫描 |
| **PyInstaller 重命名破坏现有用户** | 旧快捷方式失效 | 批次 11 全 clean rebuild + EXE 7 市场实跑；CHANGELOG 显式说明 "Breaking: renamed from FS Capture" |
| **Reviewer 3 个 Checkpoint 之间 Codex 等待时间** | sprint 总耗时膨胀 | 用户告知 Claude 何时跑 Checkpoint，目标每个 Checkpoint 用户响应 ≤ 24h |
| **GitHub release workflow 首次跑失败** | tag 已 push 无法回滚 | 先在 fork / branch 上 dry-run；workflow 用 `environment: production` 加手动 approval |

---

## 不在 v1.0 范围（明确排除）

- 日语 / 韩语 UI 字符串 — 仅 zh + en，dict 结构支持但翻译留 v1.1+
- 其他新市场（SGX / ASX / 印度 / SEDAR / Bundesanzeiger / AMF）— 每个独立 1 周 mini-sprint，v1.1+
- EDINET / NSM 全文搜索 UX
- 增量更新粒度超过 (ticker, period_year, period_type)
- Sidecar schema v2 / 破坏性字段改动
- 已下载 PDF 浏览 GUI
- `config.toml` 通用迁移框架（language `zh_CN→zh` 仅一行 if）
- 跨平台打包（macOS / Linux）— Windows-only PyInstaller for v1.0
- 本地化文件名（`_年报.pdf` 在 en UI 下也保留）
- Plugin 自动发现 — `plugins/__init__.py` 维持显式 if-chain

---

## 关键文件路径（Codex 速查）

### 新建
- `development/app/ui/i18n.py`
- `development/app/core/incremental.py`
- `development/plugins/jp/__init__.py` + `name_resolver.py` + `reports.py` + `edinet_api.py` + `edinet_web.py`
- `development/plugins/uk/__init__.py` + `name_resolver.py` + `reports.py` + `nsm_web.py`
- `development/tests/test_i18n.py` / `test_language_switch.py` / `test_incremental.py` / `test_sidecar_migration.py` / `test_jp_*.py` (×2) / `test_uk_*.py` (×2)
- `development/tests/fixtures/edinet_*.json` / `nsm_*.json`（实地 spike 后保存）
- `CHANGELOG.md`、`.github/workflows/release.yml`
- 资源重命名：`filings_atlas.ico` / `filings_atlas_logo.png`、`filings_atlas.spec`

### 重写
- `development/app/ui/strings.py`（97 keys × 2 lang）
- `development/app/core/sidecar.py`（write/read/iter + migrate）
- `development/app/core/models.py::Exchange`（+JP +UK）
- `development/plugins/__init__.py::get_plugin`（+JP +UK 分支）
- `README.md`（双语 + 7 市场表 + 双语截图）

### 修改（行号已确认 2026-05-23）
- `app/main.py`:47,62-65,75,104,110
- `app/ui/main_window.py`:36,89
- `app/core/settings.py`:42,51 + 新 EDINETCfg + 新 RateLimitsCfg.edinet/nsm
- `app/ui/settings_dialog.py`:42（+ EDINET Key 输入框 + 语言 combo）
- `app/ui/main_view.py`:87（tuple +JP +UK + 增量按钮 + _start_job 改签名）
- `app/ui/exchange_selector.py`:67-80（+JP +UK 分支）
- `app/core/orchestrator.py`:78-88（write_sidecar 调用保持，no behavior change）
- `plugins/ashare/reports.py`:24-31, 315（删 IPO helper）
- `plugins/hk/reports.py`:24-27, 466（删 filing helper）
- 11 UI 文件各加 `_retranslate`
- `ARCHITECTURE.md`（v1.0 章节 + §11 "如何加新市场"）
- `PROJECT_RETROSPECTIVE.md`（§13 v1.0 postscript）
- `CLAUDE.md` + `AGENTS.md` 同步（§10 v1.0 进度 + §3 文档地图）

### 不动
- `app/core/http.py`、`ratelimit.py`、`cache.py`、`pdf_renderer.py`、`output_paths.py::report_output_path`（其他函数）
- 5 现有 plugin 的 `name_resolver` / `reports`（除 ashare/hk IPO helper 拆除 + 加 `cached_or_load` 已在 v0.8 完成）
- KR `dart_web.py` / HK `fiscal_year.py` / `_pdf_verify.py` / TW reports.py 业务级重试
- `VBA Captor/` 子项目

---

## 致 Codex 的提示

1. **依次做批次**：1 → 2 → 3 → 4 → **暂停等 Checkpoint A**；5 → 6 → 7 → **暂停等 Checkpoint B**；8 → 9 → **暂停等 Checkpoint C**；10 → 11。
2. **批次 1 是纯机械重命名**：先把它做完得到清单，再启动 i18n 工作避免品牌名硬编码进 strings.py 后还要改两次。
3. **批次 3 的英文翻译**：97 keys 是大体力活，**逐 key 翻**，参考 style guide；翻完整体扫一遍统一性，再提 commit。Checkpoint A Claude 会全审。
4. **批次 6 sidecar migration 的孤儿场景**：必须容错 — 旧 sidecar payload 不含 `exchange` 字段时跳过 + log warning（不抛）。
5. **批次 8/9 第 1 天 API spike**：spike 代码放 `development/tmp_pytest_ok/`（gitignored），结论写到 commit message 或 `docs/plans/2026-05-23-{jp,uk}-api-spike-notes.md`。**spike 不要进 commit**。
6. **批次 11 git tag**：tag push 前先在本地完成 PyInstaller smoke + 7 市场各 1 票实跑。tag 推上去就触发 release workflow，**别一边修代码一边 tag**。
7. **commit message**：`v1.0: <一句话变更>（批次 N）`。每批一个 commit，不合并。

---

**最后更新**：2026-05-23（批次 3 strings.py 双语拆分完成）

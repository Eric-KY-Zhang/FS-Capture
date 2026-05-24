# SPRINT v1.0 — SG 新加坡市场 + 抓取性能优化 + 体积压缩

**日期**：2026-05-24
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex
**Reviewer**：Claude Code
**状态**：待 Codex 启动
**预计工作量**：3-4 周
**发布策略**：**首个 GitHub release**（v1.0.0 tag + artifact）

---

## Context

Filings Atlas / 全球披露图谱 已完成 7 市场（A/HK/US/KR/TW/JP/UK）+ 双语 UI + sidecar 迁移 + 增量更新 + 品牌升级 + 主界面语言切换。该批工作**原计划作为 v1.0 首发 GitHub release**，但用户 2026-05-24 决定**版本号重定**：

- **已完成工作 → v0.9**（内部迭代，不发 release）
- **本 sprint → v1.0**（首发 GitHub release）

新 v1.0 核心目标（用户决策 2026-05-24）：

1. **新增第 8 个市场：新加坡 SGX**（覆盖年报 + 中期 + IPO）
2. **抓取性能优化**（首要诉求 — "速度最重要"）
3. **视情况压缩软件体积**（次要 — "只能在不影响性能及用户体验的情况下"）

### 决策表

| 决策项 | 选定方案 |
|---|---|
| 版本号 | 已完成工作 v1.0 → **v0.9**；本 sprint = **v1.0**（首发 release） |
| Sprint 划分 | **单个 v1.0 mega sprint**（7 批次 + 3 Reviewer Checkpoint） |
| SG 数据源 | **SGXNet announcements**（公网 JSON API + HTML 双解析，无 API Key） |
| SG ticker 规范化 | **大写 + 去 `.SI` 后缀**（`D05` / `d05.si` / `D05.SI` 都接受 → `D05`） |
| SG plugin 模式 | **单模式公网爬虫**（同 UK，不需 Key，**不需 Playwright**） |
| SG period | ANNUAL（年报）+ H1（中期）+ IPO_PROSPECTUS（招股说明书） |
| 性能优化哲学 | **基线测量 → 单项变更 → A/B 验证 → 锁定**，每项可独立回滚 |
| 体积优化范围 | 仅"零体验代价"项；明确**不动** Chromium / akshare / pandas |
| Reviewer Checkpoint | 3 个（A: SG plugin + smoke；B: 性能 A/B 数据；C: 体积 + 8 市场回归） |

### 明确不做（用户已锁定）

- ❌ Edge WebView2 替代 Chromium（影响 KR/JP/UK 体验）
- ❌ 自维护 A 股代码表替代 akshare（年维护成本 + 新股 IPO 延迟）
- ❌ 删除 OpenDartReader 走纯公网 fallback（v0.7 已做双模式，无需破坏）
- ❌ 跨平台打包（macOS / Linux）
- ❌ 多 part 并发下载单 PDF（HTTP Range 分块）
- ❌ benchmark 自动化 CI（手动跑即可）
- ❌ 第三语言 UI

---

## 批次概览

| 批次 | 名称 | 工作量 | 输出 |
|---|---|---|---|
| 0 | 版本号大调整（文档 + tag） | 1 天 | 文档同步、文件 git mv、`pyproject.toml` 0.9.0、git tag 重打 |
| 1 | SG plugin API spike | 1 天 | `tmp_pytest_ok/sg_spike.py` + `docs/plans/2026-05-25-sg-spike-report.md` |
| 2 | SG plugin 三件套实现 | 3-4 天 | `plugins/sg/*.py` + 单测 |
| 3 | SG UI 接入 + 双语字符串 + smoke | 2 天 | `app/ui/*.py` + `strings.py` + 4 票 smoke |
| **🔴 A** | **Reviewer Checkpoint A** | 0.5 天 | Claude 验收 |
| 4 | 性能基线 benchmark | 1.5 天 | `tests/benchmarks/test_baseline_perf.py` + `docs/perf/v0.9_baseline.md` |
| 5 | 抓取性能优化（6 子项） | 4-5 天 | `app/core/*.py` 改动 + A/B 数据 |
| 6 | 体积压缩（5 子项） | 2 天 | `filings_atlas.spec` 调整 + `docs/perf/v1.0_size.md` |
| **🔴 B** | **Reviewer Checkpoint B** | 0.5 天 | Claude 验收性能 + 体积 |
| 7 | v1.0 发布打磨 + tag | 1.5 天 | `pyproject.toml` 1.0.0 + `CHANGELOG.md` + 截图 + git tag |
| **🔴 C** | **Reviewer Checkpoint C** | 0.5 天 | Claude 最终验收 |

**总计**：~17 工作日 = 3-4 周

---

## 批次 0 — 版本号大调整（文档 + tag，零功能变更）

### 0.1 文档同步（"v1.0" → "v0.9"）

#### CLAUDE.md / AGENTS.md 改动（必须双文件同步）

**§ 10 当前版本进度表**：
- 当前 v1.0 行 "🟡 Release Candidate" → 改为 v0.9 行 "✅ 已完成 2026-05-23"
  - 描述改为：7 市场（含 JP/UK）+ 双语 UI + sidecar 迁移 + 增量更新 + 品牌升级
- **新增 v1.0 行**：状态 "🟡 待启动"
  - 计划文档：`roadmap/SPRINT_v1.0_sg_and_perf.md`
  - 描述：SG 新加坡市场 + 抓取性能优化 + 视情况压缩体积（**首发 GitHub release**）

**§ 11 路线图速览**：
- 原"v1.0（4-6 周）"段拆分为两段：
  - **v0.9（已完成）**：原 v1.0 11 批次内容（重命名 + i18n + sidecar + 增量 + JP/UK）—— 改 prefix 表述
  - **v1.0（3-4 周，7 批次 + 3 Checkpoint）**：本 sprint 内容（SG + 性能 + 体积）
- 末尾"明确不做"列表保留并合并

**§ 13 紧急情况下方** "最后更新" 日期：改为 2026-05-24，说明改为"v1.0 sprint 启动（SG + 性能）"

#### PROJECT_RETROSPECTIVE.md

- § 13（如果存在 v1.0 postscript）改名为 § 13 v0.9 postscript
- 内容保留（描述 v0.9 实施情况），但 header 改为"v0.9"

#### README.md

- 顶部如有"v1.0" / "Release v1.0" 字样：审查并删除（v1.0 尚未发布）
- 保留产品名 "Filings Atlas / 全球披露图谱"（与版本号解耦）
- "What's new" 段若提到 v1.0：暂时删除或改为"v1.0 即将发布"

#### CHANGELOG.md（如已存在）

- 已有的 v1.0 段改为 v0.9 段
- 新建空 v1.0 段（最后批次 7 填写）

### 0.2 文件重命名（git mv 保留 history）

```bash
cd "E:\Claude+CODEX Project\FS Capture"
git mv roadmap/SPRINT_v1.0_filings_atlas.md roadmap/SPRINT_v0.9_filings_atlas.md
git mv roadmap/ROADMAP_v0.6.1_to_v1.0.md roadmap/ROADMAP_v0.6.1_to_v0.9.md
```

随后**编辑文件内部措辞**：
- `SPRINT_v0.9_filings_atlas.md` 第一行 header `# SPRINT v1.0 ...` → `# SPRINT v0.9 ...`
- 内部所有"v1.0"叙述改"v0.9"
- `ROADMAP_v0.6.1_to_v0.9.md` 同上，顶部说明改为"v0.6.1 → v0.9"

### 0.3 版本数字

| 文件 | 当前 | 改为 |
|---|---|---|
| `development/pyproject.toml:3` | `version = "1.0.0"` | `version = "0.9.0"` |
| `development/version_info.txt` | （如有 1.0.0）| 0.9.0 |
| `development/filings_atlas.spec` | （如有 version 元数据）| 0.9.0 |

**重要**：批次 7 最后再升回 1.0.0。

### 0.4 Git tag 调整（✅ 已由 Planner 完成，Codex 跳过）

**完成时间**：2026-05-24（Planner 验证用户授权后执行）

**用户确认**：origin 上未推送 v1.0 / v1.0.0 tag（`git ls-remote --tags origin` 输出为空），属非破坏性本地操作。

**已执行**：
```bash
git tag -d v1.0      # was b80298f
git tag -d v1.0.0    # was 8b4d882
git tag v0.9.0 4832479  # 指向"v1.0: UI 打磨 — 主界面语言切换 + 全球披露图谱品牌升级"
```

**当前 `git tag -l`**：`v0.1.0 / v0.5 / v0.6 / v0.9.0`

**Codex 验证**：批次 0 启动时 `git tag -l` 应显示上述 4 个 tag，不含 v1.0 / v1.0.0；如有出入立刻停并联系用户。

### 0.5 验证

```bash
cd "E:\Claude+CODEX Project\FS Capture"

# A. 文档无 v1.0 残留（除明确指向本 sprint 的）
grep -rn "v1\.0" development docs roadmap README.md CLAUDE.md AGENTS.md PROJECT_RETROSPECTIVE.md | grep -v "本 sprint\|即将发布\|SPRINT_v1.0_sg_and_perf"

# B. 单元测试不退化
cd development && pytest -m "not e2e" -v  # 维持 145/145

# C. ruff 干净
ruff check .

# D. 版本号
python -c "import tomllib; print(tomllib.loads(open('pyproject.toml','rb').read().decode())['project']['version'])"  # 输出 0.9.0
```

### 0.6 Commit

```
v1.0: 版本号大调整 — 原 v1.0 降级为 v0.9（批次 0）

- pyproject.toml version 1.0.0 → 0.9.0
- 文档同步：CLAUDE/AGENTS § 10-11、PROJECT_RETROSPECTIVE § 13、README、CHANGELOG
- 文件重命名：SPRINT_v1.0_filings_atlas → v0.9，ROADMAP v1.0 → v0.9
- 本地 git tag v1.0 / v1.0.0 删除，重打 v0.9.0 指向 4832479
- 145 tests 全绿，ruff 0 warning
```

### Codex 自检 checklist 批次 0

- [x] ~~用户已授权 git tag 删除（对话中明确）~~ — 已由 Planner 完成
- [x] ~~`git tag -l` 显示 v0.9.0 存在，v1.0 / v1.0.0 不存在~~ — 已确认（v0.1.0 / v0.5 / v0.6 / v0.9.0）
- [ ] `pyproject.toml` version = "0.9.0"
- [ ] `grep -rn "v1\.0" ...` 输出仅含本 sprint 引用
- [ ] `pytest -m "not e2e"` 145/145
- [ ] `ruff check .` 0 warning
- [ ] CLAUDE.md 与 AGENTS.md 内容逐字一致

---

## 批次 1 — SG plugin API spike（先验证再实现）

### 1.1 目的

在批次 2-3 大规模铺设代码前，先用一次性脚本验证 SGXNet 端点真实可用。spike 失败立刻停（回 Planner 重设计），不浪费精力实施 plugin。

### 1.2 实施位置

`tmp_pytest_ok/sg_spike.py`（一次性脚本，commit 后保留 1 周参考用）

### 1.3 Spike 内容

#### 1.3.1 SGXNet announcements 端点验证

候选端点（按可信度排序）：
1. `https://api.sgx.com/announcements/v1.0/search` — 官方 JSON API
2. `https://www.sgx.com/securities/company-announcements` — HTML 列表页（兜底）
3. `https://links.sgx.com/...` — 个别 announcement 详情页

测试方式：
```python
# tmp_pytest_ok/sg_spike.py
import httpx

with httpx.Client(headers={
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "application/json",
    "Referer": "https://www.sgx.com/securities/company-announcements",
}) as client:
    # 测 DBS（D05）2024 年报
    resp = client.get("https://api.sgx.com/announcements/v1.0/search", params={
        "params": "symbol=D05&announcement_type=Annual Report",
        "date_range": "2024-01-01,2024-12-31",
    })
    print(resp.status_code, resp.headers.get("content-type"))
    print(resp.text[:2000])
```

**期望产出**：找到一个稳定 JSON 端点，可按 ticker + announcement_type + date_range 过滤。

#### 1.3.2 三类 period 数据可见性

各测试一票：
- **ANNUAL** — DBS Group (`D05`) 2024 Annual Report
- **H1** — UOB (`U11`) 2024 H1/Interim Report
- **IPO** — 任意一个 2023-2024 SG IPO（如有，spike 时确定）

每类记录：
- announcement_type 字段的精确字面值（可能是 `"Annual Report"` 或 `"Annual Reports"` 或 `"Annual_Report"`）
- PDF 直链是否在 JSON 返回中（vs 需二次 GET HTML 解析）
- PDF 是否 anonymous 可下载（无 cookie / 无 referer）

#### 1.3.3 反爬阶段验证

- 连续请求 10 次（间隔 0.5s），观察是否 429 / 403
- 缺 User-Agent / Referer 是否拒绝
- 是否有 cookie/session 强制
- 是否需要 Playwright（SGXNet 改 SPA 化）

#### 1.3.4 Ticker 格式

- 测试 `D05` / `d05` / `D05.SI` / `d05.si` 输入是否都能命中
- 确定规范化规则（plan 已定：大写 + 去 `.SI` 后缀）
- 是否需要查询时附加交易所代码（如 `SGX:D05`）

### 1.4 Spike 报告输出

写到 `docs/plans/2026-05-25-sg-spike-report.md`，包含：

```markdown
# SG SGXNet Spike Report — 2026-05-25

## 端点

**主端点**：[URL]
**响应格式**：[JSON / HTML]
**curl 复现**：
\`\`\`bash
curl -H "User-Agent: ..." -H "Referer: ..." "[URL]"
\`\`\`

## 三类 period 实测

| Period | Ticker | announcement_type 字面值 | PDF 直链字段 | 实跑 PDF 大小 |
|---|---|---|---|---|
| ANNUAL | D05 | `"Annual Report"` | `attachment_path` | 8.2 MB |
| H1 | U11 | `"Half Year Results"` | ... | ... |
| IPO | XXXX | `"Prospectus"` | ... | ... |

## 反爬

- 必要 headers: ...
- 限流敏感度: ...
- Cookie/session: ...
- Playwright 需求: ❌ 不需要 / ✅ 需要（因 SPA）

## 已知坑

- [ ] ...
- [ ] ...

## 建议 plugin 实现

- 数据源：JSON API（直接）/ HTML 解析（fallback）
- 限流：1.5 QPS
- 缓存：name_resolver 全表 24h
- Playwright：[需要/不需要]
```

### 1.5 失败回退

如果 spike 证明端点已弃用 / 强 captcha / SPA：
- **不进入批次 2**
- 把 spike 报告填上失败结论
- 通知 Reviewer 回 Planner 重设计（备选数据源：MAS OPERA 仅基金，ACRA 仅注册资料，价值有限）

### 1.6 Commit

```
v1.0: SG plugin API spike（批次 1）

- tmp_pytest_ok/sg_spike.py — 一次性验证 SGXNet 端点
- docs/plans/2026-05-25-sg-spike-report.md — spike 结论
- 三类 period（ANNUAL/H1/IPO）数据可见性确认
- 反爬阶段：[结论]
- Playwright 需求：[结论]
```

### Codex 自检 checklist 批次 1

- [ ] spike 脚本可独立运行（`python tmp_pytest_ok/sg_spike.py`）
- [ ] spike 报告含 3 票 ANNUAL + 1 票 H1 + 1 票 IPO 真实 PDF 下载验证
- [ ] 反爬阶段明确记录（必要 headers / 限流 / Playwright 与否）
- [ ] 报告含 curl 复现命令（Reviewer 可验证）
- [ ] 如 spike 失败：明确结论 + STOP，不进批次 2

---

## 批次 2 — SG plugin 三件套实现

### 2.1 新建文件

按 JP/UK 单模式公网模板，新建 `development/plugins/sg/`：

```
plugins/sg/
  ├─ __init__.py          # SGShare(ExchangePlugin) 类，3 方法委托
  ├─ name_resolver.py     # resolve(code), fetch_company(ticker)
  ├─ reports.py           # download_reports(ticker, period, output_root)
  └─ sgxnet_web.py        # SGXNet HTTP 适配层（搜索 + PDF 直链）
```

### 2.2 复用现有基础设施（强制）

| 设施 | 文件 | 用途 |
|---|---|---|
| HTTP 客户端 | `app/core/http.py::default_client(source="sgxnet")` | **强制走** —— 不允许 SG plugin 自建 httpx.Client |
| 大 PDF 续传 | `app/core/http.py::stream_to_file` | SG PDF 下载用 |
| 限流 | `app/core/ratelimit.py` | SG 注册 source `sgxnet` |
| 单飞缓存 | `app/core/cache.py::cached_or_load` | SG name_resolver 全表 24h |
| 输出路径 | `app/core/output_paths.py::report_output_path` | 扁平命名契约 |
| Plugin ABC | `plugins/base.py::ExchangePlugin` | 3 方法签名严格遵守 |

### 2.3 Models 扩展

`app/core/models.py::Exchange`（约第 8-16 行）：

```python
class Exchange(str, Enum):
    A_SHARE = "A_SHARE"
    HK = "HK"
    US = "US"
    KR = "KR"
    TW = "TW"
    JP = "JP"
    UK = "UK"
    SG = "SG"  # ← 新增
```

`Exchange.display_name`（zh + en mapping）：

```python
@property
def display_name(self) -> str:
    return {
        # ...现有 7 个...
        Exchange.SG: "新股",   # 在 zh dict
    }[self]
```

如果 display_name 走 i18n（请查证当前实现），en 侧加 `"Singapore"`。

### 2.4 Plugin 注册

`plugins/__init__.py::get_plugin`（参考 JP/UK 现有分支）：

```python
def get_plugin(exchange: Exchange) -> ExchangePlugin:
    # ...现有分支...
    if exchange is Exchange.SG:
        from .sg import SGShare
        return SGShare()
    raise NotImplementedError(f"No plugin for {exchange}")
```

### 2.5 Settings 扩展

`app/core/settings.py::RateLimitsCfg`：

```python
class RateLimitsCfg(BaseModel):
    cninfo: float = 5.0
    hkexnews: float = 3.0
    sec: float = 8.0
    dart: float = 5.0
    akshare: float = 4.0
    nsm: float = 2.0
    twse: float = 2.0
    edinet: float = 2.0
    sgxnet: float = 1.5  # ← 新增
```

**不需要** `SGCfg` 类（SG 无 API Key）。

### 2.6 Period 兼容性

检查 `app/core/models.py::PeriodType`：

- 已有 `ANNUAL` / `Q1` / `Q2` / `Q3` / `IPO_PROSPECTUS`？如有 → 直接复用，约定 H1 = Q2（在 ARCHITECTURE.md 文档化）
- 已有 `H1` / `H2`？如有 → 优先用 H1
- 都无 → 新增 `H1` 枚举值（影响面广，需检查所有 plugin 是否需更新）

**默认假设**：复用 Q2 等同 H1（最小改动），文档化即可。

### 2.7 plugins/sg/__init__.py 骨架

```python
"""Singapore Exchange (SGX) plugin — SGXNet public disclosures."""
from __future__ import annotations

from pathlib import Path

from app.core.models import Company, Period, ReportFile, Ticker
from plugins.base import ExchangePlugin

from .name_resolver import fetch_company as _fetch_company
from .name_resolver import resolve as _resolve
from .reports import download_reports as _download_reports


class SGShare(ExchangePlugin):
    def resolve_name(self, code: str) -> Ticker:
        return _resolve(code)

    def fetch_company(self, ticker: Ticker) -> Company:
        return _fetch_company(ticker)

    def download_reports(
        self,
        ticker: Ticker,
        period: Period,
        output_root: Path,
    ) -> list[ReportFile]:
        return _download_reports(ticker, period, output_root)
```

### 2.8 plugins/sg/name_resolver.py 要求

- `resolve(code: str) -> Ticker`：接受 `D05` / `d05` / `D05.SI` / `d05.si` 输入；规范化为大写 + 去 `.SI`
- `fetch_company(ticker: Ticker) -> Company`：通过 SGXNet 搜索 API 获取公司名 + 行业（如可得）；走 `cached_or_load`
- 失败必须 `raise ValueError`（按 plugins/base.py 契约，**不能返回 None**）

```python
def resolve(code: str) -> Ticker:
    normalized = code.strip().upper().removesuffix(".SI")
    if not re.match(r"^[A-Z0-9]{1,5}$", normalized):
        raise ValueError(f"Invalid SG ticker format: {code}")
    # 调 SGXNet 拉公司名
    name = _fetch_name(normalized)
    if not name:
        raise ValueError(f"SG ticker {normalized} not found on SGXNet")
    return Ticker(code=normalized, name=name, exchange=Exchange.SG)
```

### 2.9 plugins/sg/reports.py 要求

`download_reports(ticker, period, output_root) -> list[ReportFile]`：

1. 根据 `period.period_type` 决定 announcement_type 过滤：
   - `ANNUAL` → `"Annual Report"`
   - `Q2` / `H1` → `"Half Year Results"`（或 spike 报告中实际字面值）
   - `IPO_PROSPECTUS` → `"Prospectus"`
2. 调 sgxnet_web 拉对应公告列表
3. 按 `period.year` 过滤（注意发布年 vs 报告年的差异，参考 UK NSM 处理）
4. 下载 PDF 用 `stream_to_file`
5. 文件名走 `report_output_path` / `report_output_path_for_filing`
6. 写 sidecar（用 `app/core/sidecar.py::write_sidecar`）
7. **返回 `[]` 表示该 period 无披露**（非异常，符合 base.py 契约）

### 2.10 plugins/sg/sgxnet_web.py 要求

- `search_announcements(symbol, announcement_type, date_range) -> list[dict]`
- 走 `default_client(source="sgxnet")`
- HTML fallback（如 spike 报告显示 JSON 不稳定）
- 限流由 default_client 自动注入（基于 source）

### 2.11 单元测试

#### `tests/test_sg_name_resolver.py`

```python
import pytest
from plugins.sg.name_resolver import resolve


@pytest.mark.parametrize("input_code,expected", [
    ("D05", "D05"),
    ("d05", "D05"),
    ("D05.SI", "D05"),
    ("d05.si", "D05"),
])
def test_resolve_normalizes_ticker(input_code, expected, mocked_sgxnet):
    """Ticker 规范化：大写 + 去 .SI 后缀"""
    ticker = resolve(input_code)
    assert ticker.code == expected
    assert ticker.exchange.value == "SG"


def test_resolve_invalid_format_raises_value_error():
    with pytest.raises(ValueError, match="Invalid SG ticker format"):
        resolve("INVALID-CODE")


def test_resolve_unknown_ticker_raises_value_error(mocked_sgxnet_empty):
    with pytest.raises(ValueError, match="not found"):
        resolve("ZZZZ")
```

#### `tests/test_sg_reports.py`

测试三类 period：

```python
def test_download_annual_report(tmp_path, mocked_sgxnet_annual):
    ticker = Ticker(code="D05", name="DBS Group", exchange=Exchange.SG)
    period = Period(year=2024, period_type=PeriodType.ANNUAL)
    reports = download_reports(ticker, period, tmp_path)
    assert len(reports) == 1
    assert reports[0].file_path.name.startswith("SG_D05_DBS")
    assert reports[0].file_path.exists()


def test_download_h1_interim(tmp_path, mocked_sgxnet_h1):
    period = Period(year=2024, period_type=PeriodType.Q2)  # H1 复用 Q2
    reports = download_reports(...)
    ...


def test_download_ipo_prospectus(tmp_path, mocked_sgxnet_ipo):
    period = Period(year=2024, period_type=PeriodType.IPO_PROSPECTUS)
    reports = download_reports(...)
    ...


def test_download_no_disclosure_returns_empty(tmp_path, mocked_sgxnet_empty):
    """无披露 period 返回 [] 而非异常"""
    reports = download_reports(ticker, period, tmp_path)
    assert reports == []
```

#### `tests/test_sg_settings.py`

```python
def test_sgxnet_rate_limit_default():
    cfg = RateLimitsCfg()
    assert cfg.sgxnet == 1.5


def test_sgxnet_rate_limit_hot_reload(settings_path):
    """限流热更新（参考 v0.6.1 测试）"""
    ...
```

### 2.12 验证

```bash
cd development
pytest tests/test_sg_*.py -v
ruff check plugins/sg/
pytest -m "not e2e" -v  # 全量回归，预期 ~150 passed
```

### 2.13 Commit

```
v1.0: SG plugin 三件套实现（批次 2）

- plugins/sg/{__init__,name_resolver,reports,sgxnet_web}.py
- Exchange.SG 枚举 + display_name
- RateLimitsCfg.sgxnet 默认 1.5 QPS
- plugins/__init__.py 注册 SG 分支
- 单测：resolve 规范化、3 类 period 下载、settings 热更新
- 145 → 150 tests，ruff 干净
```

### Codex 自检 checklist 批次 2

- [ ] `plugins/sg/` 4 文件齐全
- [ ] HTTP **100% 走 `default_client(source="sgxnet")`**，无自建 Session（grep 验证）
- [ ] `resolve` 失败 raise ValueError（不返回 None）
- [ ] `download_reports` 无披露返回 `[]`（不抛异常）
- [ ] `Exchange.SG` + `RateLimitsCfg.sgxnet` 已添加
- [ ] 3 类 period（ANNUAL / Q2/H1 / IPO_PROSPECTUS）单测全过
- [ ] 全量 pytest 不退化（145 + ~5 = ~150）

---

## 批次 3 — SG UI 接入 + 双语字符串 + smoke

### 3.1 UI 接线

#### `app/ui/main_view.py`

在 exchange 元组（约第 87-96 行）追加 `Exchange.SG`：

```python
for ex in (
    Exchange.A_SHARE,
    Exchange.HK,
    Exchange.US,
    Exchange.KR,
    Exchange.TW,
    Exchange.JP,
    Exchange.UK,
    Exchange.SG,  # ← 新增
):
    panel = ExchangePanel(ex)
    self._panels[ex] = panel
    layout.addWidget(panel)
    panel.setVisible(False)
```

#### `app/ui/exchange_selector.py`

- chip 注册位置（参考 JP/UK 当前注册）添加 SG
- `_zh_name()` / `_meta()` 映射（或等价方法）加 SG 分支

#### **不**需要 `settings_dialog.py` 改动（SG 无 API Key）

### 3.2 双语字符串

`app/ui/strings.py` 同时在 `STRINGS["zh"]` 和 `STRINGS["en"]` 添加：

| Key | zh | en |
|---|---|---|
| `ES_NAME_SG` | "新股" | "Singapore" |
| `ES_META_SG` | "新加坡交易所 SGX" | "Singapore Exchange" |
| `EP_TITLE_SG` | "新股 · Singapore" | "Singapore" |
| `EP_SUBTITLE_SG` | "SGXNet 公开披露" | "SGXNet public disclosures" |
| `TR_PLACEHOLDER_SG` | "如 D05 / U11 / Z74" | "e.g. D05 / U11 / Z74" |

### 3.3 测试增强

- `tests/test_ui_strings.py` 现有断言（zh/en key 对齐 + en 无 CJK）自动覆盖 5 个新 key（无需修改测试，但需 verify 通过）
- `tests/test_language_switch.py` 遍历 8 市场 chip 文本断言

### 3.4 e2e smoke（用户实测）

```powershell
$env:FS_CAPTURE_RUN_E2E="1"
cd development
pytest tests/e2e -v -k "sg"
```

实跑清单：
- DBS Group (`D05`) 2024 ANNUAL
- UOB (`U11`) 2024 ANNUAL + 2024 H1
- Singtel (`Z74`) 2024 ANNUAL
- 1 个近期 SG IPO（spike 期间确定的 ticker）

**验证**：
- PDF 落地到 `output/` 根目录
- 文件名严格 `SG_{code}_{name}_{year}_{kind_zh}.pdf` 格式
- sidecar 落到 `data/cache/sidecars/SG/{stem}.meta.json`
- pypdf 可解码（无 corrupted PDF）

### 3.5 Commit

```
v1.0: SG UI 接入 + 双语 smoke（批次 3）

- main_view exchange tuple + exchange_selector chip 注册
- strings.py 5 个 SG key × zh/en 双份
- 4 票 e2e smoke（DBS/UOB/Singtel + IPO）
- sidecar 落到 data/cache/sidecars/SG/
- 文件命名扁平契约不破
- ~155 tests，ruff 干净
```

### Codex 自检 checklist 批次 3

- [ ] 启动 GUI 看到 8 个 exchange chip（含 SG）
- [ ] 切换 en 后 SG chip 文字变 "Singapore"
- [ ] `tests/test_ui_strings.py` 自动通过新 key
- [ ] `tests/test_language_switch.py` 遍历测试 8 市场无 CJK 残留
- [ ] e2e smoke 4 票 PDF 落地，命名格式正确
- [ ] sidecar 落 `data/cache/sidecars/SG/`

---

## 🔴 Reviewer Checkpoint A（批次 3 后）

**Reviewer**：Claude Code 启动新会话，执行：

```bash
cd "E:\Claude+CODEX Project\FS Capture/development"
pytest -m "not e2e" -v
ruff check .
```

**逐项审**：
1. ☐ SG plugin 与 JP/UK shape 一致（diff 对比）
2. ☐ HTTP 100% 走 default_client（grep `httpx.Client\(`）
3. ☐ 4 票 smoke PDF 落地 + 扁平命名契约未破
4. ☐ 双语下 SG chip 与面板正确（手动切换两次）
5. ☐ sidecar 落 `data/cache/sidecars/SG/`
6. ☐ ~155 tests 全绿
7. ☐ `Exchange.SG` 在所有需要的地方注册（plugins / models / UI）

**Reviewer 通过条件**：7 项全过。任一失败 → 写回填清单交 Codex 修复。

---

## 批次 4 — 性能基线 benchmark

### 4.1 目的

建立 v0.9 基线，后续优化每项都要 A/B 对比。**不做基线就无法证明优化有效**。

### 4.2 新建测试

`tests/benchmarks/__init__.py`（空文件）
`tests/benchmarks/test_baseline_perf.py`：

```python
"""v0.9 baseline performance benchmark.

Run with: pytest -m benchmark --benchmark-only
Disabled by default (heavy network + time).
"""
import time
from statistics import median

import pytest

from app.core.models import Exchange, Period, PeriodType, Ticker
from plugins import get_plugin


@pytest.mark.benchmark
@pytest.mark.parametrize("exchange,ticker_code", [
    (Exchange.A_SHARE, "600519"),
    (Exchange.HK, "0700"),
    (Exchange.US, "AAPL"),
    (Exchange.KR, "005930"),
    (Exchange.TW, "2330"),
    (Exchange.JP, "7203"),
    (Exchange.UK, "ULVR"),
    (Exchange.SG, "D05"),
])
def test_single_ticker_2024_annual(tmp_path, exchange, ticker_code):
    """8 市场各 1 票 2024 年报，记录抓取时长"""
    plugin = get_plugin(exchange)
    ticker = plugin.resolve_name(ticker_code)
    period = Period(year=2024, period_type=PeriodType.ANNUAL)

    durations = []
    for _ in range(3):  # 跑 3 次取中位数
        start = time.monotonic()
        reports = plugin.download_reports(ticker, period, tmp_path)
        durations.append(time.monotonic() - start)

    median_sec = median(durations)
    print(f"\n[BENCH] {exchange.value} {ticker_code} ANNUAL 2024: {median_sec:.2f}s")
    assert reports  # at least 1 PDF
```

测量场景（全部加 `@pytest.mark.benchmark` marker）：

1. **冷启动到主界面可交互**（exe 启动时间）
2. **单 ticker 单 period 抓取**（8 市场各 1 票 2024 年报）
3. **批量抓取**（每市场各 3 票 × 2024 年报 = 24 任务，并发 4）
4. **name_resolver 冷热缓存差异**（同 ticker 第一次 vs 第二次）
5. **HTTP/2 vs HTTP/1.1**（DART 强制 HTTP/1.1 → 测试启用 HTTP/2 的收益）

### 4.3 pyproject.toml 加 marker

```toml
[tool.pytest.ini_options]
markers = [
    "e2e: end-to-end smoke tests requiring network",
    "benchmark: performance benchmarks (slow, default off)",
]
```

### 4.4 输出

`docs/perf/v0.9_baseline.md` markdown 表格：

```markdown
# v0.9 Baseline Performance

**Date**: 2026-05-XX
**Hardware**: Windows 11, 16GB RAM, 100Mbps WAN
**Build**: v0.9.0 (commit XXXXXXX)

## Scenario 1: 单 ticker 单 period

| Exchange | Ticker | Period | Median (s) | Std Dev (s) | Memory Peak (MB) |
|---|---|---|---|---|---|
| A | 600519 | 2024 ANNUAL | ... | ... | ... |
| HK | 0700 | 2024 ANNUAL | ... | ... | ... |
| ... |

## Scenario 2: 批量抓取（24 任务，max_workers=4）

总时长：... 秒
平均 task 时长：... 秒
峰值并发：... 个

## Scenario 3: name_resolver 缓存

| Ticker | Cold (ms) | Warm (ms) | Speedup |
|---|---|---|---|

## Scenario 4: HTTP/2 vs HTTP/1.1

| Source | HTTP/2 | HTTP/1.1 | Δ |
|---|---|---|---|
```

### 4.5 验证

```bash
cd development
pytest -m benchmark --benchmark-only -s
# 数据落到 stdout，由 Codex 整理填 docs/perf/v0.9_baseline.md
```

### 4.6 Commit

```
v1.0: 性能基线 benchmark（批次 4）

- tests/benchmarks/test_baseline_perf.py
- pyproject.toml 加 benchmark marker
- docs/perf/v0.9_baseline.md 5 个测量场景
- 8 市场各 1 票完整 run 计时数据
- 后续优化 A/B 对比基准
```

### Codex 自检 checklist 批次 4

- [ ] 5 个测量场景齐全
- [ ] 8 市场基线数据有效（无 0s 或异常值）
- [ ] `docs/perf/v0.9_baseline.md` 表格可读
- [ ] `pytest -m benchmark --benchmark-only` 可复现
- [ ] 非 benchmark 测试不受影响

---

## 批次 5 — 抓取性能优化（速度优先）

### 5.1 优化清单（按预期 ROI 排序）

按每子项独立 commit，做完即跑 batch 4 benchmark 对比 v0.9 baseline。**任一项负收益立刻回滚**（commit message 标 `experimental` 留 Reviewer 决定）。

### 5.2 子项 5.1 — 并发度提升（预期 -20% 总时长）

**改动**：

```python
# app/core/settings.py::ConcurrencyCfg
class ConcurrencyCfg(BaseModel):
    max_workers: int = 6  # 4 → 6
```

**限流校准**：检查每数据源每 QPS × max_workers 是否不破限：
- cninfo 5 QPS × 6 workers = 30 req/30s? 需测
- sec 8 QPS × 6 = 48 req/6s 应安全
- nsm 2 QPS × 6 = 12 req/6s 接近上限
- sgxnet 1.5 × 6 = 9 req/6s 接近上限

**现网压测**：每数据源连续抓 10 票确认无 429 / 403。

**如 ban**：单独提该 source rate 而非降 max_workers。

### 5.3 子项 5.2 — name_resolver 启动期预热（预期首次抓取 -50%）

**改动**：`app/ui/main_view.py` 在 panel `setVisible(True)` 时后台跑 `_load_name_map()`：

```python
def _sync_exchange_panels(self) -> None:
    selected = set(self.exchange_selector.selected())
    for ex, panel in self._panels.items():
        visible = ex in selected
        panel.setVisible(visible)
        if visible and not panel.resolved_tickers() and not panel._rows:
            panel.add_row()
        # 新增：visible 后台预热 name_resolver 缓存
        if visible:
            QThreadPool.globalInstance().start(
                _NameResolverPrewarmTask(ex)
            )
```

`_NameResolverPrewarmTask` 调对应 plugin 的 `_load_name_map`，失败静默（不让 UI block）。

### 5.4 子项 5.3 — stream_to_file chunk_size 调优（预期大 PDF -10%）

**改动**：`app/core/http.py::stream_to_file` `chunk_size` 64 KB → 256 KB

```python
def stream_to_file(
    client: httpx.Client,
    url: str,
    dest: Path,
    chunk_size: int = 256 * 1024,  # 64 * 1024 → 256 * 1024
    ...
):
```

**风险**：进度反馈粒度变粗。需检查 ProgressDock 信号频率，确认 UI 流畅。

### 5.5 子项 5.4 — Playwright context 池化（预期 KR/UK 渲染 -30%）

**改动**：`app/core/pdf_renderer.py` 把 `_PW_CONTEXT_SEMAPHORE` 替换为 `_PW_CONTEXT_POOL`（Queue）：

```python
import queue
import threading

_PW_LOCK = threading.Lock()
_PW_RUNTIME = None
_PW_BROWSER = None
_PW_CONTEXT_POOL: queue.Queue = queue.Queue()
_PW_POOL_SIZE = 6  # = max_workers


def _ensure_pool():
    global _PW_RUNTIME, _PW_BROWSER
    with _PW_LOCK:
        if not _PW_CONTEXT_POOL.empty():
            return
        from playwright.sync_api import sync_playwright
        _PW_RUNTIME = sync_playwright().start()
        _PW_BROWSER = _launch_browser(_PW_RUNTIME)
        for _ in range(_PW_POOL_SIZE):
            context = _PW_BROWSER.new_context(user_agent=DEFAULT_UA)
            _PW_CONTEXT_POOL.put(context)


def render_url_to_pdf(url, dest, ...):
    _ensure_pool()
    context = _PW_CONTEXT_POOL.get()  # block until available
    try:
        page = context.new_page()
        try:
            page.goto(url, ...)
            page.pdf(path=str(dest), ...)
        finally:
            page.close()
    finally:
        # 清理 context 状态（防 cookie/cache 跨任务污染）
        context.clear_cookies()
        # context.clear_permissions() — 检查 API 可用性
        _PW_CONTEXT_POOL.put(context)
```

**测试**：`tests/test_playwright_pool.py` —— 顺序两次 render 不共享 cookie。

### 5.6 子项 5.5 — HTTP 连接池调优（预期 -5%）

**改动**：`app/core/http.py::default_client`：

```python
def default_client(source: str = "generic") -> httpx.Client:
    return httpx.Client(
        http2=source != "dart",
        limits=httpx.Limits(
            max_connections=100,
            max_keepalive_connections=40,  # 默认 20 → 40
        ),
        ...
    )
```

**DART HTTP/2 重新评估**：v0.7 因兼容禁掉，httpx >= 0.27 可能已修。spike 测一次。

### 5.7 子项 5.6 — sidecar 写入优化（预期 -3%）

**改动**：`app/core/sidecar.py::write_sidecar`：

```python
def write_sidecar(meta: SidecarMeta, dest: Path) -> None:
    tmp = dest.with_suffix(dest.suffix + ".tmp")
    tmp.write_text(meta.model_dump_json(indent=2), encoding="utf-8")
    tmp.replace(dest)  # atomic rename
```

### 5.8 单元测试

- `tests/test_perf_concurrency.py` — 6 并发不破坏限流
- `tests/test_playwright_pool.py` — context 复用后 cookie 清空
- `tests/test_resolver_prewarm.py` — 后台预热信号 + 缓存命中

### 5.9 验证

```bash
cd development
pytest -m "not e2e" -v
ruff check .
pytest -m benchmark --benchmark-only -s
# 数据追加到 docs/perf/v0.9_baseline.md 的 "v1.0 优化后" 列
```

### 5.10 Commit（每子项独立）

```
v1.0: 性能优化 5.1 - 并发度 4→6 + 限流校准（批次 5）
v1.0: 性能优化 5.2 - name_resolver 启动期预热（批次 5）
v1.0: 性能优化 5.3 - stream_to_file chunk_size 64KB→256KB（批次 5）
v1.0: 性能优化 5.4 - Playwright context 池化（批次 5）
v1.0: 性能优化 5.5 - HTTP 连接池调优（批次 5）
v1.0: 性能优化 5.6 - sidecar 原子写（批次 5）
```

### Codex 自检 checklist 批次 5

- [ ] 6 子项每项独立 commit + 独立 benchmark 对比
- [ ] 任一子项负收益 → 单独 revert（commit message 留迹）
- [ ] Playwright pool 无 cookie 跨任务泄漏（独立性测试通过）
- [ ] 限流压测每市场连续 10 票无 429/403
- [ ] benchmark 总体性能 vs v0.9 baseline 至少 -15%（如不到，Checkpoint B 由 Reviewer 决定是否接受）

---

## 批次 6 — 体积压缩（仅"零体验代价"项）

### 6.1 子项 6.1 — spec excludes 补齐（预期 -30~50 MB）

**改动**：`development/filings_atlas.spec:40-45`：

```python
excludes=[
    # 原有
    "matplotlib", "scipy", "torch", "tensorflow",
    "PySide6.QtQuick", "PySide6.QtMultimedia", "PySide6.QtWebEngineCore",
    "PySide6.Qt3D", "PySide6.QtCharts",
    # 新增
    "numpy.tests", "numpy.testing",
    "pandas.tests", "pandas.plotting",
    "notebook", "jupyter", "IPython",
    "pytest", "pytest_httpx", "ruff",
    "tkinter",
],
```

**注意**：`pandas` 主体不能 exclude（akshare/OpenDartReader 强依赖），只剪 tests/plotting。

### 6.2 子项 6.2 — PySide6 翻译精简（预期 -15~20 MB）

**改动**：`development/build.bat` 加 post-build 步骤：

```batch
@echo off
pyinstaller --noconfirm --clean filings_atlas.spec
if errorlevel 1 exit /b 1

REM 删除非 zh_CN / en 的 PySide6 翻译
for /r "dist\Filings Atlas\_internal\PySide6\translations" %%F in (*.qm) do (
    echo %%~nF | findstr /v /i "zh_CN en_US" >nul && del "%%F"
)

REM 删除 Qt Designer / linguist 资源（dev only）
del /q "dist\Filings Atlas\_internal\PySide6\designer*.exe" 2>nul
del /q "dist\Filings Atlas\_internal\PySide6\linguist*.exe" 2>nul
```

### 6.3 子项 6.3 — collect_all → selective collect（预期 -10~30 MB，需谨慎）

**改动**：`filings_atlas.spec`：

```python
# 原：
# playwright_datas, playwright_binaries, playwright_hiddenimports = collect_all("playwright")

# 新：仅 Chromium，不带 Firefox/WebKit
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

playwright_datas = collect_data_files("playwright")
# 过滤掉 firefox/webkit 二进制
playwright_datas = [(src, dst) for src, dst in playwright_datas 
                    if "firefox" not in src.lower() and "webkit" not in src.lower()]
playwright_binaries = []  # Playwright runtime 自己管浏览器二进制下载
playwright_hiddenimports = collect_submodules("playwright")
```

**风险**：Playwright runtime 启动时找不到 chromium 二进制会 fail。需 spike 验证：
- 打包后第一次 `render_url_to_pdf` 是否会自动下载 chromium（不应，因 PyInstaller 离线分发）
- 或需手动包含 `~/.cache/ms-playwright/chromium-XXXX` 到 datas

**跳过条件**：如打包后 KR/UK e2e smoke 失败，立刻 revert 这一项。

### 6.4 子项 6.4 — UPX 压缩（预期 -5~8 MB）

**改动**：`filings_atlas.spec`：

```python
exe = EXE(
    ...
    upx=True,  # False → True
    ...
)

coll = COLLECT(
    ...
    upx=True,  # False → True
    upx_exclude=[  # 保护关键 DLL 避免 Defender 误报
        "vcruntime*.dll",
        "python*.dll",
        "Qt6*.dll",
        "PySide6*.dll",
    ],
    ...
)
```

**验证 Windows Defender 不误报**：打包后右键 `Filings Atlas.exe` → "用 Microsoft Defender 扫描"。

### 6.5 子项 6.5 — 死代码扫描（预期 -5 MB）

```bash
cd development
ruff check . --select F401  # 未使用 import
# 或
pip install vulture
vulture app/ plugins/ --min-confidence 80
```

扫描清理 v0.2 重定位后遗留的 Excel/数据抓取死代码（已应该清完，但 verify 一次）。

**注意**：删除前 grep 确认无 dynamic import / getattr 引用。

### 6.6 验证

```powershell
cd development
pyinstaller --noconfirm --clean filings_atlas.spec
# 测体积
Get-Item ..\Filings\ Atlas\Filings\ Atlas.exe | Select-Object Length
Get-ChildItem ..\Filings\ Atlas -Recurse | Measure-Object -Property Length -Sum
# 测启动时间
Measure-Command { Start-Process "..\Filings Atlas.exe" -PassThru; Start-Sleep 5 }
# 8 市场 smoke
$env:FS_CAPTURE_RUN_E2E="1"
pytest tests/e2e -v
```

### 6.7 输出

`docs/perf/v1.0_size.md`：

```markdown
# v1.0 Bundle Size Optimization

## Before (v0.9)

- Total: 340 MB
- Top 5: Playwright 150MB, PySide6 120MB, akshare 50MB, OpenDartReader 18MB, base_library 15MB

## After (v1.0)

- Total: XXX MB
- Delta: -XX MB (-XX%)

## Per-optimization

| Optimization | Before (MB) | After (MB) | Δ (MB) |
|---|---|---|---|
| 6.1 spec excludes | ... | ... | ... |
| 6.2 PySide6 翻译 | ... | ... | ... |
| 6.3 selective collect | ... | ... | ... |
| 6.4 UPX | ... | ... | ... |
| 6.5 死代码 | ... | ... | ... |

## Startup time

- v0.9: X.X s
- v1.0: X.X s (no regression)

## Windows Defender scan

✅ Clean / ⚠️ Warning: ...
```

### 6.8 Commit

```
v1.0: 体积压缩（批次 6）

- spec excludes 补齐：numpy.tests / pandas.tests / pytest / tkinter
- PySide6 翻译精简：仅保留 zh_CN + en_US
- Playwright selective collect：仅 Chromium，不带 Firefox/WebKit
- UPX 压缩启用（保护 vcruntime/Qt6/PySide6 DLL）
- 死代码清理
- bundle: 340 MB → XXX MB (-XX%)
- 启动时间不退化
- Windows Defender 实时扫描通过
- 8 市场 e2e 全绿
```

### Codex 自检 checklist 批次 6

- [ ] bundle 实际减重 ≥ 30 MB
- [ ] 启动时间不退化（vs v0.9 baseline）
- [ ] Windows Defender 实时扫描 0 报警
- [ ] 8 市场 e2e smoke 全绿（含新增 SG）
- [ ] 子项 6.3（selective collect）如失败已 revert 并记录

---

## 🔴 Reviewer Checkpoint B（批次 6 后）

**Reviewer**：Claude Code 启动新会话，执行：

```bash
cd "E:\Claude+CODEX Project\FS Capture/development"
pytest -m "not e2e" -v
pytest -m benchmark --benchmark-only -s  # 对比 docs/perf/v0.9_baseline.md
```

**逐项审**：
1. ☐ 性能 A/B 数据：每项优化 vs v0.9 baseline 有正收益（或回滚标记）
2. ☐ 总体性能 ≥ v0.9 -15%
3. ☐ Playwright pool 无 cookie 泄漏（`tests/test_playwright_pool.py` 通过）
4. ☐ 限流连续 50 票各市场无 429/403
5. ☐ bundle 实际减重 ≥ 30 MB
6. ☐ 启动时间不退化
7. ☐ Windows Defender 实时扫描 0 报警
8. ☐ 8 市场 e2e smoke 全绿
9. ☐ docs/perf/v0.9_baseline.md + v1.0_size.md 数据完整

**Reviewer 通过条件**：9 项全过。

---

## 批次 7 — v1.0 发布打磨 + GitHub release

### 7.1 版本数字

| 文件 | 改动 |
|---|---|
| `development/pyproject.toml:3` | `version = "0.9.0"` → `"1.0.0"` |
| `development/version_info.txt` | 0.9.0 → 1.0.0 |
| `development/filings_atlas.spec` | （如有 version 元数据）1.0.0 |

### 7.2 README 截图重拍

- 8 chips 含 SG 可见
- 双语 UI（zh 默认 + en 切换截图各一张）
- 替换 `docs/screenshots/` 现有图

### 7.3 CHANGELOG.md v1.0.0 段

```markdown
## v1.0.0 — 2026-XX-XX

**🎉 首个 GitHub Release**

### 新功能
- ✨ 新增第 8 个市场：新加坡 SGX（DBS / UOB / Singtel + IPO）
- ⚡ 抓取速度提升 XX%（vs v0.9）
- 📦 bundle 体积减少 XX MB（vs v0.9）

### 内部改进
- 并发度默认 4 → 6
- name_resolver 启动期预热
- Playwright context 池化
- HTTP 连接池调优
- spec excludes 补齐

### 兼容性
- v0.9 落地的 sidecar 在 v1.0 启动后可直接读
- output 文件命名扁平契约保持不变
- 用户 config.toml 全兼容
```

### 7.4 ARCHITECTURE.md 更新

- §"如何加新市场" 用 SG 作为最新示例（替换或并列 JP）
- 新增 §"性能优化注意事项"：
  - 并发数与限流的关系
  - Playwright pool 的 cookie 清理
  - chunk_size vs 进度反馈粒度的 tradeoff

### 7.5 PROJECT_RETROSPECTIVE.md § 14 v1.0 postscript

```markdown
## § 14 — v1.0 postscript (2026-XX-XX)

### SG plugin 实施

- spike 期间发现：SGXNet announcement_type 字面值 [实际值]
- 限流敏感度：[实测]
- Playwright 需求：[结论]

### 性能优化

| 子项 | 预期 | 实测 | 状态 |
|---|---|---|---|
| 5.1 并发 4→6 | -20% | ... | ✅/❌ |
| 5.2 预热 | -50% 首次 | ... | ✅/❌ |
| ... |

### 体积压缩

- 实际减重：340 MB → XXX MB
- Playwright selective collect [成功/回滚]

### 不做的项（明确）

- WebView2 替代 Chromium — 影响 KR/JP/UK 体验，未来再评估
- akshare 替换 — 年维护成本高
```

### 7.6 git tag（**需用户授权再推**）

```bash
cd "E:\Claude+CODEX Project\FS Capture"
git tag v1.0.0 HEAD
# 不主动推 origin，等用户在 Reviewer Checkpoint C 通过后亲自：
# git push origin v1.0.0
```

### 7.7 PyInstaller 最终 smoke

```powershell
cd development
pyinstaller --noconfirm --clean filings_atlas.spec
# 测启动 + 8 市场 smoke
$env:FS_CAPTURE_RUN_E2E="1"
pytest tests/e2e -v
```

### 7.8 GitHub release artifact（用户授权后）

```powershell
cd "E:\Claude+CODEX Project\FS Capture"
Compress-Archive -Path "Filings Atlas/*" -DestinationPath "Filings_Atlas_v1.0.0_windows_x64.zip"
gh release create v1.0.0 "Filings_Atlas_v1.0.0_windows_x64.zip" `
    --title "Filings Atlas v1.0.0 — 全球披露图谱" `
    --notes-file CHANGELOG.md
```

### 7.9 Commit

```
v1.0: 发布打磨 + v1.0.0 tag（批次 7）

- pyproject.toml version 0.9.0 → 1.0.0
- CHANGELOG.md v1.0.0 段
- ARCHITECTURE.md 加 SG 示例 + 性能优化笔记
- PROJECT_RETROSPECTIVE.md § 14 v1.0 postscript
- README 截图重拍（8 chips + 双语）
- git tag v1.0.0 已打（待用户授权推 origin）
```

### Codex 自检 checklist 批次 7

- [ ] `pyproject.toml` version = "1.0.0"
- [ ] `git tag -l` 含 v1.0.0
- [ ] **未推 origin**（等用户授权）
- [ ] CHANGELOG / README / ARCHITECTURE 准确反映 v0.9 → v1.0
- [ ] PyInstaller 重新打包后 8 市场 e2e 全绿
- [ ] exe 双击启动 < 5 秒

---

## 🔴 Reviewer Checkpoint C（批次 7 后）

**Reviewer**：Claude Code 启动新会话，执行：

```bash
cd "E:\Claude+CODEX Project\FS Capture/development"
pytest -m "not e2e" -v  # 预期 ~160 passed
ruff check .
$env:FS_CAPTURE_RUN_E2E="1"
pytest tests/e2e -v  # 8 市场全绿
```

**逐项审**：
1. ☐ ~160 tests 全绿 + ruff 干净
2. ☐ 8 市场 e2e smoke 全绿
3. ☐ bundle 体积 ≤ v0.9（数字明确）
4. ☐ 启动时间 ≤ v0.9
5. ☐ CHANGELOG 准确
6. ☐ README 8 chips 截图含 SG
7. ☐ ARCHITECTURE 含 SG 示例 + 性能笔记
8. ☐ PROJECT_RETROSPECTIVE § 14 完整
9. ☐ `git tag -l` 含 v1.0.0
10. ☐ `git ls-remote --tags origin | grep v1.0.0` **不含**（未推）
11. ☐ pyproject version = 1.0.0
12. ☐ sidecar 兼容（v0.9 落地的 sidecar 在 v1.0 启动后可读）
13. ☐ output 命名扁平契约未破
14. ☐ Windows Defender 实时扫描 0 报警

**Reviewer 通过条件**：14 项全过。

**通过后**：通知用户亲自推 tag + 发 GitHub release：

```bash
cd "E:\Claude+CODEX Project\FS Capture"
git push origin v1.0.0
gh release create v1.0.0 "Filings_Atlas_v1.0.0_windows_x64.zip" --notes-file CHANGELOG.md
```

---

## 风险与缓解（Top 6）

| # | 风险 | 缓解 |
|---|---|---|
| 1 | git tag v1.0 / v1.0.0 删除 destructive | 批次 0 用户明确 ack origin 推送状态；如已推则需更复杂的公共历史改写 |
| 2 | SGXNet 端点反爬 / SPA 化导致 spike 失败 | 批次 1 单独切 spike，失败立刻 stop，不投入 plugin 实现 |
| 3 | max_workers 4→6 触发数据源 ban | 批次 5 现网压测 50 票，发现 ban 即回滚或单独降单 source rate |
| 4 | Playwright pool cookie 跨任务污染 | `tests/test_playwright_pool.py` 专测独立性；release 时 `clear_cookies` + `clear_permissions` |
| 5 | UPX 压缩触发 Windows Defender 误报 | 批次 6 后做 Defender 扫描；如误报立刻回滚 `upx=False` |
| 6 | 体积优化 6.3 (selective collect) 破坏 Playwright 启动 | 跳过条件已定义；打包后 KR/UK smoke 失败立刻 revert |

---

## 关键文件路径速查

### 新建
- `tmp_pytest_ok/sg_spike.py`
- `docs/plans/2026-05-25-sg-spike-report.md`
- `docs/perf/v0.9_baseline.md`、`docs/perf/v1.0_size.md`
- `development/plugins/sg/{__init__,name_resolver,reports,sgxnet_web}.py`
- `development/tests/test_sg_{name_resolver,reports,settings}.py`
- `development/tests/test_{perf_concurrency,playwright_pool,resolver_prewarm}.py`
- `development/tests/benchmarks/test_baseline_perf.py`
- `CHANGELOG.md` v1.0 段（如不存在则新建）

### 重命名（git mv）
- `roadmap/SPRINT_v1.0_filings_atlas.md` → `SPRINT_v0.9_filings_atlas.md`
- `roadmap/ROADMAP_v0.6.1_to_v1.0.md` → `ROADMAP_v0.6.1_to_v0.9.md`

### 修改（按批次）
- 批次 0：`CLAUDE.md`、`AGENTS.md`、`PROJECT_RETROSPECTIVE.md`、`README.md`、`pyproject.toml`、`version_info.txt`、`filings_atlas.spec`
- 批次 2：`app/core/models.py`、`plugins/__init__.py`、`app/core/settings.py`
- 批次 3：`app/ui/main_view.py`、`app/ui/exchange_selector.py`、`app/ui/strings.py`
- 批次 5：`app/core/{orchestrator,http,pdf_renderer,sidecar,settings}.py`、`app/ui/main_view.py`
- 批次 6：`development/filings_atlas.spec`、`development/build.bat`
- 批次 7：`pyproject.toml`、`CHANGELOG.md`、`ARCHITECTURE.md`、`PROJECT_RETROSPECTIVE.md`、`README.md`、`docs/screenshots/*`

### 不动（红线）
- `app/main.py` stdio 守护逻辑
- `app/core/http.py::default_client` 签名（仅 limits 内部调，不变 ABI）
- `app/core/output_paths.py::report_output_path` 文件命名契约
- `plugins/base.py` ABC 签名
- 7 现有 plugin 业务逻辑（仅 settings 注册位增加）
- `VBA Captor/` 子项目

---

## 验证

### 单元 + 集成测试（每批必绿）
```bash
cd development
pytest -m "not e2e" -v
ruff check .
```
**v0.9 基线**：145 → **v1.0 预计 ~160 passed**（+SG 10 + perf 5）

### 性能 benchmark
```bash
cd development
pytest -m benchmark --benchmark-only -s
```

### 8 市场 e2e smoke
```powershell
$env:FS_CAPTURE_RUN_E2E="1"
cd development
pytest tests/e2e -v
```

### SG smoke 实跑
- DBS (`D05`) 2024 ANNUAL
- UOB (`U11`) 2024 ANNUAL + 2024 H1
- Singtel (`Z74`) 2024 ANNUAL
- 1 个 spike 期间确定的近期 SG IPO

### bundle + 启动时间
```powershell
cd development
pyinstaller --noconfirm --clean filings_atlas.spec
Get-Item ..\"Filings Atlas\Filings Atlas.exe" | Select-Object Length
Measure-Command { Start-Process "..\Filings Atlas\Filings Atlas.exe" -PassThru; Start-Sleep 5 }
```

---

## 最终 Reviewer Checklist

- [ ] 8 市场 e2e 全绿（含新增 SG）
- [ ] 双语切换无回归
- [ ] v1.0 性能 ≥ v0.9 各场景（benchmark 数据支撑）
- [ ] v1.0 bundle 体积 ≤ v0.9（含具体 MB 差）
- [ ] Windows Defender 实时扫描 0 报警
- [ ] sidecar 兼容（v0.9 落地的 sidecar 在 v1.0 启动后可读）
- [ ] output 命名扁平契约未破
- [ ] `pyproject.toml` version = "1.0.0"
- [ ] `git tag -l` 含 v1.0.0
- [ ] `git ls-remote --tags origin | grep v1.0.0` **不含**（等用户推）
- [ ] CHANGELOG / README / ARCHITECTURE / PROJECT_RETROSPECTIVE 准确反映 v0.9 → v1.0
- [ ] CLAUDE.md ≡ AGENTS.md（双文件同步）

---

## 下一步（用户操作）

1. **批准本 SPRINT 文档**
2. **明确授权批次 0 的 git tag 删除操作**（在对话中说明 origin 是否已推 v1.0/v1.0.0 tag）
3. **把本文档 + plan 文件路径喂给 Codex** 启动批次 0
4. 每批次完成后 Codex 提交 commit，Planner（Claude Code）按 Reviewer Checkpoint 验收
5. 通过 Checkpoint C 后用户**亲自**推 v1.0.0 tag 并发 GitHub release

---

**Planner 签名**：Claude Code (Opus 4.7)
**日期**：2026-05-24
**Plan 文件**：`C:\Users\kaiyu\.claude\plans\distributed-enchanting-tarjan.md`

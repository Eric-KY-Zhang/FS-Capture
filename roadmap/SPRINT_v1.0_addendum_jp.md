# SPRINT v1.0 ADDENDUM — JP EDINET 公网爬虫真实现

**日期**：2026-05-24
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex
**Reviewer**：Claude Code
**状态**：待 Codex 启动
**预计工作量**：2.5-3.5 天（4 批次 + 1 Reviewer Checkpoint）
**发布策略**：合并到 v1.0 首发 GitHub release（**撤回当前 v1.0.0 本地 tag，完工后重打**）

---

## Context

v1.0 主 sprint（SG + 性能 + 体积）已完成并打 v1.0.0 本地 tag（未推 origin）。但 Reviewer 检查时发现：

**JP plugin 当前必须依赖 EDINET API key 才能工作**，`plugins/jp/edinet_web.py` 是个**占位 stub**而非真公网爬虫：

```python
# edinet_web.py:29-34（占位代码）
def list_documents(submit_date):
    try:
        return edinet_api.list_documents(submit_date, api_key=None)  # 直接调 API，无 key 必失败
    except PermissionError:
        return []  # ← 失败返回空，无真实 fallback
```

这与 v0.9 时 KR 的真双模式（OpenDartReader API + `dart_web.py` 真公网爬虫）不一致。v0.9 批次 8 实施 JP 时只抄了 KR 的 shape，没有真做公网爬虫。

**用户决策**（2026-05-24）：
- 不接受"JP 必须 key"作为 v1.0 release 状态
- 撤回本地 v1.0.0 tag → 补做 JP 公网爬虫 → 重打 v1.0.0 tag → push release
- v1.0 首发 GitHub release 必须做到"**8 市场都开箱即用**"

---

## 已锁定决策

| 决策项 | 选定方案 |
|---|---|
| Sprint 归属 | **v1.0 addendum**（非 v1.0.1 patch；撤回 tag 重打）|
| 数据源 | **EDINET 公网搜索页** `https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx`（ASP.NET WebForms）|
| 反爬层级 | 无 captcha / 无 Cloudflare / 无登录 / 无 API key 强制 |
| 实现技术 | **纯 httpx + BeautifulSoup**（无需 Playwright；类比 TW MOPS v0.6 实现）|
| 限流 | **新增独立 source** `edinet_web`，默认 1.0 QPS（独立于带 key 的 `edinet = 2.0`）|
| 模板参考 | `plugins/kr/dart_web.py`（KR 真双模式的成熟模板） |
| Reviewer Checkpoint | 1 个（批次 3 后） |

### Spike 已确认事实

- 搜索字段 `D_Keyword` 接受"提出者/発行者/ファンド/**証券コード**"——可直接用 ticker（如 `7203`）搜，无需 ticker → EDINET code 预映射
- 期间过滤字段 `D_KIKAN`
- 书类种别过滤字段 `Syorui1` `Syorui2` `Syorui3` `Syorui4`
- 隐藏字段 `__VIEWSTATE` / `__EVENTVALIDATION` / `D_Hidden_Param` —— 必须 GET 拿 viewstate 再 POST
- **v0.9 时 edinet_web.py 里的 selector（`W0018vD_KEYWORD` 等）是错的**，必须按真实字段重写

### 未确认事实（批次 1 spike 必须验证）

- PDF 下载链接的 `onclick='PdfClick(...)'` 是 URL 拼接还是 postback
- 翻页是否也是 postback（需带 viewstate）
- session cookie 持久化要求（ASPSESSIONID 等）

---

## 批次概览

| 批次 | 名称 | 工作量 | 输出 |
|---|---|---|---|
| 1 | 实地 spike（GET viewstate → POST 搜 → PDF 下载全链路）| 0.5-1 天 | `tmp_pytest_ok/jp_web_spike.py` + `docs/plans/2026-05-25-jp-edinet-web-spike.md` |
| 2 | `edinet_web.py` 真实现 + 限流注册 | 1-1.5 天 | `plugins/jp/edinet_web.py` 重写 + `settings.py` 加 `edinet_web` |
| 3 | 单测 + e2e smoke + Reviewer Checkpoint | 0.5 天 | `tests/test_jp_edinet_web.py` + e2e 真跑 7203/6758/9984 无 key 状态 |
| 4 | 文档 + v1.0.0 tag 移动 | 0.5 天 | CHANGELOG / README / RETROSPECTIVE 更新 + tag 重打 |

---

## 批次 1 — 实地 spike

### 1.1 目的

确认 PDF 下载链路（`PdfClick(...)` 是 URL 拼接还是 postback）+ 翻页机制 + session cookie 要求，再投入批次 2 实现。

### 1.2 实施位置

`tmp_pytest_ok/jp_web_spike.py`（一次性脚本，commit 保留 1 周）

### 1.3 spike 步骤

```python
# tmp_pytest_ok/jp_web_spike.py 骨架
import httpx
from bs4 import BeautifulSoup

SEARCH_URL = "https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx"

with httpx.Client(
    headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "text/html,application/xhtml+xml,*/*",
        "Accept-Language": "ja,en;q=0.9",
    },
    follow_redirects=True,
    timeout=60.0,
) as client:
    # Step 1: GET 拿 viewstate
    r = client.get(SEARCH_URL)
    print("GET status:", r.status_code, "cookies:", list(client.cookies.keys()))
    soup = BeautifulSoup(r.text, "lxml")
    viewstate = soup.select_one("input[name='__VIEWSTATE']")["value"]
    eventvalidation = soup.select_one("input[name='__EVENTVALIDATION']")["value"]
    hidden_param = soup.select_one("input[name='D_Hidden_Param']")
    print("viewstate length:", len(viewstate), "eventvalidation length:", len(eventvalidation))
    print("hidden_param:", hidden_param.get("value") if hidden_param else None)
    
    # Step 2: POST 搜 7203（丰田）2024 年报
    payload = {
        "__VIEWSTATE": viewstate,
        "__EVENTVALIDATION": eventvalidation,
        "D_Keyword": "7203",
        "Syorui1": "on",  # 有価証券報告書（年报）
        "D_KIKAN": "2024",  # 期间，spike 实测确认格式
        # ... 其他 hidden + checkbox 字段
    }
    if hidden_param:
        payload["D_Hidden_Param"] = hidden_param.get("value", "")
    
    r2 = client.post(SEARCH_URL, data=payload)
    print("POST status:", r2.status_code, "response length:", len(r2.text))
    
    # Step 3: 解析搜索结果，找 PDF 链接
    soup2 = BeautifulSoup(r2.text, "lxml")
    rows = soup2.select("table.Grid tr")  # 或 actual 真实 selector
    print("rows found:", len(rows))
    for tr in rows[:3]:
        pdf_link = tr.select_one("a[onclick*='PdfClick']")
        if pdf_link:
            print("onclick:", pdf_link.get("onclick"))
            print("href:", pdf_link.get("href"))
    
    # Step 4: 拼出 PDF URL 并下载（关键 spike 项）
    # 解 PdfClick(doc_id, ...) 参数，拼真实 URL
    # 或者再发一次 POST 作为 postback
```

### 1.4 spike 报告输出

写到 `docs/plans/2026-05-25-jp-edinet-web-spike.md`，含：

```markdown
# JP EDINET 公网爬虫 Spike Report — 2026-05-25

## 端点
- 搜索 URL: https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx
- PDF 下载 URL: [实际确认]

## 真实表单字段
| 字段 | 用途 | 必填 |
|---|---|---|
| __VIEWSTATE | viewstate | 是 |
| __EVENTVALIDATION | eventvalidation | 是 |
| D_Hidden_Param | hidden state | [实测] |
| D_Keyword | 关键词（接受 ticker）| 是 |
| Syorui1-4 | 书类种别 checkbox | 至少 1 个 |
| D_KIKAN | 期间 | [实测格式] |
| ... | ... | ... |

## PdfClick 解析

[关键 spike 结论]：
- 类型：URL 拼接 / postback
- 真实 PDF URL 格式：`...`
- 或：需要再发 POST 带 [...]

## 翻页机制

- 类型：postback / URL 参数
- 实测方式：[...]

## Session cookie

- 必须 cookies：[...]
- 单 session 内可连续多次搜：是 / 否

## 限流验证
- 连续 10 次 POST（间隔 1s）：[结论]

## 三票验证

| Ticker | 公司 | Period | 搜到 | PDF 下载 | 大小 |
|---|---|---|---|---|---|
| 7203 | 丰田 | 2024 年报 | ✅ | ✅ | X MB |
| 6758 | 索尼 | 2024 年报 | ✅ | ✅ | X MB |
| 9984 | 软银 | 2024 年报 | ✅ | ✅ | X MB |

## 实现建议
- 是否走 default_client(source="edinet_web")
- 是否需要 SECTION cookie 持久化（同 client 复用）
- viewstate 是否可缓存（一次 GET 可服多次 POST？还是每次 POST 都要新 GET？）

## 失败回退
[如 spike 失败]：
- PdfClick 是 postback 且需要 JS 触发 → 切 Playwright（工作量 +1 天）
- 强 IP 限流 → 调低 rate，警告用户
```

### 1.5 失败回退条件

如果 spike 发现：
- PDF 下载强制需要 Playwright（PdfClick 是不可逆向的 JS 调用 + token）→ 转 Playwright 实现，工作量 +1 天
- 强 IP 限流（< 0.5 QPS）→ 实现保留但 UI 提示用户慢速

### 1.6 Commit

```
v1.0: JP EDINET 公网爬虫 spike（addendum 批次 1）

- tmp_pytest_ok/jp_web_spike.py
- docs/plans/2026-05-25-jp-edinet-web-spike.md
- 3 票（7203/6758/9984）2024 年报 PDF 真实下载验证
- PdfClick 解析结论：[结果]
- 翻页机制：[结果]
- 是否需要 Playwright：[结论]
```

### Codex 自检 checklist 批次 1

- [ ] spike 脚本可独立运行
- [ ] 3 票 PDF 真实落地（文件头 `%PDF`）
- [ ] PdfClick 解析方式确认（URL 拼接 OR postback）
- [ ] 翻页机制确认
- [ ] 报告含 curl/Python 复现命令（Reviewer 可独立验证）

---

## 批次 2 — `edinet_web.py` 真实现

### 2.1 改写文件

`plugins/jp/edinet_web.py`（完全重写，替换当前占位 stub）：

```python
"""EDINET public-web fallback (no-key mode) for JP filings."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup
from loguru import logger

from app.core.http import default_client
from app.core.ratelimit import limiter
from app.core.settings import load_settings

_SEARCH_URL = "https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx"

_SELECTORS = {
    # 基于批次 1 spike 报告的真实字段，**不是 v0.9 的假名字**
    "viewstate": "input[name='__VIEWSTATE']",
    "eventvalidation": "input[name='__EVENTVALIDATION']",
    "hidden_param": "input[name='D_Hidden_Param']",
    "result_rows": "...",  # spike 确认
    "pdf_link": "...",     # spike 确认
}

# Syorui code mapping (spike 阶段确认 EDINET 公网的 doc_type 编码)
_SYORUI_BY_DOC_TYPE = {
    "120": ("Syorui1", "annual"),       # 有価証券報告書
    "130": ("Syorui1", "annual"),       # 訂正有価証券報告書
    "140": ("Syorui2", "q1_q3"),        # 四半期報告書
    "160": ("Syorui3", "interim"),      # 半期報告書
    # ...
}


def _edinet_web_rate() -> float:
    return load_settings().rate_limits.edinet_web


def _fetch_viewstate(client) -> dict[str, str]:
    """GET 搜索页拿 viewstate token + cookie."""
    limiter("edinet_web", _edinet_web_rate()).acquire_blocking()
    response = client.get(_SEARCH_URL)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "lxml")
    return {
        "__VIEWSTATE": soup.select_one(_SELECTORS["viewstate"])["value"],
        "__EVENTVALIDATION": soup.select_one(_SELECTORS["eventvalidation"])["value"],
        "D_Hidden_Param": (soup.select_one(_SELECTORS["hidden_param"]) or {}).get("value", ""),
    }


def _search_payload(viewstate: dict[str, str], *, keyword: str, year: int, doc_types: set[str]) -> dict[str, str]:
    payload = dict(viewstate)
    payload["D_Keyword"] = keyword
    payload["D_KIKAN"] = ...  # spike 确认格式
    for code in doc_types:
        field_name, _ = _SYORUI_BY_DOC_TYPE.get(code, (None, None))
        if field_name:
            payload[field_name] = "on"
    return payload


def _parse_pdf_url(soup, *, base_url: str) -> str | None:
    """从搜索结果 row 解出 PDF 真实 URL（参考 spike 报告）."""
    # 基于批次 1 spike 结论实现：
    # - URL 拼接模式：解 onclick="PdfClick('XXX', 'YYY')" 拼成 https://...XXX...
    # - postback 模式：返回特殊 sentinel，由 download_document_pdf 走 postback 路径
    ...


def list_documents(submit_date: str | date, *, year: int = None, ticker: str = None) -> list[dict[str, Any]]:
    """List EDINET documents for a given date or year/ticker combo.
    
    与 edinet_api.list_documents 返回相同的归一化字典格式，使 reports.py 切换无感。
    """
    # 公网模式不是按"提交日"枚举，而是按 ticker + 期间搜（更高效）
    # reports.py 当前按 submit_date 枚举一年 365 天，公网模式可优化为单次按 ticker 查
    ...


def download_document_pdf(doc_id: str, dest: Path, *, viewstate_session: dict | None = None) -> int:
    """Download PDF for a doc_id (走 spike 确认的 URL 拼接或 postback 路径)."""
    ...
```

### 2.2 settings.py 改动

`development/app/core/settings.py::RateLimitsCfg`：

```python
class RateLimitsCfg(BaseModel):
    # ... 现有
    edinet: float = 2.0
    edinet_web: float = 1.0  # ← 新增，公网更慢
    # ... 其他
```

### 2.3 reports.py 改动

`plugins/jp/reports.py` 现有结构（双模式 if/else）保持不变，但需要适配公网模式的 API shape 差异：

```python
# 当前
def _list_documents(submit_date: str) -> list[dict[str, Any]]:
    api_key = _edinet_api_key()
    if api_key:
        from .edinet_api import list_documents
        return list_documents(submit_date, api_key=api_key)
    from .edinet_web import list_documents
    return list_documents(submit_date)  # ← 公网模式按 submit_date 不高效
```

**关键设计决策**：是否改 `_list_filings` 的循环结构？

**方案 A**（保守，最小改动）：保持按 submit_date 枚举 365 天的结构，每天单独查（**公网模式很慢，但 reports.py 改动最小**）
**方案 B**（推荐，按 ticker 一次查）：reports.py 改 `_list_filings` 直接调 `edinet_web.list_documents(year=period.year, ticker=ticker.code)`，公网模式跳过 365 天枚举

**Codex 决定**：按 spike 报告中"公网搜索是否能按期间一次返回全年结果"决定。如果可以，方案 B；如果有结果数限制（如最多 100 条/页），方案 A 但加 ticker 过滤。

### 2.4 验证

```bash
cd development
pytest tests/test_jp_*.py -v
ruff check plugins/jp/
```

### 2.5 Commit

```
v1.0: EDINET 公网爬虫真实现（addendum 批次 2）

- plugins/jp/edinet_web.py 完全重写（_SELECTORS 真实字段、viewstate 处理、PDF 下载）
- settings.py RateLimitsCfg 加 edinet_web = 1.0 QPS
- plugins/jp/reports.py 适配公网模式 API shape（方案 A/B）
- KR dart_web.py 模板风格保持一致
```

### Codex 自检 checklist 批次 2

- [ ] _SELECTORS 全部基于 spike 报告的真实字段
- [ ] HTTP 100% 走 `default_client(source="edinet_web")`，无自建 httpx.Client
- [ ] viewstate 正确处理（GET 拿，POST 带）
- [ ] session cookie 同 client 持久化
- [ ] reports.py 双模式分支干净（有 key 调 edinet_api，无 key 调 edinet_web）
- [ ] 限流注册 `edinet_web = 1.0 QPS`
- [ ] ruff 干净

---

## 批次 3 — 单测 + e2e smoke + Reviewer Checkpoint

### 3.1 单测

新建 `development/tests/test_jp_edinet_web.py`：

```python
def test_fetch_viewstate_parses_three_hidden_fields(mocked_search_get):
    """GET 搜索页正确解出 __VIEWSTATE / __EVENTVALIDATION / D_Hidden_Param"""

def test_search_payload_includes_viewstate_and_doc_type_checkbox():
    """POST payload 含 viewstate + 期间 + Syorui 勾选"""

def test_list_documents_returns_normalized_rows_matching_api_shape(mocked_search_post):
    """公网 list_documents 返回与 edinet_api.list_documents 同 shape 的归一化字典"""

def test_download_document_pdf_writes_pdf_header(mocked_pdf_response, tmp_path):
    """download_document_pdf 写入文件以 %PDF 开头"""

def test_list_documents_handles_pagination(mocked_paginated_search):
    """翻页正确（依据 spike 确认的翻页机制）"""

def test_edinet_web_rate_default_is_1_qps():
    from app.core.settings import RateLimitsCfg
    assert RateLimitsCfg().edinet_web == 1.0
```

修改现有 `test_jp_reports.py` 增加无 key 场景：

```python
def test_reports_uses_edinet_web_when_no_api_key(monkeypatch, tmp_path):
    """API key 缺失时 reports._list_documents 走 edinet_web 而非 edinet_api"""
    monkeypatch.setattr(reports, "_edinet_api_key", lambda: "")
    # mock edinet_web.list_documents 返回 1 行
    # 调 reports.download，断言走了 edinet_web 路径
```

### 3.2 e2e smoke（**用户实跑，必跑**）

```powershell
$env:FS_CAPTURE_RUN_E2E="1"
# 临时清空 EDINET key 模拟"用户没申请"场景
Remove-Item env:EDINET_API_KEY -ErrorAction SilentlyContinue
Remove-Item env:EDINET_SUBSCRIPTION_KEY -ErrorAction SilentlyContinue
# config.toml 也清掉 [edinet] api_key

cd "E:\Claude+CODEX Project\FS Capture\development"
python -m pytest tests/e2e/test_smoke.py -v -k "jp"
```

预期 3 票全绿（公网模式）：
- 7203 丰田 2024 年报
- 6758 索尼 2024 年报
- 9984 软银 2024 年报

PDF 落地到 `output/` 根目录，文件名 `JP_{code}_{name}_2024_年报.pdf`。

### 3.3 🔴 Reviewer Checkpoint（批次 3 后）

Reviewer Claude Code 启动新会话，跑：

```bash
cd "E:\Claude+CODEX Project\FS Capture/development"
pytest -m "not e2e" -v
ruff check .
```

**逐项审**：
1. ☐ `plugins/jp/edinet_web.py` 是真实现非 stub（grep `return []` 不应在 PermissionError catch 后立刻 return）
2. ☐ `_SELECTORS` 字典字段名与 spike 报告一致（不再是 `W0018...` 假名字）
3. ☐ HTTP 100% 走 `default_client(source="edinet_web")`（grep `httpx.Client(`）
4. ☐ `RateLimitsCfg.edinet_web = 1.0` 已加
5. ☐ `reports.py` 双模式分支正确（无 key 走公网，有 key 走 API）
6. ☐ ~175 tests 全绿（170 + jp_edinet_web 5-6 个新测试）
7. ☐ ruff 0 warning
8. ☐ e2e smoke 3 票无 key 状态 PDF 真实落地（用户实跑确认）
9. ☐ 文件命名扁平契约未破

**通过条件**：9 项全过。

### 3.4 Commit

```
v1.0: JP 公网爬虫单测 + e2e smoke（addendum 批次 3）

- tests/test_jp_edinet_web.py 5-6 个测试
- tests/test_jp_reports.py 加无 key 路径测试
- e2e smoke 3 票（7203/6758/9984）无 key 状态实跑确认
- ~175 tests 全绿
```

### Codex 自检 checklist 批次 3

- [ ] 单测覆盖 viewstate 解析 / payload 构造 / 翻页 / PDF 下载 / rate
- [ ] reports.py 无 key 路径单测覆盖
- [ ] e2e smoke 3 票真跑成功
- [ ] PDF 文件头 `%PDF` 验证
- [ ] 全量 pytest 不退化

---

## 批次 4 — 文档更新 + v1.0.0 tag 移动

### 4.1 CHANGELOG.md 调整

修改 v1.0.0 段：

**删除**：
> EDINET Subscription-Key is strongly recommended for the current build.
> Public fallback is retained, but the official EDINET API requires Subscription-Key.

**新增**（v1.0 段的 Added 节）：
> - Japan EDINET public web fallback fully implemented (`edinet_web.py`):
>   - JP downloads now work without an EDINET API key out of the box.
>   - API key remains an optional accelerator with higher rate limit (2.0 vs 1.0 QPS).
>   - Public mode tested with Toyota (`7203`), Sony (`6758`), SoftBank (`9984`) 2024 annual reports.

### 4.2 README.md 调整

- 8 markets 表中 Japan 行的 "Key required" 列从 "Yes (recommended)" 改为 "No (key optional, accelerates)"
- 中文版同步

### 4.3 PROJECT_RETROSPECTIVE.md 调整

§ 14 新增小节：

```markdown
### 14.5 JP 公网爬虫真实现（v1.0 addendum）

v0.9 批次 8 实施 JP 时只做了 API key 路径，`edinet_web.py` 是占位 stub。v1.0 Reviewer 阶段发现并补做了真公网爬虫：

- 数据源：EDINET WEEE0030.aspx ASP.NET WebForms 搜索页
- 反爬：无 captcha / 无登录，但需要处理 `__VIEWSTATE` / `__EVENTVALIDATION` / `D_Hidden_Param`
- 真字段名（不是 v0.9 stub 里的 W0018 假名字）：`D_Keyword` / `Syorui1-4` / `D_KIKAN`
- PdfClick 解析：[批次 1 spike 结论]
- 限流：edinet_web 1.0 QPS（API mode 2.0 QPS）
- 工作量：v1.0 addendum，约 3 天
- 教训：以后新增市场必须有真正的 spike 报告作为前置条件，不能只抄 shape
```

### 4.4 ARCHITECTURE.md 调整

§"如何加新市场"章节如有提 EDINET 例子，更新为含公网爬虫的"真双模式"。

### 4.5 UI strings 调整（如有需要）

如果 settings_dialog 或 onboarding_dialog 有"EDINET key 推荐"类提示，改为"可选（提升速度）"。

### 4.6 v1.0.0 tag 移动

```bash
cd "E:\Claude+CODEX Project\FS Capture"

# 验证 origin 上未推（应该仍为空）
git ls-remote --tags origin | grep "v1\.0" || echo "(confirmed not on origin)"

# 移动本地 v1.0.0 tag 到最新 commit
git tag -d v1.0.0
git tag v1.0.0 HEAD  # 即 addendum 批次 4 的最后 commit
```

### 4.7 Commit（含 tag 移动）

```
v1.0: JP 公网真双模式文档与 release 收尾（addendum 批次 4）

- CHANGELOG: 删除 "key strongly recommended"，新增 "EDINET public fallback fully implemented"
- README: 8 markets 表 Japan "Key required" 改为 "No (optional)"
- PROJECT_RETROSPECTIVE § 14.5: JP 公网爬虫 v1.0 addendum 实施记录
- UI 提示语调整（如有）
- 本地 v1.0.0 tag 移动到最新 commit
- 待用户审核通过后 push origin v1.0.0
```

### Codex 自检 checklist 批次 4

- [ ] CHANGELOG 无 "key strongly recommended" 残留
- [ ] README 中英双份都更新
- [ ] PROJECT_RETROSPECTIVE § 14.5 完整
- [ ] git tag -l 仍显示 v1.0.0，但 `git rev-parse v1.0.0` 已是新 commit
- [ ] origin 仍无 v1.0.0（等用户授权 push）

---

## 关键文件路径速查

### 新建
- `tmp_pytest_ok/jp_web_spike.py`
- `docs/plans/2026-05-25-jp-edinet-web-spike.md`
- `development/tests/test_jp_edinet_web.py`

### 重写
- `development/plugins/jp/edinet_web.py`（替换占位 stub）

### 修改
- 批次 2：`development/app/core/settings.py`、`development/plugins/jp/reports.py`
- 批次 3：`development/tests/test_jp_reports.py`（加无 key 路径）
- 批次 4：`CHANGELOG.md`、`README.md`、`PROJECT_RETROSPECTIVE.md`、`ARCHITECTURE.md`（如需）

### 不动
- `development/plugins/jp/edinet_api.py`（保持现状，仅作 API mode）
- `development/plugins/jp/__init__.py`（薄包装层不动）
- `development/plugins/jp/name_resolver.py`（按需微调，但 `_edinet_api_key` 函数不动）
- 7 个其他市场 plugin
- `app/core/http.py` / `pdf_renderer.py` / `output_paths.py` 等基础设施

---

## 复用现有基础设施

| 设施 | 文件 | 用途 |
|---|---|---|
| HTTP 单 session | `app/core/http.py::default_client` | 必须走它，不自建 httpx.Client |
| 限流 | `app/core/ratelimit.py::limiter` | 注册 `edinet_web` source |
| 设置 | `app/core/settings.py::RateLimitsCfg` | 加 `edinet_web` 字段 |
| 模板参考 | `plugins/kr/dart_web.py` | KR 真双模式的成熟实现 |
| 模板参考 | `plugins/tw/reports.py` | TW MOPS 同样是 ASP-style，纯 httpx 实现的 |

---

## 风险与缓解（Top 4）

1. **PdfClick 是 postback 而非 URL 拼接**
   - **缓解**：批次 1 spike 单独验证；如确认是 postback 且 token 不可逆向，转 Playwright（工作量 +1 天）

2. **viewstate 频繁变化导致 session 失效**
   - **缓解**：参考 KR dart_web 的 client 持久化；每个 ticker 用独立 session 或每 N 个 ticker 重新拿 viewstate

3. **公网限流被 ban**
   - **缓解**：限流默认 1.0 QPS（保守）；批次 3 e2e 实跑确认；如 ban 立刻降到 0.5 QPS + UI 提示

4. **公网搜索返回结果数限制**
   - **缓解**：批次 1 spike 验证翻页机制；如有结果数限制，按 ticker + 年份精确搜（结果数应 < 10 条）

---

## 不在本 sprint 范围

- ❌ 切 Playwright 兜底（除非批次 1 spike 失败必须用）
- ❌ 修改 `edinet_api.py`（保持现状）
- ❌ 修改 `name_resolver.py`（_edinet_api_key 不动）
- ❌ JP IPO 招股说明书支持（v1.1+）
- ❌ EDINET 公网爬虫的全文搜索能力

---

## 验证

### 单元测试
```bash
cd development
pytest -m "not e2e" -v  # 预期 ~175 passed
ruff check .
```

### e2e smoke（**必跑**，模拟用户无 key 场景）
```powershell
Remove-Item env:EDINET_API_KEY -ErrorAction SilentlyContinue
Remove-Item env:EDINET_SUBSCRIPTION_KEY -ErrorAction SilentlyContinue
# 临时编辑 config.toml 删除 [edinet] api_key

$env:FS_CAPTURE_RUN_E2E="1"
cd "E:\Claude+CODEX Project\FS Capture/development"
python -m pytest tests/e2e/test_smoke.py -v -k "jp"
```

预期 3 票全绿：7203 / 6758 / 9984 2024 年报 PDF 真实下载到 `output/`。

### v1.0.0 tag 状态
```bash
cd "E:\Claude+CODEX Project\FS Capture"
git tag -l  # 含 v1.0.0
git rev-parse v1.0.0  # 应等于 HEAD（addendum 批次 4 commit）
git ls-remote --tags origin  # 仍不含 v1.0.0
```

---

## 最终 Reviewer Checklist

- [ ] 8 市场 e2e 全绿（含 SG + 含 JP 无 key 模式）
- [ ] `edinet_web.py` 不再是占位 stub
- [ ] `_SELECTORS` 字段名与真实 EDINET 页面一致
- [ ] reports.py 无 key 路径单测覆盖
- [ ] CHANGELOG / README / PROJECT_RETROSPECTIVE 准确反映"JP open-box ready"
- [ ] ~175 tests 全绿 + ruff 干净
- [ ] git tag v1.0.0 已移动到 addendum 完工 commit
- [ ] origin 仍无 v1.0.0（等用户审核通过后 push）
- [ ] sidecar 兼容（v0.9 + v1.0 主 sprint 落地的 sidecar 都可读）

---

## 下一步（用户操作）

1. **批准本 addendum SPRINT 文档**
2. **把本文档喂给 Codex 启动批次 1**（spike 优先）
3. 批次 1 spike 后回 Planner 评估是否需要切 Playwright
4. 批次 3 后 Reviewer Checkpoint
5. 通过后用户实跑 e2e 验证无 key 模式
6. 然后 push v1.0.0 tag + 发 GitHub release

---

**Planner 签名**：Claude Code (Opus 4.7)
**日期**：2026-05-24
**关联 SPRINT**：`roadmap/SPRINT_v1.0_sg_and_perf.md`（主 sprint）

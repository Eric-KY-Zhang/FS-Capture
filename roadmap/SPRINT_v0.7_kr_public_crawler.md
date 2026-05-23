# SPRINT v0.7 — KR 公网爬虫去 Key + 测试补强

**日期**：2026-05-23
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex（待委派）
**Reviewer**：Claude Code
**状态**：待 Codex 实施
**预计工作量**：1 周
**发布策略**：内部迭代，**不发 GitHub release**（按用户指示，GitHub 只发 v1.0）

---

## Context

v0.6.1 patch sprint（2026-05-23 验收通过，6 commit `dfa461d → 3fa6c12`，72/72 tests pass）已修完 7 个确认 bug。v0.7 进入功能型迭代，主项目是**让韩股插件去 DART API Key 化**——v0.6 以前韩股代码虽然写完但用户从未实跑（因为没注册 Key），现在通过 DART 公网披露页爬虫达到"用户零配置即可用韩股"，用户填了 Key 仍享 OpenAPI 路径（双模式并存）。

v0.7 的次要项目是**补 v0.6 留下的 2 个测试盲区**——US 老 ticker 的 `submissions.files[]` 分页 fallback 路径，以及 TW e2e smoke（v0.6 收官时只跑了单元测试和集成测试，没做真实网络 e2e）。

### v0.7 不做的事（已澄清）

经 Planner 在写 SPRINT 前的代码侦察，**原 ROADMAP 中以下三项在 v0.7 不做**：

| 原计划项 | 状态 | 原因 |
|---|---|---|
| ~~httpx `verify=<str>` 弃用警告统一修~~ | ❌ 已修 | `app/core/http.py:44, 54-60` 已经用 `ssl.create_default_context(cafile=certifi.where())` + `_ssl_context()` lru_cache 缓存。提交 `f243a36 Follow-up cleanup: httpx verify + fiscal_year + doc consistency` 已完成 |
| ~~HK 真实场景测试（09988、00005 etc.）~~ | ❌ 已覆盖 | v0.6.1 commit `1564fad` + `aedaf71` 已新增 `test_alibaba_march_fiscal_year_accepts_same_year_filing_window` + `test_pdf_verification_drops_wrong_year_candidate` + `test_pdf_verification_all_fail_falls_back_with_warning` + 4 个 `fiscal_year_lookup_*` 测试 |
| ~~Plugin 重试逻辑统一（删 TW/KR 自己写的常量）~~ | ❌ 误判 | TW `_EMPTY_RETRIES`（line 110-129）处理 MOPS 偶发返回 200 但 body 空；`_DOWNLOAD_RETRIES`（line 243-295）处理 TWSE 临时 URL session-bound 失效。这些是**业务级**重试（HTTP 状态码正常但语义错），与 HTTP 层 `@retry` 处理的网络级失败是不同 layer 的关注点。KR 实际没有自己写的重试（`grep retry development/plugins/kr/` 无匹配）。**保持现状**，文档化"业务 vs 网络重试边界"延后 |
| ~~UI 字符串集中化~~ | ⏳ 挪到 v0.8 | 12 个 UI 文件共 110 处中文字符串，scope 太大。v0.7 KR 改造 + 测试补强已饱和；UI 集中化在 v0.8 lint 清债期间一并做 |

---

## Part A — KR 公网爬虫去 Key 化（主项目）

主体方案来自 `docs/plans/2026-05-23-kr-no-api-key.md`（已与用户对齐）。本节为 SPRINT 文档化的精确版本——行号已对照 v0.6.1 后的代码 verify 过。

### A.1 方案要点（不变）

1. **不引入 `dart-fss` 库**——自己写轻量公网客户端，避免 +20MB 依赖污染 PyInstaller 体积
2. **双模式并存**：`dart.api_key` 配置保留，语义从"必填"降级为"可选加速器"
3. **corp_code 映射动态按需查**：每只票按 stock_code 单查 DART 公网搜索，结果缓存到 disk（key=`kr:corp:{code}`，永久缓存）
4. **公网爬虫限速 3 req/s**（比 OpenAPI 的 5 req/s 更保守）
5. **HTML 选择器集中到 `_SELECTORS` 字典**，便于改版维护

### A.2 改动文件

#### 新建 `development/plugins/kr/dart_web.py`

DART 公网客户端，封装三个核心函数：

| 函数 | 用途 | 输入 | 输出 |
|---|---|---|---|
| `resolve_corp(stock_code) -> dict \| None` | stock_code → corp_code | `"005930"` | `{"corp_code": "00126380", "corp_name": "삼성전자"}` 或 `None` |
| `list_filings(corp_code, bgn_de, end_de, detail_type) -> pd.DataFrame` | 拉报告列表 | corp_code + 日期范围 + A001/A002/A003 | DataFrame，columns **与 OpenDartReader 完全对齐**（`rcept_no` / `report_nm` / `rcept_dt` / `corp_code`） |
| `list_ipo_filings(corp_code) -> pd.DataFrame` | 拉 IPO 招股书 | corp_code | 同上，但走 `publicType=C` + `detailType=C001` + 空 detailType 两轮（兼容现有 `_list_ipo_filings` 双轮抓取） |

**实现约束**：
- 复用 `plugins/kr/reports.py:_dart_client()`（curl_cffi impersonate session），import 不要循环依赖
- 复用 `app/core/ratelimit.py::limiter("dart_web", _DART_WEB_RATE)`
- 复用 `app/core/cache.py::get_cache()` 缓存 `kr:corp:{stock_code}` 永久（**不设 TTL**，corp_code 终身不变）
- HTML 解析优先 BeautifulSoup4（**v0.7 允许新增依赖**，确认已有则复用；查 `requirements.txt`）
- 模块顶部定义 `_SELECTORS` 字典，集中所有 CSS / XPath 选择器（便于 DART 改版时单点修复）

**目标 URL**（已在 KR plan 中预研验证）：
- `https://dart.fss.or.kr/dsab007/main.do?textCrpNm={code}` — corp_code 搜索
- `https://dart.fss.or.kr/dsac001/search.ax` — POST 搜索备用
- `https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}` — 公告详情页

#### 修改 `development/plugins/kr/name_resolver.py`

**注意 v0.6.1 已改的部分必须保留**：
- `fetch_company` 中的 `induty_code` 兼容代码（line 93-99）— Codex 改动 #4 引入，**保留**

具体改法：

**`_dart()`（line 38-45）**
```python
def _dart():
    api_key = (load_settings().dart.api_key or "").strip()
    if not api_key:
-       raise ValueError("尚未配置 DART API 密钥。韩股官方披露数据需要 DART OpenAPI ...")
+       return None
    return _dart_for_key(api_key)
```

**`_load_map()` → 重构为 `resolve_one(stock_code)`（替换 line 52-71）**

旧版 `_load_map` 一次性拉全表（依赖 `dart.corp_codes`，必须有 Key）。新版按需单查：

```python
def resolve_one(stock_code: str) -> dict | None:
    """stock_code -> {corp_code, corp_name}; cache permanently."""
    cache = get_cache()
    cache_key = f"kr:corp:{stock_code}"
    cached = cache.get(cache_key)
    if cached:
        return cached  # type: ignore[return-value]

    # 1) Prefer OpenAPI when available (fast, well-structured)
    dart = _dart()
    if dart is not None:
        try:
            df = dart.corp_codes
            df = df[df["stock_code"].astype(str).str.zfill(6) == stock_code]
            if not df.empty:
                row = df.iloc[0]
                info = {"corp_code": str(row["corp_code"]), "corp_name": str(row["corp_name"])}
                cache.set(cache_key, info)  # no TTL: corp_code is permanent
                return info
        except Exception as exc:
            logger.warning(f"DART OpenAPI corp lookup failed for {stock_code}: {exc}")
            # fall through to public crawler

    # 2) Fallback to public crawler
    from .dart_web import resolve_corp
    info = resolve_corp(stock_code)
    if info is not None:
        cache.set(cache_key, info)
    return info
```

**`resolve(code)`（line 74-85）**
```python
def resolve(code: str) -> Ticker:
    norm = _normalize_code(code)
    info = resolve_one(norm)
    if not info:
        raise ValueError(f"未找到韩股代码 {code}（请确认代码格式，如 005930）")
    return Ticker(exchange=Exchange.KR, code=norm,
                  name=info["corp_name"], external_id=info["corp_code"])
```

**`fetch_company()`（line 88-107）**

保留 v0.6.1 改动 #4 的 induty_code 兼容代码。无 Key 时跳过 OpenAPI 直接返回最小 Company（已是 try-except 容错）：

```python
def fetch_company(ticker: Ticker) -> Company:
    industry: str | None = None
    extra: dict = {}
    dart = _dart()
    if dart is not None:
        try:
            info_df = dart.company(corp=ticker.external_id or ticker.code)
            if info_df is not None and not info_df.empty:
                row = info_df.iloc[0]
                # DART OpenAPI spells this field as "induty_code"; keep upstream name.
                industry_value = row.get("induty_code", row.get("industry_code", ""))
                industry = str(industry_value) or None
                extra = {k: str(v) for k, v in row.items()}
        except Exception as exc:
            logger.warning(f"DART company info failed for {ticker.code}: {exc}")
    return Company(ticker=ticker, listing_date=None, industry=industry,
                   currency="KRW", extra=extra)
```

#### 修改 `development/plugins/kr/reports.py`

**`_list_filings()`（line 126-142）→ 分支化**
```python
def _list_filings(ticker: Ticker, period: Period) -> pd.DataFrame:
    dart = _dart()
    bgn = f"{period.year}0101"
    end = f"{period.year + 1}0630" if period.type is PeriodType.ANNUAL else f"{period.year}1231"
    detail = _DETAIL_KIND[period.type]

    if dart is not None:
        try:
            df = dart.list(corp=_corp(ticker), start=bgn, end=end,
                           kind="A", kind_detail=detail, final=True)
            return df if isinstance(df, pd.DataFrame) else pd.DataFrame()
        except Exception as exc:
            logger.warning(f"DART list failed for {ticker.code}: {exc}")
            return pd.DataFrame()

    # Public crawler path
    from .dart_web import list_filings
    return list_filings(ticker.external_id or ticker.code, bgn, end, detail)
```

**`_list_ipo_filings()`（line 145-163）→ 同样分支化**
```python
def _list_ipo_filings(ticker: Ticker) -> pd.DataFrame:
    dart = _dart()
    if dart is not None:
        # existing OpenAPI two-pass logic unchanged
        ...

    from .dart_web import list_ipo_filings
    return list_ipo_filings(ticker.external_id or ticker.code)
```

**PDF 下载部分（line 255-322）保持不变** —— `_dart_client()` + curl_cffi impersonate 已是公网爬虫，本来就不依赖 Key。

#### 修改 `development/app/ui/main_view.py:172-184`

**砍掉 KR + 无 Key 的硬拦截弹窗**：

```python
- # Korea needs DART key
- kr_in_use = any(t.exchange == Exchange.KR for t in tickers)
- if kr_in_use and not self.settings.dart.api_key:
-     ret = QMessageBox.question(
-         self,
-         "DART 密钥缺失",
-         "韩股官方披露数据主要来自 DART。当前未配置 DART API 密钥，..."
-         QMessageBox.Yes | QMessageBox.No,
-     )
-     if ret == QMessageBox.Yes:
-         self._open_settings()
-     return
+ # Korea: DART API key is optional; without it we fall back to public crawler.
```

可选：在 status bar 加一行非阻塞提示"未配置 DART Key，韩股将走公网（速度较慢）"——但这是 nice-to-have，不影响 SPRINT 验收。

#### 修改 `development/app/ui/onboarding_dialog.py:37-50`

**文案降级**：
```python
- dart = QLabel(
-     "韩股需要 DART API Key。可在 "
-     '<a href="https://opendart.fss.or.kr/">opendart.fss.or.kr</a> 注册后粘贴到设置。'
- )
+ dart = QLabel(
+     "韩股默认通过 DART 公网披露页抓取，无需配置。如需更快更稳的体验，"
+     '可在 <a href="https://opendart.fss.or.kr/">opendart.fss.or.kr</a> '
+     "免费申请 API Key 后填入设置。"
+ )
```

按钮文案：
- `"现在就配 DART"` → `"先填 Key（可选）"`
- `"稍后再说"` → 不变（"跳过"也可，按 Codex 审美选）

#### 修改 `development/app/ui/settings_dialog.py`

DART Key 输入框的 placeholder / 标签文案：
- placeholder：`"可选——留空走 DART 公网爬虫"`
- 帮助文案：`"DART OpenAPI Key（可选）：填入后享受更快更稳的官方接口。"`

具体 widget 位置 Codex 自行 grep `dart.api_key` 或 `DART` 在 settings_dialog.py 中定位。

#### 修改 `development/app/core/settings.py:31-37` (`RateLimitsCfg`)

新增 `dart_web` 字段：
```python
class RateLimitsCfg(BaseModel):
    cninfo: float = 5.0
    hkexnews: float = 3.0
    sec: float = 8.0
    dart: float = 5.0
+   dart_web: float = 3.0
    akshare: float = 4.0
    twse: float = 2.0
```

`dart_web.py` 使用 `limiter("dart_web", load_settings().rate_limits.dart_web)`，不要硬编码 3.0。

#### 修改 `development/requirements.txt`

- `OpenDartReader` **保留**（双模式仍需）
- 新增 `beautifulsoup4>=4.12` 和 `lxml>=5.0` 仅在不存在时（Codex 先 grep 现有 imports）；如果 akshare 等已经传递依赖了 bs4，不要重复加
- 不要新增 `dart-fss`

### A.3 测试

#### 现有测试不受影响

`tests/integration/test_plugins.py:200-236` 中 KR 用例已用 monkeypatch `_load_map` 和 `_list_filings`——新分支被 monkeypatch 覆盖，无需改动即继续通过。**Reviewer 会确认 KR 已有 3 个 KR 测试用例（`test_kr_resolve_name_uses_corp_code_map` / `test_kr_fetch_company_uses_dart_induty_code` / `test_kr_download_reports_selects_q3_by_month`）继续通过**。

#### 新增单元测试

新建 `development/tests/test_kr_dart_web.py`：

1. **`test_dart_web_resolve_corp_parses_search_result`**
   - 用 fixture HTML（保存在 `tests/fixtures/dart_search_005930.html`，从真实 DART 搜索页保存）
   - mock `_dart_client` 返回静态 HTML
   - 断言 `resolve_corp("005930")` 返回 `{"corp_code": "00126380", "corp_name": "삼성전자"}`

2. **`test_dart_web_list_filings_schema_matches_opendartreader`**
   - mock HTML/JSON 响应
   - 断言返回 DataFrame 的 columns 包含 `["rcept_no", "report_nm", "rcept_dt", "corp_code"]`
   - 字段类型与 OpenDartReader 一致（字符串）

3. **`test_dart_web_resolve_corp_returns_none_for_unknown_code`**
   - mock 空搜索结果
   - 断言返回 None（不抛错）

4. **`test_kr_no_api_key_falls_back_to_public_crawler`**
   - 清空 `settings.dart.api_key`
   - monkeypatch `dart_web.resolve_corp` 返回 `{"corp_code": "00126380", "corp_name": "삼성전자"}`
   - 断言 `KRShare().resolve("005930")` 返回正确 Ticker，且没有 raise

5. **`test_kr_with_api_key_prefers_openapi_path`**
   - 设置 dart.api_key
   - 用 fake `OpenDartReader` 包含 005930
   - 断言走 OpenAPI 路径，`dart_web.resolve_corp` 未被调用

#### Fixtures

`tests/fixtures/dart_search_005930.html` — 从真实 DART 公网搜索 005930 的页面保存（Codex 实施前 step 1 抓一份）

#### Smoke 验收（实跑）

**清空 `config.toml` 中 `dart.api_key`**，依次跑：

| 票 | 期间 | 验证点 |
|---|---|---|
| 013890 Zinus（KOSDAQ） | ANNUAL 2024 | KOSDAQ 中盘股代表 |
| 005930 三星电子 | ANNUAL 2024 + Q3 2024 + audit_report | 大盘 + 多 period_type |
| 000660 SK海力士 | ANNUAL 2023 + 2024 | 多年回溯 |
| 068760 셀트리온제약（原 CJ헬스케어） | IPO_PROSPECTUS | IPO C001 双轮抓取（**注**：CJ헬스케어已被韩美药品收购退市；Codex 选一个**现存**的 IPO 案例如 2023-2024 间上市的中小盘，比如 **377300 카카오페이** 或 **361610 SK아이이테크놀로지**） |
| 035420 NAVER | ANNUAL 2024 + Q2 | 大盘补充 |

**每只票期望**：PDF 落地，文件名扁平契约 `{exchange}_{code}_{name}_{year}_{kind_zh}.pdf`。

**再补一轮"有 Key"路径回归**：填入临时 Key 后跑 005930 ANNUAL 2024，确认 OpenAPI 路径仍工作。

---

## Part B — US `submissions.files[]` 分页 fallback 测试

### B.1 现状

`development/plugins/us/reports.py:141-160` 的 `_filter_filings` 当 `recent` 没有匹配时，会遍历 `filings.get("files")` 拉历史分页。`development/plugins/us/reports.py:163-177` 的 `_all_filing_rows` 也走同样路径。

**测试盲区**：`grep "files" tests/` 无任何匹配 — 现有 `tests/test_selection_logic.py` 和 `tests/integration/test_plugins.py` 都只 mock 了 `recent.*`，分页路径**从未被测试覆盖**。

### B.2 测试要求

新增 `development/tests/test_us_paged_submissions.py`：

1. **`test_filter_filings_falls_back_to_paged_files_when_recent_empty`**
   - mock `get_json` 第一次返回 `{"filings": {"recent": {"form": [], ...}, "files": [{"name": "CIK0000789019-submissions-001.json"}]}}`
   - mock 第二次（拉 paged）返回包含目标 10-K
   - 断言 `_filter_filings(...)` 返回 paged 的 row

2. **`test_filter_filings_returns_empty_when_neither_recent_nor_files_match`**
   - 两次都返回无匹配
   - 断言返回 `[]`

3. **`test_all_filing_rows_merges_recent_and_files`**
   - mock recent 含 form A，files 含 form B
   - 断言返回 [A, B] 且经 `_dedupe_rows` 去重

4. **`test_paged_file_fetch_warns_on_error_and_continues`**
   - mock paged 拉取抛异常
   - 断言走异常分支 logger.warning，**不抛**，继续返回 recent 或 []

**不需要实跑**：MSFT FY2018 这种历史票实跑比较慢且不稳定，单元 mock 即可覆盖路径。

---

## Part C — TW e2e Smoke

### C.1 现状

`development/tests/integration/test_tw_plugin.py` 全部走 monkeypatch 路径（v0.6.1 commit `3fa6c12` 已改为 mock-only 提高 portability）。`tests/e2e/test_smoke.py` 只覆盖 A/HK/US/KR 4 市场，TW 未在 e2e。

### C.2 改动

修改 `development/tests/e2e/test_smoke.py`：

新增 1 个 e2e 测试 `test_tw_smoke_annual_taiwan_semiconductor`：
- 走 `FS_CAPTURE_RUN_E2E=1` 环境变量门禁（与现有 4 市场一致）
- 抓 `2330 台積電` `ANNUAL 2024`
- 断言：PDF 落地 + `path.stat().st_size > 1_000_000`（台积电年报通常 5-15MB，阈值 1MB 偏松但合理）+ 文件名扁平契约

不修改 `pyproject.toml` 的 e2e 配置——沿用现有 e2e marker 即可。

---

## 实施顺序（Codex 分批 commit）

| 批次 | 目标 | 验证 |
|---|---|---|
| **1** | 抓 fixture HTML（`tests/fixtures/dart_search_005930.html`）+ 新建 `dart_web.py` 实现 `resolve_corp` + 加单测 `test_dart_web_resolve_corp_parses_search_result` 和 `test_dart_web_resolve_corp_returns_none_for_unknown_code` | `pytest tests/test_kr_dart_web.py -v` 绿 |
| **2** | 实现 `dart_web.list_filings`（A001 only）+ 单测 `test_dart_web_list_filings_schema_matches_opendartreader` | smoke：清空 Key 后跑 005930 ANNUAL 2024 单条链路 |
| **3** | 重构 `name_resolver._load_map → resolve_one` + `reports._list_filings` 分支化 + 单测 `test_kr_no_api_key_falls_back_to_public_crawler` + `test_kr_with_api_key_prefers_openapi_path` | 5 只 smoke 票全季度跑通（无 Key） |
| **4** | 实现 `dart_web.list_ipo_filings` + 接 `_list_ipo_filings` | IPO 票（Codex 选一个 2023-2024 间上市的） PROSPECTUS 跑通 |
| **5** | `RateLimitsCfg.dart_web` 配置项 + UI 三处文案修改 + 砍掉 `main_view.py:172-184` 硬拦截 | 启动 EXE，无 Key 配置下 KR 流程顺畅 |
| **6** | **Part B**：新建 `tests/test_us_paged_submissions.py` 4 个测试 | `pytest tests/test_us_paged_submissions.py -v` 绿 |
| **7** | **Part C**：`tests/e2e/test_smoke.py` 新增 TW 用例 | `FS_CAPTURE_RUN_E2E=1 pytest tests/e2e/ -v` 5 市场全绿（视用户实测时间） |
| **8** | 文档收尾：`PROJECT_RETROSPECTIVE.md §11 v0.7 postscript`；`ARCHITECTURE.md` plugin 表 KR 行更新（数据源描述加"+ 公网爬虫 fallback"）；`docs/plans/2026-05-23-kr-no-api-key.md` 顶部加"已实施于 v0.7"标记 | 全量回归 `pytest -m "not e2e" -v` 全绿 |

每批结束后 Codex 提 commit，commit message 格式：`v0.7: <一句话变更>（批次 N）`。

---

## 测试矩阵

### 单元测试（必绿）

```bash
cd development
pytest -m "not e2e" -v
```

预期：v0.6.1 基线 72 passed + 本期新增 ~9 个用例（5 个 KR + 4 个 US 分页） → **81 passed**。

### e2e Smoke（用户实测时）

```powershell
$env:FS_CAPTURE_RUN_E2E="1"
cd development
pytest tests/e2e -v
```

预期：5 市场全部 PDF 落地。

### KR 双模式手动验证（用户实测）

**模式 A（无 Key）**：
1. 启动 EXE，设置中确认 DART Key 为空
2. 抓 005930 ANNUAL 2024 + 013890 ANNUAL 2024 + 一只 IPO
3. 观察日志：应出现"DART OpenAPI not configured, falling back to public crawler"或类似提示
4. PDF 落地，文件名 `KR_005930_삼성전자_2024_年报.pdf`

**模式 B（有 Key）**：
1. 设置填入 DART Key
2. 抓 005930 ANNUAL 2024
3. 观察日志：应走 OpenAPI 路径（不出现 fallback 提示）
4. PDF 落地，命名同上

---

## Reviewer Checklist

### 改动正确性

#### KR 公网爬虫
- [ ] `dart_web.py` 输出 DataFrame 列名与 OpenAPI 严格一致：`rcept_no` / `report_nm` / `rcept_dt` / `corp_code`（无下游适配改动）
- [ ] `name_resolver._dart()` 无 Key 时**返回 None** 而非抛 `ValueError`
- [ ] `name_resolver.resolve_one()` 实现按需单查 + 永久缓存（`kr:corp:{code}`）
- [ ] `name_resolver.fetch_company()` 保留 v0.6.1 改动 #4 的 `induty_code or industry_code` 兼容代码
- [ ] `reports._list_filings` / `_list_ipo_filings` 分支化后仍通过现有 3 个 KR 集成测试
- [ ] PDF 下载链路（`reports.py:255-322`）**未改动**
- [ ] HTML 选择器集中在 `dart_web._SELECTORS` 字典（grep 该模块内其他 `find_*` / `select_*` 调用应只引用 `_SELECTORS`）
- [ ] 限速读取 `settings.rate_limits.dart_web` 而非硬编码 3.0
- [ ] `RateLimitsCfg` 新增 `dart_web: float = 3.0` 字段
- [ ] `main_view.py:172-184` KR DART Key 硬拦截**已移除**
- [ ] `onboarding_dialog.py:37-50` 文案降级为"可选 Key"
- [ ] `settings_dialog.py` DART Key 输入框 placeholder/帮助文案已改

#### 测试盲区补强
- [ ] `tests/test_us_paged_submissions.py` 至少 4 个测试用例，覆盖：
  - recent 空 + files 命中
  - recent 空 + files 也无匹配
  - recent + files 合并去重
  - paged 拉取异常时 warning + continue
- [ ] `tests/e2e/test_smoke.py` 新增 TW 用例，门禁与 4 市场一致

#### 实跑验证
- [ ] 5 只 smoke 票全部 PDF 落地（无 Key 模式）：013890 / 005930 / 000660 / IPO 票（Codex 选定的）/ 035420
- [ ] "有 Key"路径回归通过：005930 ANNUAL 2024 不出现 fallback 提示
- [ ] 所有 KR PDF 文件名符合 `KR_{code}_{name}_{year}_{kind_zh}.pdf`

### 质量门禁

- [ ] `pytest -m "not e2e" -v` 全绿，**总数 ≥ 81**（v0.6.1 基线 72 + KR 5 + US 4 ≥ 81）
- [ ] e2e（如果用户实测）：5 市场全绿
- [ ] `app/core/orchestrator.py` / `app/core/http.py` / `app/main.py` **零改动**
- [ ] HK / A 股 / US plugin 代码**零改动**（KR 是 v0.7 唯一动 plugin 的市场，US/HK/TW 只新增测试）
- [ ] `pyproject.toml`：如果新增 bs4/lxml 依赖必须 commit 中明确说明并验证 `python -m build` 仍能打包成功
- [ ] 输出文件命名 schema 未变（`test_output_layout.py` 全绿）

### 文档

- [ ] `PROJECT_RETROSPECTIVE.md §11 v0.7 postscript`：记录"KR 去 Key 化的 4 个关键决策（不引 dart-fss / 双模式 / 按需查 / 选择器集中）"
- [ ] `ARCHITECTURE.md §4 各市场实现概览表` KR 行更新为"OpenDartReader（双模式，无 Key 时走公网爬虫）"
- [ ] `docs/plans/2026-05-23-kr-no-api-key.md` 顶部状态从"待 Codex 实施"改为"已实施于 v0.7（commit X）"
- [ ] **不发 GitHub release，不需要更新 README 顶部版本号或 "What's new" 段**

---

## 风险与缓解

| 风险 | 影响 | 缓解 |
|---|---|---|
| DART 公网 HTML 改版导致解析失败 | KR 无 Key 模式全挂 | 选择器集中在 `_SELECTORS` 单点修复；ParseError 时降级提示"可填 Key 走 OpenAPI 备用通道" |
| `robots.txt` 限制 | 法律 / 道德风险 | 实测 `/dsab007/`、`/dsac001/`、`/dsaf001/` 未被 disallow（仅禁 `/cdic/` 和登录后台）；PDF 下载本来就在抓 `/pdf/*`，新增改动不扩大违规面 |
| 冷门股 corp_code 查不到（KONEX 小盘 / 退市股 / OTC） | 少数票 resolve 失败 | 错误消息明确引导用户填 Key 走 OpenAPI corpCode 全量映射 |
| 公网爬虫被 ban | 短时不可用 | 限速 3 req/s + curl_cffi impersonate；README/UI 加一句"本工具仅访问公开披露页面，请合理控制频次" |
| Codex 误删 v0.6.1 已修的 `induty_code` 兼容代码 | KR 行业字段回归 None | SPRINT 文档明确标注必须保留；Reviewer Checklist 单独 1 项验证 |
| IPO smoke 票选错（CJ헬스케어已退市） | smoke 失败 | SPRINT 已建议 Codex 自选 2023-2024 间现存上市的中小盘 IPO 案例（如 카카오페이） |

---

## 关键文件路径（Codex 速查清单）

### 新建
- `development/plugins/kr/dart_web.py`
- `development/tests/test_kr_dart_web.py`
- `development/tests/test_us_paged_submissions.py`
- `development/tests/fixtures/dart_search_005930.html`

### 修改
- `development/plugins/kr/name_resolver.py`（`_dart` / `_load_map`→`resolve_one` / `resolve` / `fetch_company` 保留 induty_code 兼容）
- `development/plugins/kr/reports.py`（`_list_filings` / `_list_ipo_filings` 分支化，PDF 下载不动）
- `development/plugins/kr/__init__.py`（如有 import 变化）
- `development/app/core/settings.py`（`RateLimitsCfg` + `dart_web` 字段）
- `development/app/ui/main_view.py`（line 172-184 移除）
- `development/app/ui/onboarding_dialog.py`（line 37-50 文案）
- `development/app/ui/settings_dialog.py`（DART Key 输入框文案）
- `development/tests/e2e/test_smoke.py`（新增 TW 用例）
- `development/requirements.txt`（如需 bs4/lxml）
- `PROJECT_RETROSPECTIVE.md`（§11 postscript）
- `ARCHITECTURE.md`（§4 表 KR 行）
- `docs/plans/2026-05-23-kr-no-api-key.md`（状态标记）

### 不动
- `development/app/core/orchestrator.py`
- `development/app/core/http.py`
- `development/app/core/pdf_renderer.py`
- `development/app/core/cache.py`
- `development/app/core/ratelimit.py`（v0.6.1 已做热更新，不需要再改）
- `development/app/main.py`
- 所有 A股 / HK / US / TW plugin（仅新增测试，**不改 plugin 代码**）

---

## 后续衔接

v0.7 验收通过后：
- Planner 起草 `roadmap/SPRINT_v0.8_perf_and_lint.md`，涵盖：
  - Playwright browser 连接池
  - `_load_name_map` 加锁防惊群（KR 现在改成按需单查，可能也要加锁）
  - 大 PDF 断点续传
  - diskcache 批写
  - **UI 字符串集中化（从 v0.7 挪过来的 110 处）**
  - ruff lint 清 114 个历史 warnings
  - Playwright 移除决断（需用户参与）
  - PyInstaller 体积优化

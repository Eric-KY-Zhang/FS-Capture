# 韩股（KR）插件去 DART API Key 化

**日期**：2026-05-23
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex
**Reviewer**：Claude Code
**状态**：已实施于 v0.7（2026-05-23，内部迭代，不发 release）

---

## Context

FS Capture 当前韩股插件硬依赖 DART OpenAPI Key 才能跑通——`name_resolver._load_map()` 用 OpenDartReader 加载 stock_code↔corp_code 映射，`reports._list_filings()` 用 `dart.list()` 拉报告列表。用户没注册 DART Key，KR 代码虽编完但从未实跑。同时 `main_view.py:172-184` 有硬拦截弹窗，KR 勾选且无 Key 时整个抓取链路被掐断。

调研三个并行 Explore agent 后确认：DART 公网披露页（`dart.fss.or.kr/dsab007/main.do`、`dsac001/search.ax`、`dsaf001/main.do`）完全公开免 Key，可按 stock_code/corp_code/日期/披露类型搜索并解析 HTML 取得 `rcept_no`，而现有 PDF 下载链路（`reports.py:294-322`）本身已经在走公网 + curl_cffi impersonate，不依赖 Key。**改造范围只涉及两个 OpenAPI 调用点的新增"公网分支"，不动 PDF 下载、不动文件命名契约**。

目标：用户不填 Key 时韩股自动走 DART 公网爬虫；用户填了 Key 仍享受 OpenAPI 的更稳更快路径（双模式并存）。

## 方案要点（已与用户对齐）

1. **不引入 `dart-fss` 库**——自己写轻量公网客户端，避免 +20MB 依赖污染 PyInstaller 体积
2. **双模式并存**：`dart.api_key` 配置项保留，语义从"必填"降级为"可选加速器"
3. **corp_code 映射动态按需查**：每只票按 stock_code 单查 DART 公网搜索，结果缓存到 disk（key=`kr:corp:{code}`，永久缓存）
4. **公网爬虫限速 3 req/s**（比 OpenAPI 的 5 req/s 更保守）
5. **Smoke 验收**：013890 Zinus（KOSDAQ）、005930 三星电子、000660 SK海力士、068760 CJ헬스케어（IPO 案例）、035420 NAVER

## 改动文件

### 新建

- **`development/plugins/kr/dart_web.py`** —— DART 公网客户端，封装三个函数：
  - `resolve_corp(stock_code) -> {corp_code, corp_name} | None`：GET/POST `dart.fss.or.kr/dsab007/main.do?textCrpNm={code}` 或 `dsac001/search.ax`，解析 HTML 表格首行
  - `list_filings(corp_code, bgn_de, end_de, detail_type) -> pd.DataFrame`：POST 公网搜索（`publicType=A`、`detailType=A001/A002/A003`），输出 columns **与 OpenDartReader 完全对齐**（`rcept_no`/`report_nm`/`rcept_dt`/`corp_code`），让 `_select_filing` 等下游函数零改动
  - `list_ipo_filings(corp_code) -> pd.DataFrame`：同上，`publicType=C`、`detailType=C001` + 空 detailType（兼容现有 `_list_ipo_filings` 的两轮抓取）
  - 复用 `_dart_client()`（curl_cffi）、`limiter("dart_web", _DART_WEB_RATE)`、`get_cache()`
  - HTML 选择器集中到模块顶部 `_SELECTORS` 字典，便于改版维护

### 修改

- **`development/plugins/kr/name_resolver.py`**
  - `_dart()`（38-45）：无 Key 时不再抛 `ValueError`，返回 `None`
  - `_load_map()`（52-71）→ 重构为 `resolve_one(stock_code)`：先查 cache（`kr:corp:{code}`），未命中时优先用 `_dart()` 走 OpenAPI，回落到 `dart_web.resolve_corp`；缓存命中无 TTL（corp_code 终身不变）
  - `resolve()`（74-85）：调 `resolve_one` 替代旧 `_load_map`
  - `fetch_company()`（88-107）：保持现有 try-except 结构，无 Key 时跳过行业字段（已是降级容错）

- **`development/plugins/kr/reports.py`**
  - import 区（17）：`from .name_resolver import _dart` 改为可选 import
  - `_list_filings()`（126-142）：`if _dart() is not None: ...走 OpenAPI...; else: return dart_web.list_filings(...)`
  - `_list_ipo_filings()`（145-163）：同样分支化
  - PDF 下载部分（255-322）**保持不变**

- **`development/app/ui/main_view.py:172-184`**
  - 删除 KR + 无 Key 的硬拦截弹窗。改为信息提示（QToolTip 或不阻塞的状态栏文案），允许直接开始

- **`development/app/ui/onboarding_dialog.py:37-44`**
  - 文案从"韩股需要 DART API Key"改为"韩股默认走 DART 公网；可选填 Key 以加速并提升稳定性"
  - 按钮"现在就配 DART"改为"先填 Key（可选）"或"跳过"

- **`development/app/ui/settings_dialog.py`** （DART Key 输入框）
  - placeholder/帮助文案改为"可选——留空走公网爬虫"

- **`development/app/core/settings.py:31-37`** （`RateLimitsCfg`）
  - 新增 `dart_web: float = 3.0` 字段；`dart_web.py` 用此字段而非硬编码

- **`development/requirements.txt`**
  - `OpenDartReader` 保留（双模式仍需）
  - 不新增依赖（已有 curl_cffi、httpx；如需 HTML 解析能力补 beautifulsoup4/lxml，先确认现有依赖）

## 复用现有基础设施

- `app/core/cache.py:get_cache()` — corp_code 永久缓存
- `app/core/ratelimit.py:limiter()` — 3 req/s 限速
- `app/core/http.py:default_client()` — fallback httpx
- `plugins/kr/reports.py:_dart_client()`（107-119）— curl_cffi impersonate session，**直接复用**
- `plugins/kr/reports.py:_DART_HEADERS`（23-31）— 浏览器伪装请求头

## 测试

### 现有测试不受影响

`tests/integration/test_plugins.py` 中 KR 相关用例（~200-236）已经 monkeypatch `_load_map` 和 `_list_filings`——新分支被 monkeypatch 覆盖，无需改动即继续通过。

### 新增单元测试

- `test_kr_dart_web_resolve_corp`：贴 `tests/fixtures/dart_search_005930.html`，mock httpx，断言返回 `corp_code=00126380`
- `test_kr_dart_web_list_filings_schema`：mock HTML/JSON，断言 DataFrame 列名/类型与 OpenAPI 完全对齐
- `test_kr_no_api_key_path`：清空 `settings.dart.api_key`，monkeypatch `dart_web.*`，验证 `resolve()` 不抛错且能正确返回 Ticker

### Smoke 验收（实跑）

清空 `config.toml` 中 `dart.api_key`，依次跑：
- `013890`（Zinus, KOSDAQ）— ANNUAL 2024
- `005930`（三星电子）— ANNUAL 2024 + Q3 2024 + audit_report
- `000660`（SK海力士）— ANNUAL 2024 + 多年度回溯
- `068760`（CJ헬스케어）— IPO_PROSPECTUS（验 C001 + 增补窗口）
- `035420`（NAVER）— ANNUAL 2024 + Q2

每只都应该有 PDF 落地到 `out_dir/`，文件名仍符合 `{exchange}_{code}_{year}_{period_type}_{kind}.pdf` 扁平契约（由 `tests/test_output_layout.py` 锁定）。

再补一轮"有 Key"路径回归（保留双模式）：填入临时 Key 后同样跑 005930 ANNUAL，确认 OpenAPI 路径仍工作。

## 实施顺序（Codex 分批执行）

| 批次 | 目标 | 验证 |
|---|---|---|
| 1 | 新建 `dart_web.py`，实现 `resolve_corp` + `list_filings`（仅 A001），加单测 | 005930 ANNUAL 单条链路跑通 |
| 2 | 接管 `name_resolver` + `reports._list_filings`，保留 OpenAPI 分支 | 5 只 smoke 票全季度跑通（无 Key） |
| 3 | 实现 `list_ipo_filings`，接 `_list_ipo_filings` | 068760 IPO_PROSPECTUS 跑通 |
| 4 | UI 三处文案 + 砍掉 main_view 硬拦截 | 启动 EXE 无 Key 配置下 KR 流程顺畅 |
| 5 | 限速配置项 + README/UI 免责声明 + e2e smoke 注释更新 | 回归 + 用户 alpha 验收 |

## 风险与缓解

- **DART 公网 HTML 改版**：选择器集中在 `dart_web._SELECTORS`，改版时单点修复；ParseError 时降级提示"可填 Key 走 OpenAPI 备用通道"
- **robots.txt**：实测 `/dsaf001/`、`/dsab007/`、`/dsac001/` 未被 disallow（只禁 `/cdic/` 和登录后台）；PDF 下载部分本来就在抓 `/pdf/*`，新增改动不扩大违规面。README/UI 加一句"本工具仅访问公开披露页面，请合理控制频次"
- **冷门股 corp_code 查不到**：KONEX 小盘股 / 已退市股 / OTC 股可能命中失败，错误信息明确引导用户填 Key 走 OpenAPI corpCode 全量映射
- **公网爬虫速率**：限速 3 req/s + curl_cffi impersonate 已经是行业常见做法（dart-fss、scrapdart 等开源项目都这么用），被 ban 风险可控

## 关键文件路径（Codex 速查清单）

- `development/plugins/kr/name_resolver.py`
- `development/plugins/kr/reports.py`
- `development/plugins/kr/dart_web.py`（新建）
- `development/plugins/kr/__init__.py`
- `development/app/core/settings.py`
- `development/app/ui/main_view.py`
- `development/app/ui/onboarding_dialog.py`
- `development/app/ui/settings_dialog.py`
- `development/tests/integration/test_plugins.py`
- `development/tests/fixtures/dart_search_005930.html`（新增 fixture）

---

## Reviewer Checklist（Claude Code 收尾验收用）

- [ ] `dart_web.py` 输出 DataFrame 列名与 OpenAPI 严格一致（无下游适配改动）
- [ ] `name_resolver._dart()` 无 Key 时返回 None 而非抛错
- [ ] `reports._list_filings`/`_list_ipo_filings` 分支化后仍通过 `tests/integration/test_plugins.py`
- [ ] `main_view.py:172-184` 硬拦截已移除
- [ ] 5 只 smoke 票实跑（无 Key 模式）PDF 落地且文件名扁平契约不破
- [ ] 单独跑一遍"有 Key"路径回归确认双模式可用
- [ ] HTML 选择器集中在 `_SELECTORS` 字典
- [ ] 限速读取 `settings.rate_limits.dart_web` 而非硬编码

# FS Capture 总路线图 v0.6.1 → v0.9

**起草日期**：2026-05-23
**最后更新**：2026-05-24（v0.9 已完成；首个 GitHub Release 延后到 `SPRINT_v1.0_sg_and_perf.md`）
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex
**Reviewer**：Claude Code
**当前版本**：v0.9 internal（2026-05-23 完成，7 市场 A/HK/US/KR/TW/JP/UK）
**发布策略**：v0.6.x / v0.7 / v0.8 / v0.9 内部迭代不发 release；**GitHub Release 延后到 `SPRINT_v1.0_sg_and_perf.md`**

---

## 协作模型

```text
┌─────────────────┐   plan    ┌──────────────┐   commits   ┌────────────────┐
│ Claude (Planner)│ ────────▶ │ Codex (Worker)│ ──────────▶ │Claude (Reviewer)│
└─────────────────┘           └──────────────┘             └────────────────┘
       ▲                                                            │
       └────────────── 缺陷回填 / 二轮迭代 ──────────────────────────┘
```

- **Planner（Claude）**：每个版本出独立 SPRINT 计划文件，含 Context / 改动文件 / 测试方案 / 实施顺序 / Reviewer Checklist
- **Worker（Codex）**：按 SPRINT 文档端到端实施，提交 commits + 自检报告
- **Reviewer（Claude）**：照 Checklist 验收，缺陷回填给 Codex 二轮，通过后由用户最终确认发布

每个版本一份 SPRINT 文档放在 `roadmap/` 下，命名规则 `SPRINT_v{X.Y.Z}_{kebab-slug}.md`。

---

## 版本概览

| 版本 | 主题 | 工作量 | 状态 | 验收门槛 |
|---|---|---|---|---|
| **v0.6.1** | Patch — bug 修复 + 体验微调 | 1-2 天 | ✅ **已完成 2026-05-23** | 7 个 bug 修完，72/72 tests pass，5 市场 smoke 实跑通过 |
| **v0.7** | KR 公网爬虫去 Key + 测试补强 | 1 周 | ✅ **已完成 2026-05-23**（11 commit，85/85 tests，KR 无 Key 4 家实跑通过） | KR 无 Key 跑通 + US 分页测试 + Playwright audit 渲染验证 |
| **v0.8** | 性能 + UI 字符串集中化 + lint 锁定 | 2 周 | ✅ **已完成 2026-05-23**（6 批次，100/100 tests，KR audit e2e + ruff 通过） | Playwright pool / 大 PDF 续传 / name_resolver 单飞缓存 / UI strings.py 集中 / lint pre-commit + CI |
| **v0.9** | **品牌重命名 Filings Atlas + 中英双语 UI + sidecar 迁移 + JP/UK + 增量 + IPO 统一**（内部迭代） | 4-6 周 | ✅ **已完成 2026-05-23**（11 批次 + 3 Checkpoint） | 7 市场 e2e 全绿 + 双语 UI + `output/` 100% PDF |

---

## v0.6.1 — Patch（详细 plan 见 `SPRINT_v0.6.1_patch.md`）

**目标**：把 v0.6 上线后审计发现的 5 个 bug 修掉，不引入新功能。

| # | 改动 | 文件 | 严重度 |
|---|---|---|---|
| 1 | TokenBucket 速率热更新 | `app/core/ratelimit.py` | 🔴 高 |
| 2 | HK PDF 验证失败应 drop 而非降权 | `plugins/hk/reports.py` | 🔴 高 |
| 3 | 补 HK 非 12 月财年公司 8-10 家 | `plugins/hk/fiscal_year.py` | 🔴 高 |
| 4 | 修 `kr/name_resolver.py:96` 列名 typo `induty_code` → `industry_code` | `plugins/kr/name_resolver.py` | 🟡 中 |
| 5 | `ashare/reports.py:62` `zip(strict=True)` 防脏数据 | `plugins/ashare/reports.py` | 🟡 中 |
| 6 | batch import 失败行反馈 | `app/ui/batch_import_dialog.py` | 🟡 中 |
| 7 | `settings.py:117` 首次加载返回 `s` 而非 `model_validate({})` | `app/core/settings.py` | 🟢 低（潜在 future bug） |

**不在 v0.6.1 范围**：性能优化、新市场、KR 公网爬虫、UI 重构。

---

## v0.7 — KR 公网爬虫 + 测试补强

**详细 SPRINT 计划**：`roadmap/SPRINT_v0.7_kr_public_crawler.md`
**主体方案**：`docs/plans/2026-05-23-kr-no-api-key.md`（已与用户对齐）

### 7.1 KR 公网爬虫去 Key 化（主项）

按 `docs/plans/2026-05-23-kr-no-api-key.md` 实施。要点：

- 新建 `plugins/kr/dart_web.py` 封装 DART 公网搜索（dsab007/dsac001/dsaf001）
- 双模式并存：有 Key 走 OpenAPI，无 Key 走公网爬虫
- `_load_map`（一次性拉全表）重构为 `resolve_one`（按需单查），corp_code 永久缓存（`kr:corp:{code}`）
- 限速 3 req/s（新增 `RateLimitsCfg.dart_web`）
- 砍掉 `main_view.py:172-184` KR DART Key 硬拦截弹窗
- HTML 选择器集中到 `_SELECTORS` 字典便于改版维护
- Smoke 实跑 5 只票：013890 / 005930 / 000660 / 一个 2023-2024 间上市的 IPO 票 / 035420

### 7.2 测试补强

- **US `submissions.files[]` 分页 fallback 测试**：新建 `tests/test_us_paged_submissions.py`，单元 mock 覆盖 4 种路径（recent 空+files 命中 / 都无匹配 / recent+files 合并 / paged 拉取异常）
- **TW e2e smoke**：`tests/e2e/test_smoke.py` 新增 TW 2330 台积电 ANNUAL 2024 用例

### 7.3 不在 v0.7 范围（已澄清）

经 v0.7 SPRINT 起草前的代码侦察，以下 ROADMAP 原计划项**已修正**：

| 原计划项 | 处置 | 原因 |
|---|---|---|
| ~~httpx `verify=<str>` 弃用警告统一修~~ | **已修，跳过** | 提交 `f243a36 Follow-up cleanup: httpx verify` 已用 `ssl.create_default_context` |
| ~~HK 选片真实场景测试~~ | **已覆盖，跳过** | v0.6.1 commit `1564fad` + `aedaf71` 已新增阿里 3 月财年 + PDF 验证 drop/fallback 共 7 个测试 |
| ~~统一 plugin 重试策略~~ | **误判，保持现状** | TW `_EMPTY_RETRIES`/`_DOWNLOAD_RETRIES` 是业务级重试（MOPS 空响应、TWSE 临时 URL 失效），与网络层 `@retry` 不同 layer；KR 实际没有自己写的重试 |
| ~~UI 字符串集中化~~ | **挪到 v0.8** | 12 个 UI 文件共 110 处中文字符串，scope 太大；与 v0.8 lint 清债一并做更合适 |

### 7.4 文档

- `PROJECT_RETROSPECTIVE.md §11 v0.7 postscript`
- `ARCHITECTURE.md §4 各市场实现概览表` KR 行（数据源加"+ 公网爬虫 fallback"）
- `docs/plans/2026-05-23-kr-no-api-key.md` 状态标记为"已实施于 v0.7"
- **不发 GitHub release，不更新 README 顶部版本号或 What's new 段**

---

## v0.8 — 性能 + UI 集中化 + Lint 锁定（详细 SPRINT 见 `SPRINT_v0.8_perf_and_ui_strings.md`）

经 v0.8 SPRINT 起草前的代码侦察，**原 ROADMAP v0.8 三项已不成立**：

| 原 ROADMAP 项 | 处置 | 实测证据 |
|---|---|---|
| ~~ruff 0 warning（清 114 个 warnings）~~ | **已完成，改为锁定** | `ruff check development/` 输出 `All checks passed!`；pyproject.toml 已严格 select；本期只在 pre-commit + CI 锁定 |
| ~~Playwright 移除决断~~ | **不可行，改池化** | v0.7 KR 4 家 smoke 实跑显示 audit 报告 100% 依赖 `render_url_to_pdf`；US HTML filings 同样依赖。Playwright 是必要基础设施 |
| ~~diskcache 批写~~ | **挪到 v0.9** | 当前 `cache.set` 都不在热循环里，ROI 低 |
| ~~Bundle 体积优化（425MB → 250MB）~~ | **挪到 `SPRINT_v1.0_sg_and_perf.md`**（用户 2026-05-24 决断） | 本期不动；零体验代价压缩留给 SG + 性能 sprint |

### 8.1 v0.8 实际范围（5 项，已完成）

| Part | 内容 | ROI |
|---|---|---|
| **A.1** | Playwright 浏览器池化（singleton browser + per-render context） | 🔴 高（批量渲染省 30-50% 时间） |
| **A.2** | 大 PDF 断点续传（`stream_to_file` 保留 `.part` + Range header + `cleanup_stale_parts`） | 🟡 中（TW IPO / HK 大年报断网恢复） |
| **A.3** | 5 plugin name_resolver 加锁防惊群（新 `app/core/cache.py::cached_or_load` helper） | 🟡 中（4 worker 启动省 3 次冗余请求） |
| **B** | UI 中文字符串物理集中化（97 处 / 11 文件 → `app/ui/strings.py`） | 🟢 低（清债，未来 i18n 铺路） |
| **C** | Lint 锁定（pre-commit + GitHub Actions CI） | 🟢 低 |

### 8.3 v0.8 实施结果

- `cached_or_load` 单飞缓存接入 A 股 / 美股 / 韩股 / 台股名称解析；港股保持单股查询路径。
- `stream_to_file` 支持 `.part` 断点续传、服务器忽略 Range 时重写、416 时清理后重试，并在任务入口清理过期 `.part`。
- Playwright 渲染改为进程级 singleton browser + per-render context，并在应用退出时关闭。
- UI 中文字符串集中到 `app/ui/strings.py`，顶层 UI 模块新增 tokenize 测试锁定无 CJK 字符串字面量。
- Lint 状态通过 `.pre-commit-config.yaml`、GitHub Actions lint/test job、ruff isort first-party 配置锁定。
- 验收：`pytest -m "not e2e" -v` 100 passed / 7 deselected；`ruff check development/` 通过；KR 005930 audit e2e 通过。

### 8.2 不在 v0.8 范围

- ~~Playwright 移除~~（必要基础设施，保留并池化）
- ~~diskcache 批写~~（挪 v0.9）
- ~~Bundle 体积优化~~（挪 `SPRINT_v1.0_sg_and_perf.md`）
- ~~i18n tr() 包装~~（v0.8 只做物理集中，i18n 留 v0.9+ 评估）

---

## v0.9 — Filings Atlas 重塑（详细 SPRINT 见 `SPRINT_v0.9_filings_atlas.md`）

经用户 2026-05-23 提出 3 项新需求，v0.9 在原 ROADMAP 5 项基础上扩展为 **8 项 / 11 批次 mega sprint**，并作为内部迭代版本完成。

### 10.0 v0.9 完整范围

| # | 范围 | 来源 |
|---|---|---|
| 1 | **品牌重命名** "FS Capture" → "Filings Atlas / 全球披露图谱" | 用户 |
| 2 | **中英双语 UI**（运行时切换，Pattern B Python dict）+ README 双语 | 用户 |
| 3 | **Sidecar JSON 迁移**：`output/` → `data/cache/sidecars/`，用户可见目录只剩 PDF | 用户 |
| 4 | **日本 EDINET plugin**（双模式：API key 可选 + 公网兜底） | ROADMAP |
| 5 | **英国 NSM plugin**（公网为主） | ROADMAP |
| 6 | **增量更新**（基于 sidecar diff + UI 独立按钮） | ROADMAP |
| 7 | **IPO 路径统一**（删 ashare/hk IPO helper，收敛 `report_output_path_for_filing`） | ROADMAP |
| 8 | **"如何加新市场" 文档**（追加 ARCHITECTURE.md） | ROADMAP |

**3 个 Reviewer Checkpoint**：批次 4 后（双语完整性）→ 批次 7 后（sidecar/增量/IPO 收敛）→ 批次 9 后（JP/UK 双模式落地）

### 10.1 日本市场（J 股）

**数据源调研结论**（待 v0.9 Sprint 起草时再做深度调研，此处先列方向）：

- **EDINET**（金融厅）— `https://disclosure.edinet-fsa.go.jp/` 有官方 API（v2，免 Key 部分功能，需注册 Key 拿全量）
  - 优势：政府背书、覆盖全部上市公司
  - 风险：日文披露 → 文件名清洗需要处理 CJK 全角 / 半角
- **TDnet**（东证）— 实时披露，但缺历史回溯
- **推荐**：EDINET 为主、TDnet 为补，注册一个免费 Key 拿全量（学习 KR 双模式经验）

**新增文件**：
- `development/plugins/jp/__init__.py`
- `development/plugins/jp/name_resolver.py`
- `development/plugins/jp/reports.py`
- 可能需要 `development/plugins/jp/_pdf_verify.py`（仿 HK）

**Smoke 票**：
- 7203 丰田、6758 索尼、9984 软银集团、9432 NTT、7974 任天堂

### 10.2 伦敦市场（UK 股 / LSE）

**数据源调研方向**：

- **LSE RNS**（Regulatory News Service）— `https://www.londonstockexchange.com/news` 有公开 RSS / HTML
- **National Storage Mechanism (NSM)** — FCA 官方披露存储（`data.fca.org.uk/#/nsm/nationalstoragemechanism`），有公开搜索接口
- **公司官网 IR 页**：英国上市公司普遍设 `/investors/` 页面，质量参差
- **推荐**：NSM 优先（监管来源最权威），LSE RNS 做补充

**新增文件**：同 JP 结构

**Smoke 票**：
- ULVR.L 联合利华、HSBA.L 汇丰（与 HK 重叠测试）、AZN.L 阿斯利康、BP.L、SHEL.L

### 10.3 增量更新检测

**目标**：用户跑过一次后，下个月想"只下新报告"。

**实现思路**：
- 利用现有 sidecar（`*.meta.json`）记录已下载报告的 ticker × period × kind × downloaded_at
- 新加 CLI / UI 入口"增量模式"：扫 `output/*.meta.json` 得到已下载清单，与各 plugin 当前 listing 做 diff，只下新增的
- 默认 1 个月窗口扫一次（用户可配）

**新增文件**：
- `development/app/core/incremental.py`
- UI 增量按钮 / settings 项

### 10.4 IPO 路径统一

当前 5 plugin 各自实现 `_ipo_output_path()` / `_filing_output_path()`，格式不一致：

- HK: `hk_{code}_ipo_{seq}_{date}_{amendment}_{label}.pdf`
- TW: 类似但有差异
- US: 调用 `report_output_path_for_filing()`（已统一）

**改造**：所有 plugin 收敛到 `app/core/output_paths.py::report_output_path_for_filing`，废弃各 plugin 自定义函数。

### 10.5 模块边界文档化

`ARCHITECTURE.md` 新增章节"如何加新市场"：把 v0.9 加 JP/UK 的实际经验固化成 step-by-step 模板，便于未来扩到东南亚（新加坡 SGX、印尼 IDX 等）。

---

## 不在路线图中（明确排除）

- **印度市场（SEBI/BSE/NSE）**：用户明确排除
- **数字财务底稿 / Excel 导出**：v0.2 已砍，永远不做
- **跨市场对标 / 估值计算**：超出工具定位
- **桌面 web 化**：当前 PySide6 桌面应用形态稳定
- **日语 / 韩语 UI**：v0.9 仅 zh + en 双语，dict 结构支持但日语/韩语翻译留 v1.1+
- **其他新市场**（SGX / ASX / SEDAR / Bundesanzeiger / AMF 等）：每个独立 1 周 mini-sprint，v1.1+
- **跨平台打包**（macOS / Linux）：v0.9 仅 Windows，跨平台留 v1.1+
- **本地化文件名**：`_年报.pdf` 中文后缀保留，不随 UI 语言切换

---

## 风险与不确定性

- **v0.7 KR 公网爬虫被 DART 改版**：选择器集中 + Key 兜底，风险可控
- **v0.8 Playwright 移除决断**：需用户参与决策
- **v0.9 EDINET API Key 注册流程**：可能涉及日文身份验证，需提前打通
- **v0.9 LSE 数据源未做实测**：现阶段只是方向，Sprint 起草时需先做一轮独立调研

---

## 历史归档

- `roadmap/archive/` — v0.1 → v0.5 Sprint 文档
- `docs/plans/2026-05-23-kr-no-api-key.md` — KR 公网方案预研（已并入 v0.7）

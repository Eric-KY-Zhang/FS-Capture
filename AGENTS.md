# CLAUDE.md / AGENTS.md — Filings Atlas 协作指南

> 本文件是 Claude Code（作为 Planner / Reviewer）与 Codex（作为 Worker）协作时的常驻上下文。`CLAUDE.md` 与 `AGENTS.md` 是同一份协作指南的双入口。每次进入这个项目，先读这里。
>
> **双文件同步规则**：任何人修改 `CLAUDE.md` 或 `AGENTS.md` 时，必须在同一次改动中同步另一份文件；除文件名入口外，两份文件的章节、规则、日期和内容必须保持一致，禁止只改一边。

---

## 1. 项目一句话

Filings Atlas / 全球披露图谱（原 FS Capture）是跨市场上市公司**官方披露 PDF**一键下载工具，覆盖 A股 / 港股 / 美股 / 韩股 / 台股（v0.6 起 5 市场，v1.0 计划 +日本/英国）。**不抓数据、不算指标、不导 Excel**——这是 v0.2 重定位后的硬约束，任何"顺手做一下"的提议都要先质疑。

仓库根：`E:\Claude+CODEX Project\FS Capture`
源码：`development/`
打包产物：根目录 `Filings Atlas.exe` + `_internal/`（PyInstaller one-folder）

---

## 2. 协作模型

```text
Planner (Claude)  →  Worker (Codex)  →  Reviewer (Claude)  →  User 验收
     ↑                                            │
     └────── 缺陷回填 / 二轮迭代 ─────────────────┘
```

| 角色 | 谁 | 产出 |
|---|---|---|
| **Planner** | Claude Code | 调研代码、设计方案，写 SPRINT 计划文档（含改动文件、测试、Reviewer Checklist） |
| **Worker** | Codex | 按 SPRINT 文档端到端实施，提 commits + 自检报告 |
| **Reviewer** | Claude Code | 照 Checklist 逐项验收，缺陷回填 |

**Claude 不写实现代码**（除非小到 1-2 行的修正补丁，或文档/SPRINT 计划本身）；Codex 不出计划。

### 标准工作流

1. **Planning（Claude Code）**
   - 进入 plan mode（read-only），并行启 Explore agent 调研代码 + 外部资源（最多 3 个）
   - 必要时用 Plan agent 设计方案
   - 关键决策点用 `AskUserQuestion` 对齐而非自行假设
   - 把 plan 写入 `docs/plans/{YYYY-MM-DD}-{slug}.md`（预研阶段）或 `roadmap/SPRINT_v{X.Y.Z}_{slug}.md`（正式 sprint）
   - 调用 `ExitPlanMode` 请用户批准

2. **Implementation（Codex）**
   - 用户把 SPRINT 文档喂给 Codex
   - Codex 按"实施顺序"分批 commit，每批跑对应验证

3. **Review（Claude Code）**
   - 用户回到 Claude Code，要求 review Codex 的 diff
   - Claude Code 按 plan 末尾的 "Reviewer Checklist" 逐项验
   - 必要时跑 e2e smoke

---

## 3. 文档与代码地图

### 文档（按 Planner 起草新 Sprint 时的阅读优先级）
- `roadmap/ROADMAP_v0.6.1_to_v1.0.md` — **总路线图**（v0.6.1 → v1.0 全景）
- `roadmap/SPRINT_v{X.Y.Z}_*.md` — 当前 sprint 详细计划（Codex 必读）
- `docs/plans/*.md` — Planner 写的预研方案（未升级为 SPRINT 前的草稿）
- `ARCHITECTURE.md` — 架构决策与契约（plugin 接口、HTTP、限流、PyInstaller）
- `PROJECT_RETROSPECTIVE.md` — 历史踩坑笔记（**新 Sprint 起草前必看 §4 §5**）
- `development/DEVELOPMENT_BRIEF.md` — 开发约束与已知技术债
- `roadmap/archive/` — v0.1 → v0.5 历史 Sprint 文档

### 代码
- `development/app/core/` — orchestrator / http / ratelimit / cache / pdf_renderer / sidecar / models / settings / output_paths / job
- `development/app/ui/` — PySide6 GUI（main_window / main_view / exchange_panel / ticker_row / progress_dock / settings_dialog / batch_import_dialog / onboarding_dialog / period_selector / output_card）
- `development/plugins/{ashare,hk,us,kr,tw}/` — 5 个市场插件（每个含 `name_resolver.py` + `reports.py`；hk 多 `fiscal_year.py` + `_pdf_verify.py`）
- `development/plugins/base.py` — Plugin ABC 契约
- `development/tests/` — pytest 测试（含 integration / e2e）

### 打包
- 配置：`development/filings_atlas.spec`
- 脚本：`development/build.bat`
- 产物：根目录 `Filings Atlas.exe` + `_internal/` (~340MB)

---

## 4. 不可违反的硬约束

1. **HTTP 一定走 `app/core/http.py::default_client`** — 不允许任何 plugin 自建 httpx Session
2. **`verify=True` 是默认且不可改** — 唯一例外：`source="twse"` 因服务端证书缺 Subject Key Identifier 扩展
3. **PyInstaller windowed 模式 stdio 是 None** — 任何 `print()` / loguru console sink / tqdm 在打包后都会崩溃。`app/main.py:8-25` 的 stdio 守护是硬性留存
4. **路径全部相对 `Path(sys.executable).parent`** — 禁用 `__file__` 推断路径
5. **输出文件命名扁平契约**：`{exchange}_{code}_{name}_{year}_{kind_zh}.pdf`（不嵌套目录），由 `app/core/output_paths.py::report_output_path` 生成，被 `tests/test_output_layout.py` 锁定
6. **Plugin ABC 契约固定**：`plugins/base.py` 三个方法签名不动；`resolve_name` 失败必须 raise ValueError（不能返回 None）；`download_reports` 返回 `[]` 表示该期间无披露（非异常）
7. **`config.toml` 字段不删不改语义**，新增字段可
8. **任何"顺手做"功能必须拒绝**：v0.1 一次顺手做了抓数据 + Excel，v0.2 推倒重来。教训刻在脑子里

---

## 5. 5 市场技术债速查（Planner 起草 Sprint 时高频参考）

| 市场 | 数据源 | 已知坑 | 状态 |
|---|---|---|---|
| **A 股** | akshare + cninfo POST API | 北交所代码是 `bj` 不是 `bse`（已修）；akshare 偶发脏数据需 `zip strict=True` | v0.6.1 修 strict |
| **港股** | 东方财富 + HKEXnews HTML | 无官方 API；选片仅按标题年份字符串，需 PDF 内容验证补充；非 12 月财年表覆盖不足 | v0.6.1 修 |
| **美股** | SEC submissions API | 字段名陷阱：`reportDate` 不是 `periodOfReport`（已修）；老 ticker 走 `submissions.files[]` 分页 fallback 已补单测 | v0.7 已测 |
| **韩股** | OpenDartReader（DART OpenAPI）+ DART 公网披露页 | Key 可选；无 Key 走 `dart_web.py` 公网 fallback；选择器集中在 `_SELECTORS` | v0.7 已完成 |
| **台股** | TWSE ISIN + MOPS | TWSE 证书 hygiene 缺陷必须 `verify=False`；MOPS Big5 编码（不能用 `resp.text`）；ROC 年份 = AD - 1911；mtype=F 是股东会年份 | e2e 已覆盖 |

---

## 6. Plan / Sprint 文档模板

无论是 `docs/plans/` 的预研草稿还是 `roadmap/SPRINT_*.md` 的正式 sprint，结构一致：

- **Context** — 为什么做、不做什么
- **方案要点** — 已与用户对齐的关键决策（预研草稿独有）
- **改动清单** — 按文件分组，每项含：
  - 现状代码引用（带 `file:line`）
  - 问题描述
  - 修复要求
  - 单元测试要求
- **复用现有基础设施** — 引用已有 utility，避免 Codex 重复造轮子
- **实施顺序** — 分批表格，每批一个可独立验证的 milestone + commit
- **测试矩阵** — 单测命令 + smoke 实跑清单
- **风险与缓解** — Codex 实施时的边界与降级路径
- **关键文件路径** — Codex 速查清单
- **Reviewer Checklist** — 逐项可勾选

---

## 7. Planner 行为准则

### 做什么
- 起草 SPRINT 时先读 `PROJECT_RETROSPECTIVE.md §4 §5` 避免重蹈覆辙
- 用 `Glob` / `Grep` / `Read` 定位要改的文件，**给出确切行号**（不要让 Codex 猜）
- bug 报告区分"确认 bug" vs "[推测]"，subagent 报告必须交叉验证（之前出现过 subagent 误判 `us/reports.py:306` 的 sort 是 bug、实际不是）
- bug 格式：`[严重度] file.py:line — 问题 — 复现 — 建议修复`

### 不做
- ❌ 直接帮 Codex 写实现代码（除非 1-2 行的小修补）
- ❌ 主动扩大 Sprint 范围（哪怕看到顺手能修的小 bug，写进下一个 Sprint）
- ❌ 越过 Codex 直接 commit 业务代码（文档和 SPRINT 计划除外）
- ❌ 接受"差不多就行"的 Reviewer 通过——Checklist 要逐项

---

## 8. Reviewer 行为准则

- 逐项过 Checklist，**只看代码不看 Codex 自述**
- 跑 `cd development && pytest -m "not e2e" -v` 确认全绿
- 检查"不应变动"列表确实没被动
- 验证文件命名扁平契约：`{exchange}_{code}_{name}_{year}_{kind_zh}.pdf`
- 如有缺陷，写回填清单（不要直接帮 Codex 改）
- 全过后通知用户做最终验收

---

## 9. Worker（Codex）行为准则

> 本节给 Codex 看的速读版。

### 接到新 Sprint
1. 完整读对应的 `roadmap/SPRINT_v{X.Y.Z}_*.md`
2. 完整读 `ARCHITECTURE.md` + `PROJECT_RETROSPECTIVE.md §4 §5` + 本文件 §4 §5
3. 按 SPRINT 的"实施顺序"分批做，每批一个 commit
4. commit message 格式：`v{X.Y.Z}: <一句话变更>（改动 #N）`

### 实施中
- 行号变了可以调整，但**改动范围不能超出 SPRINT 列出的文件**
- 遇到 SPRINT 没覆盖的边界情况，**写 TODO 留给 Planner 下一轮**，不要自己扩范围决断
- 新增依赖必须先在 commit message 显式说明
- 不碰 `app/core/orchestrator.py`、`app/core/http.py`、`app/main.py` 除非 SPRINT 明确允许

### 自检报告（每批 commit 后）
- 跑 `pytest -m "not e2e" -v` 输出贴一段
- 列出本批改动文件
- 标注与 SPRINT 文档的偏离（若有）

### Smoke
- 最后一批必须实跑 SPRINT "测试矩阵" 中的票
- 把 PDF 落地的文件名贴出来证明扁平契约不破
- 标准 e2e 使用 `FS_CAPTURE_RUN_E2E=1`；台股 IPO 全历史 sweep 属于慢速大体积测试，需额外设置 `FS_CAPTURE_RUN_SLOW_E2E=1` 才运行

---

## 10. 当前版本进度

| 版本 | 状态 | 计划文档 |
|---|---|---|
| v0.6 | ✅ 已发布 2026-05-17 | `roadmap/archive/` |
| v0.6.1 | ✅ **已完成 2026-05-23**（6 commit `dfa461d → 3fa6c12`，72/72 tests，5 票 smoke 实跑） | `roadmap/SPRINT_v0.6.1_patch.md` |
| v0.7 | ✅ **已完成 2026-05-23**（11 commit `2dc5fd2 → d07d133`，85/85 tests，KR 无 Key 4 家实跑） | `roadmap/SPRINT_v0.7_kr_public_crawler.md` |
| v0.8 | ✅ **已完成 2026-05-23**（6 批次，100/100 tests，Playwright 池化 + 断点续传 + UI strings + lint 锁定） | `roadmap/SPRINT_v0.8_perf_and_ui_strings.md` |
| **v1.0** | 🟡 **实施中：批次 1 品牌重命名已完成**（11 批次 + 3 Reviewer Checkpoint，首发 GitHub release）— Filings Atlas / 全球披露图谱 + 双语 UI + sidecar 迁移 + JP/UK + 增量 | `roadmap/SPRINT_v1.0_filings_atlas.md` |

**发布策略**：v0.6.x / v0.7 / v0.8 内部迭代不发 release；**GitHub Release 只发 v1.0**。意味着 v0.7/v0.8 期间：
- 不更新 README 顶部版本号或 "What's new" 段
- 但仍写 `PROJECT_RETROSPECTIVE.md` postscript 章节，保留迭代历史
- 不打 git tag

---

## 11. 路线图速览

详细见 `roadmap/ROADMAP_v0.6.1_to_v1.0.md`。

- **v0.6.1**（已完成）：7 个 bug 修复（ratelimit 热更新 / HK PDF drop / HK 财年表扩 / KR `induty_code` 兼容 / A 股 strict zip / batch import 反馈 / settings 返回值）
- **v0.7**（已完成）：KR 公网爬虫去 Key（双模式）+ US 分页 fallback 单元测试 + Playwright audit 渲染兜底实跑验证。**架构清债项已澄清**：httpx verify 已在 v0.6 后期修过；plugin 重试统一是误判；HK 真实场景测试 v0.6.1 已覆盖；UI 字符串挪到 v0.8
- **v0.8**（已完成）：Playwright **池化**（不删）+ 大 PDF 断点续传 + name_resolver 单飞缓存 + UI 字符串集中（11 文件）+ Lint pre-commit/CI 锁定。**已校准**：ruff 已 0 warning 改为锁定；Playwright 不可删（KR/US 必要依赖）；bundle 体积挪到 v0.9/v1.0
- **v1.0**（4-6 周，11 批次 + 3 Reviewer Checkpoint）：**首发 GitHub release**。①品牌重命名 "FS Capture" → "Filings Atlas / 全球披露图谱"（仓库目录名不变）②中英双语 UI（Pattern B Python dict + Signal，运行时切换）+ README 双语 ③Sidecar 从 `output/` 迁移到 `data/cache/sidecars/`（启动一次性迁移）④日本 EDINET plugin（双模式）⑤英国 NSM plugin（公网为主）⑥增量更新（独立按钮，基于 sidecar diff）⑦IPO 路径统一（删 ashare/hk helper）⑧"如何加新市场" 文档

**明确不做**：印度市场、数字财务底稿、跨市场对标、桌面 web 化。

---

## 12. 用户偏好

- **语言**：中文沟通，代码注释和文档保持中文为主
- **环境**：Windows 11，PowerShell 7
- **背景**：Excel-VBA 出身的金融审计 / 研究背景，看重 UI 设计感 + 数据正确性
- **协作偏好**：Planner / Reviewer 严格把关，Worker 端到端交付；不喜欢中途反复确认琐碎细节
- **特别警惕**："顺手做一下"提议（参见 §1 + §4 第 8 条）

---

## 13. 紧急情况

- **打包后启动闪退**：先查 `crash.log`（EXE 同目录）；最常见原因是 stdio = None 相关
- **数据源 schema 改版**：DART 公网爬虫的 selector 集中在 `_SELECTORS` 字典；其他 plugin 类似集中化是 v0.7+ 改造目标
- **限流被 ban**：调高 `settings.rate_limits.*`，v0.6.1 修完后无需重启即生效

---

**最后更新**：2026-05-23（v1.0 批次 1 品牌重命名完成，下一批次为 i18n 脚手架）

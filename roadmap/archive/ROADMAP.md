# FS Capture 优化路线图

> **创建日期**：2026-05-09
> **基线**：v0.1（A/HK/US 端到端验证；KR 编码完成未实跑）
> **定位**：**纯 PDF 抓取工具**——一键下载 4 市场上市公司的官方披露文件（年报 / 审计报告 / 季报 / 半年报 / IPO 招股书）
> **明确不做**：抓三大报表数字、算财务指标、生成 Excel 底稿、跨市场对标
> **协作模式**：Claude Code（planner + reviewer）↔ `current.md` ↔ Codex（generator）
> **执行节奏**：每个 sprint 独立 commit，Reviewer 验收后才进下一 sprint

## 1. 整体目标

把 FS Capture 从 v0.1（混合工具：PDF + 财务数据 + Excel 底稿）推到 v0.5（纯 PDF 抓取生产可用）：

- **重定位**：移除所有抓三大报表数字、财务指标计算、Excel 底稿装填的代码与文档
- **精确性**：HK 年报选片在多候选公告中能挑对正报；KR 季报关键词能区分 Q1/Q3
- **健壮性**：消除工作线程不安全（KR `os.chdir`）+ UI 假提示（cancel 按钮）+ 资源泄漏（diskcache）
- **用户体验**：首次启动 onboarding + IPO 招股书链路完整 + 北交所代码兼容
- **测试覆盖**：从 6 个 unittest 扩到 plugin 集成 + 端到端冒烟，pytest 迁移

KR 整条链路实跑放在 v0.5 之后（依赖用户注册 DART Key），不在本路线图。

## 2. Sprint 概览

| Sprint | 主题 | 预估工作量 | 依赖 | 风险 |
|---|---|---|---|---|
| **01** | **工具重定位**：移除抓数据代码（fetch_financials / FinancialStatement / excel_writer 等） | 5-7 h | 无 | 中（涉及 model + ABC + 4 plugin + orchestrator + UI + 测试） |
| **02** | 健壮性：工作线程安全 + 资源管理 + UI 假提示 | 3-4 h | Sprint 01（清完无关代码后改才不会乱） | 中（OpenDartReader monkey-patch） |
| **03** | HK 报告选片精确化（PDF 内嵌文本验证 + 财年识别） | 5-7 h | Sprint 01 完成 | 高（pypdf 跨编码 + 历史选片回归） |
| **04** | UX 打磨（首次启动 onboarding + 北交所 + KR Q1/Q3 区分 + IPO 招股书完整支持 + PDF sidecar 元数据） | 4-5 h | Sprint 01 | 低 |
| **05** | 测试基建升级（pytest + plugin 集成测试 + e2e smoke + CI） | 4-6 h | 前 4 sprint 都进入稳态 | 低 |

总工作量：~22 小时 Codex 实施 + ~6 小时 Reviewer 验证。

## 3. 每个 Sprint 的成功标准

### Sprint 01 — 工具重定位
- `plugins/base.py::ExchangePlugin` ABC 只剩 3 个方法：`resolve_name` / `fetch_company` / `download_reports`
- `app/core/models.py` 删除 `FinancialStatement` / `StatementType`
- `plugins/{ashare,hk,us,kr}/financials.py` 4 个文件全部删除
- `app/exporters/excel_writer.py` 全部删除（含 templates/sina_v22_schema.json）
- `app/core/orchestrator.py` 移除 SCRAPING 阶段 + `TaskStatus.SCRAPING` 枚举
- `app/core/job.py::TaskResult.statements` 字段删除
- UI 不再显示"抓取财务数据"步骤；Job 完成提示改成"下载完成 N 个 PDF"
- 测试更新：`test_output_layout.py` 删除工作簿断言；`test_selection_logic.py` 删除 HK long-table pivot + US companyfacts 选片测试
- 端到端冒烟：600519 / 00700 / AAPL，`output/` 平铺得到 PDF，无 .xlsx 产出

### Sprint 02 — 健壮性
- KR `_download_filing` 不再用 `os.chdir`（monkey-patch `OpenDartReader.document` 接受 `output_dir` 参数 / 或 fork 实现把 cwd 参数化）
- Cancel 按钮接通：`progress_dock.cancel_requested` → orchestrator 设置 `threading.Event` → `_TaskRunnable` 在每个步骤 check
- `app/core/cache.py` 加 `close()`，挂到 `QApplication.aboutToQuit`
- `httpx.stream_to_file` 超时拆分（连接 30s + 读 None）
- `ticker_row.name_label` 改 `setTextInteractionFlags(Qt.TextSelectableByMouse)`

### Sprint 03 — HK 选片精确化
- `plugins/hk/reports.py::_select_main` 增加：
  - PDF 内嵌文本（pypdf）验证「年度」+ 财年关键词
  - 财年元数据识别（汇丰 / 阿里巴巴 / 中概股非 12 月财年）
  - 多候选时按"标题精确匹配 > PDF 文本验证 > filingDate 最早"排序
- 加 `tests/test_hk_selection.py`，用 fixture 模拟 5 个真实公司的多候选场景
- 不能让 600519 / AAPL 现有路径退化（跑现有 `test_selection_logic`）

### Sprint 04 — UX 打磨
- 首次启动 onboarding：检测 `config.toml` 不存在 → 弹引导对话框（输入第一个股票代码 + DART 注册链接）
- 北交所代码：`plugins/ashare/reports.py::_column_for` 同时支持 `bj` 与 `bse`
- KR Q1 vs Q3 关键词区分：用 `rcept_dt` 月份 + `report_nm` 关键词组合判断
- IPO 招股书链路完整支持：`PeriodType` 加 `IPO_PROSPECTUS`（或独立 flag），4 市场各 plugin 实现 IPO 招股书查询
- PDF sidecar 元数据：每个下载的 PDF 旁边写 `{filename}.meta.json`（title / period / source_url / download_time / file_hash）

### Sprint 05 — 测试基建
- 迁移 `tests/` 从 unittest 到 pytest（`pyproject.toml` 配置）
- 新增 `tests/integration/`：每市场 plugin × 2 步骤（resolve_name / download_reports）mock-based 测试（共 8 个）
  - 用 `pytest-httpx` 或 `respx` mock httpx 响应
  - fixture：4xx / RemoteDisconnected / 分页 fallback / cache hit
- 新增 `tests/e2e/smoke.py`：联网真打，每市场 1 个 ticker，跑完 resolve → reports，CI skip
- GitHub Actions：lint (ruff) + pytest unit + pytest integration（不跑 e2e）

## 4. Out of Scope（不在本路线图）

- KR 整条链路实跑（依赖用户注册 DART Key，另起 sprint）
- DART Key 平台 keyring 加密存储（v1 接受 config.toml 明文）
- MarketAdapter 抽象（4 市场 plugin 各 ~150 行 boilerplate，重构 ROI 低）
- frameless window 多显示器 / 不同 DPI 适配（用户当前单屏，需求未提）
- ~~Excel 模板注入 / 瑞华底稿 sheet~~（**已废弃**：工具不再生成任何 Excel 底稿）
- ~~跨市场字段中文化 / 三大报表 normalize~~（**已废弃**：工具不再抓财务数字）

## 5. 协作流程

```text
1. Planner（Claude Code）写 SPRINT_NN_xxx.md 详细计划，roadmap/ 下
2. Planner 把当前 sprint 的 hand-off 摘要写到 current.md（一页）
3. 用户启动 Codex 看 current.md → 跳到 SPRINT_NN 详细计划
4. Codex 端到端实施 + 在 current.md 末尾写"## Codex 回报"
5. 用户告知 Reviewer（Claude Code）开始验收
6. Reviewer 跑 frozen 回归 + spot-check + 在 current.md 末尾写"## Reviewer 验收"
7. 验收通过 → 进下一 sprint，Planner 更新 current.md（清空旧轮次或归档）
8. 验收不通过 → Planner 在 current.md 写 follow-up 任务，Codex 修复后再次验收
```

`current.md` 一页内必须装下：当前 sprint 名、阶段（Planner/Codex/Reviewer）、待办、验证、回报、验收。详细计划放 `roadmap/SPRINT_NN_xxx.md`。

## 6. 当前状态

- **进行中**：Sprint 01 — 工具重定位（已写完详细计划，等 Codex 执行）
- **下一**：Sprint 02 — 健壮性
- **待补**：Sprint 03/04/05 详细计划（前一 sprint 进入 review 时再写）

详见 [`current.md`](../current.md) 与 [`SPRINT_01_drop_financial_extraction.md`](SPRINT_01_drop_financial_extraction.md)。

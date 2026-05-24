# 项目复盘：Filings Atlas / 全球披露图谱（原 FS Capture）v0.2+

> **作者**：Eric Zhang（kaiyu199602@gmail.com）
> **日期**：2026 年 4 月底~5 月初（与 VBA Captor v1.0 release 并行收尾）
> **AI 协作**：Claude Code（planner / reviewer）+ Codex（generator）
> **当前阶段**：v0.1（混合工具）→ v0.2（重定位为纯 PDF 抓取）
> **commit 节奏**：1 个 bootstrap commit + 持续 patch；v0.2 起进入 Sprint planner-codex 循环

---

## 1. 项目目标（v0.2 重定位）

构建一个 Windows EXE 桌面工具，让财务/审计专业人士在 1-3 分钟内完成 **A 股 / 港股 / 美股 / 韩股 4 市场上市公司的官方披露 PDF 批量下载**——年报、审计报告、季报、半年报、IPO 招股书。

**核心交付**：

- 4 市场 plugin-per-exchange 架构（akshare + cninfo + HKEXnews + SEC EDGAR + DART 混合数据源）
- PySide6 自定义无边框圆角 GUI（4 chip 选所、行内确认、进度面板）
- 一键并发抓取（QThreadPool + tenacity 重试 + per-source TokenBucket 限速）
- PDF 文件名内嵌元信息（`{exchange}_{code}_{year}_{period_type}_{kind}.pdf`），扁平铺在 `output/`
- PyInstaller one-folder 打包（~340 MB，EXE 27 MB 启动器）

**v0.1 → v0.2 重定位**：

- v0.1 试图把「PDF + 三大报表数字 + Excel 底稿」三合一
- 实际跑下来发现：抓数据 + 字段口径统一 + 底稿装填的工程量与维护量都很大，跨数据源命名（akshare 英文 EM / SEC GAAP / 港股中文）对齐困难
- v0.2 重定位：FS Capture 只做"批量拿到原始 PDF"这一件事；抓数字 + Excel 底稿装填的需求请用 `VBA Captor/`（已 v1.0 release）

**跟 `VBA Captor/` 的关系**：

- VBA Captor v1.0 是 4 天 vibe coding 跑出来的 Excel/VBA 工具（10K 行 VBA + 12 个 phase）。
- FS Capture 是用户在 VBA Captor 之外的并行品——「双击 EXE 即用 + 下载官方 PDF + 现代 GUI」。
- 两者数据源有大量重叠（A 股新浪/akshare、美股 SEC、港股 HKEXnews），定位互补：VBA Captor = 数据，FS Capture = 文件。

---

## 2. 开发过程

跟 VBA Captor 的 12 phase 节奏不同，FS Capture v0.1 是**单 commit bootstrap + 持续 patch** 模式；v0.2 起改用 **roadmap 驱动的 Sprint 循环**。

### Day 0（背景）

- 用户原话："想要开发一个一键抓取 A股/港股/美股/韩股上市公司**审计报告和年报（公司自己发布的）**的工具"。
- 提交了既有 VBA 工具 `新浪财经行业数据查询V2.2.xlsm` 作为参考。
- 通过 AskUserQuestion 锁定：
  - GUI = PySide6 + 自定义 QSS
  - DART Key = 用户自行注册并粘贴
  - 多期间 = 年份区间 + 期间类型 checkbox（年报 / 季报 / 半年报 / IPO 招股书）
  - 审计报告 = 完整年报 PDF + 独立审计报告 PDF（如有）

### v0.1 Bootstrap commit（端到端实现）

单 commit 落地（含财务数据抓取 + Excel 底稿——后来在 v0.2 全部移除）：

- `app/core/`：models / settings / http / ratelimit / cache / orchestrator
- `app/ui/`：无边框圆角窗口 + 4 chip + ticker_row 行内确认 + 进度面板 + 设置对话框
- `plugins/{ashare,hk,us,kr}/`：4 市场各 3 文件
- `fs_capture.spec` + `build.bat` + `run.bat`
- `pyproject.toml` + `requirements.txt` + `config.example.toml`
- `tests/`：2 个 unittest 文件
- `DEVELOPMENT_BRIEF.md`：给 Codex 的代码审核委托书

### v0.1 持续 patch

按时间顺序的关键修复：

1. **SEC 字段名错**：网上文档普遍说 `periodOfReport`，实际接口字段是 `reportDate` → 全美股 10-K 找不到，已修。
2. **HKEXnews servlet 已 404**：旧 `titleSearchServlet` URL 退役，改 `titlesearch.xhtml`。
3. **akshare HK 聚合接口不稳**：`stock_hk_spot_em` 高频 `RemoteDisconnected`，改单股 `eastmoney stock/get` 端点。
4. **PyInstaller windowed 两个坑**：
   - `sys.stderr is None` → loguru `add(...)` `TypeError`（`app/main.py:16-25` 加 None 守护）
   - `sys.stdout is None` → akshare 内部 tqdm `.write()` crash（`app/main.py:8-14` 重定向 `os.devnull`）
5. **Orchestrator 跨 job 信号叠加**：连续 submit_job 时旧回调没断 → 加 `task_id set` 隔离 + emit 前 disconnect。
6. **报告输出布局扁平化**：用户反馈不要按公司/市场嵌套 → `output_paths.report_output_path` 把所有 metadata 写进文件名，平铺在 `output/` 下，并加 `test_output_layout.py` 锁定。

### v0.2 重定位决策

用户在 v0.1 之后审视工具定位：

- 抓数据 + 字段中文化 + 底稿装填的边际收益不高（VBA Captor 已稳定覆盖此场景）
- 维护多套字段映射表（akshare EM / SEC GAAP / 港股中文 / 韩文）的人工成本超过预期
- "拿到原始 PDF"是用户在 VBA Captor 之外没解决的痛点，FS Capture 的差异化价值在这

→ 决定砍掉财务数据抓取 + Excel 输出整条链路，工具收敛到纯 PDF 下载。Sprint 01 执行重定位代码清理。

### 当前状态

- **已端到端验证**：A 股 600519（茅台）/ 港股 00700（腾讯）/ 美股 AAPL（10-K FY2023/2024）
- **未实跑**：韩股（用户尚未注册 DART API Key）
- **测试覆盖**：单元测试覆盖 PDF 文件名规则 + US/KR 报告选片
- **打包**：PyInstaller 产物 ~340 MB，已在用户机器双击启动验证

---

## 3. 跟 VBA Captor 的协作模式差异

VBA Captor 用了 Triangle 模式（Planner / Generator / Reviewer 循环 12 phase）。FS Capture v0.1 简化为单 commit bootstrap；v0.2 起改用：

```text
        Claude Code (planner)
              ↓ 写 roadmap/SPRINT_NN_xxx.md（详细计划）+ current.md（一页交互卡）
              ↓
        Codex (generator)
              ↓ 端到端实施 sprint + 在 current.md 末尾写"## Codex 回报"
              ↓
        Claude Code (reviewer)
              ↓ 跑 frozen 回归 + spot-check + 在 current.md 末尾写"## Reviewer 验收"
              ↓
        commit & 用户测试 → 进下一 sprint
```

为什么 v0.2 重新引入 Sprint：

- v0.1 单 commit bootstrap 适合"从 0 到 1"
- v0.2 重定位需要砍代码、改 ABC、改测试，影响面大且需要 reviewer 验证 frozen API 不退化
- Sprint 节奏 + `current.md` 一页交互 + roadmap/ 详细计划，能在保证可审计性的同时不像 VBA Captor 那样重型

### 关键技巧

**roadmap 文件结构**：

- `roadmap/ROADMAP.md`：5 sprint 整体路线图（高层）
- `roadmap/SPRINT_NN_xxx.md`：每个 sprint 详细计划（背景 / 涉及范围 / 严禁动 / 步骤 / 验证 / 触发 Planner / 回报格式 / 验收清单）
- `current.md`：一页交互卡（当前 sprint hand-off + Codex 回报 + Reviewer 验收，滚动覆写）

**SPRINT_NN_xxx.md 模板**：

```text
1. 背景（为什么做这个 sprint）
2. 涉及范围（待删除文件 / 待修改文件 / 严禁动）
3. 实施步骤（按顺序，每 step 给具体 spec）
4. 验证（单元测试 + 端到端 + frozen 回归 + grep 兜底）
5. 联系 Planner 触发条件
6. Codex 完成回报格式
7. Reviewer 验收清单
```

**current.md 一页约束**：

- 任务概要（1-2 段）
- Codex 待办（按顺序的 checkbox 列表）
- 验证（自检命令）
- 严禁动
- 联系 Planner 触发条件
- "## Codex 回报"段（待填）
- "## Reviewer 验收"段（待填）

约 60-80 行，方便用户/Codex/Reviewer 一眼看到当前轮次状态。

---

## 4. 最容易出错的地方

### 4.1 PyInstaller windowed 模式 stdio = None（已修两次）

- 默认 `sys.stderr` / `sys.stdout` 在 windowed 模式下都是 `None`
- loguru `logger.add(sys.stderr)` 会 `TypeError`（None 不是文件）
- akshare / tqdm 进度条 `.write()` 会 `AttributeError`（None 没有 `.write` 方法）
- 修复：`app/main.py:8-25` 在最顶部把 None 流重定向到 `os.devnull`

教训：windowed PyInstaller **永远先想 stdio 是 None 的可能**，特别是依赖会偷偷打日志或进度条的库。

### 4.2 SEC 字段名（文档 vs 实际）

- 网上文档 / 知乎答案普遍说 `periodOfReport`
- 实际 SEC submissions API 字段是 `reportDate`
- 错了之后症状是"全美股 10-K 都找不到"，看不出 bug 在字段名而不是逻辑
- 教训：**HTTP API 必须实跑 + 断言关键字段**，不能照抄文档

### 4.3 港股没官方 API

- HKEXnews 只有 HTML 表单 → selectolax 解析（已踩过 servlet URL 404 的坑）
- akshare 港股聚合接口（`stock_hk_spot_em`）`RemoteDisconnected` 高频 → 改单股 `eastmoney stock/get`
- **HK 选片仍宽松**：`_select_main` 仅按"标题含年份字符串"筛选，同一年份的多份关联文件（补充公告、ESG 报告）可能误选；财年 vs 公告年份混淆（汇丰、阿里巴巴非 12 月财年没专门处理）
- 缓解但未根治：Sprint 03 计划补 PDF 内嵌文本验证 + 财年元数据识别

### 4.4 OpenDartReader CWD 副作用（待解决）

- `dart.document(rcept_no)` 把文件下载到**当前工作目录**（库默认行为）
- 当前实现用 `os.chdir(target_dir)` 临时切，下载完切回
- **QThreadPool 默认 4 worker，多个 KR 任务并行会互相覆盖 CWD** → 进程级状态不是线程级
- 修复方向（Sprint 02）：monkey-patch `OpenDartReader.document` 把 cwd 参数化，或 fork 一份代码

教训：**绝不在工作线程改进程级状态**（cwd / signal handler / env vars / 全局 logger 配置）。

### 4.5 SEC submissions 分页未实测

- AAPL FY2023/2024 都在 `recent.*` 数组（最近 ~1000 条）
- 老 ticker（如 AAPL 高频 8-K 之后的老年份、MSFT FY2018 之前）会落到 `filings.files[]` 分页路径
- 这条 fallback 代码已写但**未实测**
- 测试建议（Sprint 05）：用 MSFT 跑 FY2018，强制走分页

### 4.6 取消按钮假提示

- `progress_dock` 定义了 `cancel_requested` 信号
- 但**没有任何 slot 接收**，并且 `_TaskRunnable.run` 没有 cancellation token
- 用户点取消按钮后无反应，是 UI 假提示
- 修复方向（Sprint 02）：补 `threading.Event` cancel token 让 task 主动 break

### 4.7 v0.1 范围蔓延（导致 v0.2 重定位）

- 用户原始诉求是"抓审计报告和年报"，重点在 PDF
- v0.1 顺手把"抓三大报表数字 + Excel 底稿"也实现了
- 实际维护后发现：跨数据源字段口径统一是个无底洞，与 VBA Captor 范围重叠
- v0.2 砍掉数据 + Excel，回归 PDF-only 定位

教训：**面对"顺手做"的功能要警惕**。能扩张到 N 个市场的工具，附带功能的维护成本会被 N 倍放大；先把核心做扎实再考虑外延。

---

## 5. 技术债（v0.2 起）

| 项 | 说明 | 严重性 | 计划 sprint |
|---|---|---|---|
| KR 整条链路从未实跑 | DART Key 待用户注册后实测 | **高** | 无（待用户注册） |
| HK 报告选片仅按年份字符串 | 同一年份多份补充公告可能误选 | 中 | Sprint 03 |
| `cancel_requested` 信号无接收 + 无 cancel token | UI 假提示，task 无法中断 | 中 | Sprint 02 |
| KR `_download_filing` 用 `os.chdir` | 多 worker 并发下不安全 | 中 | Sprint 02 |
| diskcache 单例从未关闭 | 退出时可能留 lock 文件 | 低 | Sprint 02 |
| `httpx.stream` 全局 30s 超时 | 大 PDF + 慢网下会 timeout | 低 | Sprint 02 |
| `name_label` `QLabel` 不可选中 | 审计场景需复制公司全名 | 低 | Sprint 02 |
| 首次启动无 onboarding | 用户没引导 | 低 | Sprint 04 |
| 北交所代码识别不全（cninfo 用 `bj` 我们用 `bse`） | 北交所 ticker 失败 | 低 | Sprint 04 |
| KR Q3 vs Q1 关键词撞车 | DART 季度报告类型相同 | 低 | Sprint 04 |
| IPO 招股书链路完整性 | 4 市场暂只有部分实现 | 低 | Sprint 04 |
| frameless 多显示器 / 不同 DPI 未测 | 边缘 hit-test 边界条件 | 低 | 无（用户当前单屏） |

---

## 6. 下一阶段应避免什么

### 6.1 范围

- ❌ 不要再把"抓数据"或"做底稿"加回来。这俩在 v0.1 试过了，跨数据源字段口径维护成本太高，VBA Captor 已覆盖。
- ❌ 不要扩张到非"批量下载 PDF"以外的功能（如 PDF 分析、关键词提取、AI 总结等）；这些都是独立工具的范围。

### 6.2 实现阶段

- ❌ 不要在工作线程改进程级状态（cwd / signal handler / env vars / 全局 logger）
- ❌ 不要假设 windowed PyInstaller 的 sys.stdin/stdout/stderr 可用
- ❌ 不要 `collect_all` 一个库就以为 hidden imports 全了（akshare/OpenDartReader 在冻结模式下都需要冷启动验证）

### 6.3 验证阶段

- ❌ 不要只测 happy path：必须测异常路径（4xx / RemoteDisconnected / 老 ticker 走分页）
- ❌ 不要让 EXE 默认在 Program Files 路径下写 `config.toml`（受限路径写不下，要给降级到 `%APPDATA%`）
- ❌ 不要在没冷启动测过的状态下宣称"已端到端"（GUI 启动、首次 cache 加载、4 plugin 全跑通才算）

### 6.4 协作阶段

- ❌ 不要把 sprint plan 写完就让 Codex 一次写到底——critical 修复后要重新审一遍 frozen 区
- ❌ 不要让 Codex 自由发挥 GUI 视觉——给目标截图 + 像素级 spec
- ❌ 不要让 `current.md` 超过一页——失去"一眼看到当前轮次状态"的意义

---

## 7. 后续模板化经验

### 7.1 Sprint 节奏与 v0.1 bootstrap 的差异

- v0.1 bootstrap 适合从零起、用户对"快速看到能跑的 demo"需求高
- Sprint 模式适合：
  - 工具进入维护期，需要保 frozen API 不退化
  - 改动影响面大需要 reviewer 验证
  - 多人协作（even if "Claude vs Codex"）需要明确同步点

### 7.2 SPRINT_NN_xxx.md 章节模板（FS Capture v0.2）

```text
1. 背景（为什么做这个 sprint，跟前/后 sprint 关系）
2. 涉及范围（待删除文件 / 待修改文件 / 严禁动）
3. 实施步骤（按顺序，每 step 给具体 spec）
4. 验证（单元测试 + 端到端 + frozen 回归 + grep 兜底）
5. 联系 Planner 触发条件（明确 stop-before-commit 边界）
6. Codex 完成回报格式（commit hash + 文件 delta + 测试结果 + 决策点）
7. Reviewer 验收清单（checkbox 形式，方便逐项确认）
```

### 7.3 PyInstaller 项目必须做的冷启动测试矩阵

```text
1. 双击 EXE → 窗口正常打开（不闪退）
2. 首次启动各 plugin name_resolver 都能完成 cache 加载
3. config.toml 在 Path(sys.executable).parent 可写
4. crash.log 在崩溃路径下确实写出来 + MessageBox 弹出来
5. windowed 模式下任意依赖打日志/进度条不应 crash
6. logs/ output/ cache/ 目录在不存在时被自动创建
```

### 7.4 测试金字塔（Python 项目）

```text
        ▲   端到端冒烟（每市场 1 个真 ticker，看完整 PDF 下载链路）
       ▲▲   集成测试（mock HTTP，跑 plugin 单步：name_resolver / reports）
      ▲▲▲  单元测试（选片逻辑 / 文件名规则，不发 HTTP）
```

当前覆盖：第 3 层若干 unittest 方法。第 1 层手工跑过 3 个市场。第 2 层缺，Sprint 05 要补 mock-based plugin 集成测试。

### 7.5 用户视角 doc

跟 VBA Captor 教训一致：

- 用户面向 README 必须先写"双击 .exe → 选交易所 → 填代码 → 点确认 → 选年份 → 点开始抓取"5 步流程
- 技术黑话（"WinHttp / cache / fallback / DART OpenAPI / XBRL"）要翻译成业务话
- 韩股 DART Key 必须给注册引导（写在 README 韩股说明段，不要塞到设置对话框 tooltip）
- 工具定位边界要写清楚：FS Capture 是 PDF 下载工具，不做数字 / 不做底稿

---

## 8. 总结（v0.2 时点快照，2026-05-09）

> 以下评估反映 v0.2 重定位完成时的状态；v0.5 ↔ v0.6 后续进展见 §9。

| 维度 | 评估 |
|---|---|
| 范围 | 4 市场 PDF 下载（A/HK/US 端到端验证 + KR 编码完成待 DART Key） |
| 时间 | v0.1 单 commit bootstrap + 持续 patch；v0.2 起 Sprint 循环 |
| 代码量 | v0.1 ~3,900 行 Python；v0.2 砍掉抓数据后预计 ~3,000 行 |
| 测试 | 单元覆盖 PDF 文件名规则 + US/KR 报告选片；集成 / e2e 待 Sprint 05 |
| GUI 质量 | 无边框圆角 + 4 chip + 行内确认 + 进度面板（首版可用，无 onboarding） |
| 数据源稳定性 | A 股 / 港股 / 美股 PDF 抓取已验证；HK 选片有已知粗糙；KR 未实跑 |
| 部署 | PyInstaller one-folder ~340 MB / 启动器 ~27 MB，windowed 模式两个坑已修 |
| 文档 | ARCHITECTURE + PROJECT_RETROSPECTIVE + README + DEVELOPMENT_BRIEF；v0.2 阶段的 sprint 规划与 planner↔Codex 协作日志已归档于 `roadmap/archive/` |

**当前阶段：v0.2 重定位中（Sprint 01 砍掉抓数据代码 → Sprint 02-05 健壮性 + 选片 + UX + 测试）**。

后续优先级（高 → 低）：

1. Sprint 01 — 工具重定位代码清理（移除 fetch_financials / FinancialStatement / excel_writer）
2. Sprint 02 — 健壮性（KR cwd / cancel token / cache close / httpx timeout / name_label）
3. Sprint 03 — HK 选片精确化（pypdf 文本验证 + 财年识别）
4. Sprint 04 — UX 打磨（onboarding + 北交所 + KR Q1/Q3 + IPO 招股书 + PDF sidecar 元数据）
5. Sprint 05 — 测试基建升级（pytest + 集成测试 + e2e smoke + CI）
6. KR 整条链路实跑（用户注册 DART Key 后）

— Eric Zhang（kaiyu199602@gmail.com），2026-05-09

---

## 9. 后续：v0.6 新增台股（2026-05-17）

v0.5 5 个 sprint 完成后工具趋于稳定，按用户「再加一个市场」需求新增台股插件。这是工具发布后的第一个市场扩展，验证插件架构的可扩展性。

### 9.1 数据源调研

台股披露生态由 **MOPS 公開資訊觀測站** 统一管理，单一端点 `doc.twse.com.tw/server-java/t57sb01` 即可覆盖：

- 上市（TWSE）+ 上柜（TPEx）所有公司（同一端点）
- 年报（mtype=F 股東會材料，typecode F04 中文 / FE4 英文）
- 季报、半年报（mtype=A 財務報告書，typecode AI1 合併中文 / AIA 個體）
- IPO 公开说明书 + 公司債 + 特別股发行说明书（mtype=B）

无须 API Key，与 A/HK/US 三市场一致；用户体验上无配置门槛。

### 9.2 实现教训

1. **ROC 年份转换是经典 footgun**。`roc_year = ad_year - 1911`，独立函数 + 单元测试覆盖。
2. **文件名语法因 mtype 而异**，且与"年份"语义不一致：
   - mtype=F 是「股东会年份」（meeting year），文件名 leading YYYY 才是 FY
   - mtype=A 是 FY 本身，文件名 prefix `YYYYQQ` 编码季度
3. **下载需两步**：POST step=9 返回一个 `<a href='/pdf/<file>_<timestamp>.pdf'>` 临时链接（session-bound 数分钟），再 GET 才拿到 PDF 字节。
4. **TWSE 证书 hygiene 缺陷**：服务端证书缺 Subject Key Identifier 扩展，Python 3.12+ OpenSSL 严格模式拒绝。仅对 `source=twse` 用 `verify=False`（用户授权），不污染其他四市场严格校验。
5. **Big5 编码**：MOPS HTML body 是 Big5，必须 `resp.content.decode("big5", errors="replace")` 而非 `resp.text`。

### 9.3 验证

| 报告类型 | 实跑结果 |
|---|---|
| 2024 中文年报（F04） | ✓ 9.99 MB |
| 2024 半年报（AI1 Q2） | ✓ 7.2 MB |
| IPO + 公司債 + 特別股说明书 | ✓ 60+ 份（2-16 MB 不等） |

附带修复：IPO 输出文件名 `.pdf.pdf` 重复后缀小 bug。

### 9.4 插件架构验证

新市场代码量统计（不含测试）：

| 文件 | 行数 |
|---|---|
| `plugins/tw/__init__.py` | 28 |
| `plugins/tw/name_resolver.py` | ~140 |
| `plugins/tw/reports.py` | ~320 |
| UI 改动（5 个文件，每个加一条 dict entry / 列表元素） | ~10 |

**结论**：plugin-per-exchange 架构按预期可扩展，新市场约半天到一天就能完成（含数据源探查、实现、测试、E2E 验证）。

— Eric Zhang，2026-05-17（v0.6 5 市场 release）

---

## 10. v0.6.1 postscript（2026-05-23）

v0.6.1 是 v0.6 发布后的 patch sprint，目标是修复确认 bug，不扩功能、不改架构。

- 限流：`RateLimiterRegistry.get` 支持已存在 bucket 的速率热更新，设置保存后无需重启。
- 港股：PDF 内容验证通过时只在 verified 候选中选片；全失败才回退标题/日期打分，并记录 warning。
- 港股财年：补充多只非 12 月财年公司映射，包括新鸿基地产、领展、波司登、阿里健康、阿里影业、新东方和东方甄选。
- A 股：akshare code/name 与 item/value 映射先用 `zip(strict=True)` 诊断脏数据，异常时回退旧路径。
- 韩股：确认并保留 DART 官方字段 `induty_code`，补测试锁定行业字段读取。
- UI：批量导入现在会提示整行未识别的 token，同时避免把同一行里的公司名称当作错误。

— Codex，2026-05-23（v0.6.1 patch）

---

## 11. v0.7 postscript（2026-05-23）

v0.7 完成韩股去 Key 化，同时补齐 US 分页测试和 TW e2e 覆盖；仍保持 PDF-only 定位，不发 GitHub release。

- KR 双模式：`dart.api_key` 保留为可选加速器；有 Key 优先走 OpenDartReader，无 Key 自动走 DART 公网披露页。
- 轻量公网客户端：未引入 `dart-fss`，新增 `plugins/kr/dart_web.py`，选择器集中在 `_SELECTORS`。
- 按需 corp_code：从旧的全量 `_load_map()` 改为 `resolve_one()` 单票查询，缓存 key 为 `kr:corp:{code}` 且不设 TTL。
- 公网列表：`list_filings` 输出保持 `rcept_no/report_nm/rcept_dt/corp_code`，下游选片和 PDF 下载链路无需改动。
- 真实 smoke：无 Key 跑通 `013890`、`005930`、`000660`、`454910` IPO、`035420`；有 Key 回归跑通 `005930` 年报。
- 测试补强：新增 KR 公网解析测试和 US `submissions.files[]` 分页 fallback 测试；TW e2e 已在 `tests/e2e/test_smoke.py` 覆盖。

— Codex，2026-05-23（v0.7 internal）

---

## 12. v0.8 postscript（2026-05-23）

v0.8 是性能与清债 sprint，仍保持 PDF-only 定位，不发 GitHub release、不打 tag、不更新 README 顶部版本号。

- Playwright：`render_url_to_pdf` 改为进程级 singleton browser pool，首次渲染 lazy init，每次渲染只开独立 context，并在应用退出时关闭 runtime/browser。
- 下载续传：`stream_to_file` 支持 `.part` + `Range` 断点续传，服务器忽略 Range 时从头重写，416 时清理 stale `.part` 后重试，瞬断失败时保留 `.part`。
- 名称解析：新增 `cached_or_load` 双重检查锁，A 股 / 美股 / 韩股 / 台股 name resolver 接入 single-flight；港股保持单股查询路径。
- UI 文案：11 个顶层 UI 模块的中文字符串集中到 `app/ui/strings.py`，仅做物理集中，不引入 `tr()` 包装。
- Lint 锁定：新增 pre-commit 配置、GitHub Actions lint/test job，并补 ruff isort first-party 配置。
- 验证：`pytest -m "not e2e" -v` 为 100 passed / 7 deselected；`ruff check development/` 通过；KR 005930 audit e2e smoke 通过。

— Codex，2026-05-23（v0.8 internal）

---

## 13. v0.9 postscript（2026-05-23）

v0.9 是 Filings Atlas / 全球披露图谱的内部迭代版。该版本完成可见品牌重命名、双语 UI、sidecar 迁移、增量更新、JP/UK 新市场和发布工程收口；仍严格保持 PDF-only 定位。

- 品牌：用户可见名称从 FS Capture 切换为 Filings Atlas / 全球披露图谱，仓库目录名保留不动。
- 双语：UI 支持中文 / English 运行时切换，设置持久化，TickerRow 输入在切语言后保留。
- 元数据：sidecar 从 `output/` 迁移到 `data/cache/sidecars/`，`output/` 回到只放 PDF。
- 增量：新增独立“增量更新 / Incremental Update”按钮，按 sidecar 判断已完成任务。
- 日股：新增 EDINET 插件。当前构建强烈推荐用户配置 EDINET Subscription-Key；无 Key 时官方 API 会返回 invalid subscription key，仅能尝试公网 fallback。
- 英股：新增 FCA NSM 插件。UK 不需要 Key；年报通常在次年披露，选择逻辑用 `document_date/headline` 回筛报告年度。
- 发布工程：README 改为双语，新增 CHANGELOG 和 GitHub Actions release workflow 准备，版本号调整为 `0.9.0`，公开 release 产物留到下一轮。

验证基线：`pytest -m "not e2e" -v` 为 145 passed / 7 deselected；`ruff check . --no-cache` 全绿；UK `ULVR` / `HSBA` / `AZN` 真实 smoke 均落地 `%PDF`。

— Codex，2026-05-23（v0.9 internal）

---

## 14. v1.0 postscript（2026-05-24）

v1.0 是 Filings Atlas / 全球披露图谱的首个 GitHub release 版本。该版本在 v0.9 的 7 市场基础上新增新加坡 SGX，并完成一轮以真实 benchmark 为边界的性能优化和零体验成本体积优化；PDF-only 定位保持不变。

### 14.1 SG plugin 实施

- SGXNet 当前可用端点不是 sprint 草案里的 `announcements/v1.0/search`，而是网页配置里的 `announcements/v1.1/`、`financialreports/v1.0` 与 `ipoprospectus/v1.0/`。
- API 访问需要 `User-Agent`、`Referer` 和由公开 config/CMS `qrValidator` ROT13 生成的 `authorizationToken`；不需要 Cookie 或登录态。
- 年报 `financialreports.title` 字面值为 `Annual Report`；半年报按 `category_name = "Financial Statements"` 与标题中的 `Half Yearly Results` 识别；IPO 来自 `ipoprospectus` rows。
- SGX API 返回的 `url` 多数是公告/IPO HTML 页面，不是 PDF 直链；实现需要二跳解析 `.pdf` 或 `FileOpen/*.ashx?App=IPO&FileID=...`。
- Playwright 不需要。Spike 发现 SGX 页面级 headless 访问可能遇到 Access Denied，但 HTTP API + HTML 附件解析能稳定拿到 PDF。
- 限流从 `sgxnet = 1.5` QPS 起步；spike 的 10 次连续 count 请求（0.5s 间隔）未触发 429/403。

### 14.2 性能优化

| 子项 | 预期 | 实测 | 状态 |
|---|---|---|---|
| 5.1 并发 4 → 6 | 批量总时长下降 | 批量 benchmark 从 211.76s 降到 94.22s（综合 5.x 后） | 保留 |
| 5.2 name_resolver 预热 | 首次确认更顺滑 | A/TW/US 全量映射后台预热，单元测试锁定；SG 保持按需查询 | 保留 |
| 5.3 chunk 64 KB → 256 KB | 大 PDF 下载更快 | 批量 benchmark 未见负收益；UI 进度粒度变粗但可接受 | 保留 |
| 5.4 Playwright context 池化 | KR/UK/US HTML 渲染更稳 | v0.9 中 UK AZN zero-report 消失，批量任务 0 failure | 保留 |
| 5.5 HTTP 连接池调优 | 多源下载更稳 | `httpx.Limits(100, 40)` 已锁测试 | 保留 |
| 5.6 sidecar 原子写 | 元数据写入更安全 | temp-file replace 单测通过 | 保留 |

性能结论：以 21 任务 A/HK/US/KR/TW/UK/SG benchmark 为准，总耗时下降 55.5%，达到 sprint “至少 -15%” 的门槛。单票公网下载波动较大，不作为回滚判断的唯一依据。

### 14.3 体积压缩

- 实际减重：449.0 MB → 441.4 MB，减少 7.6 MB。
- 有效项主要是 PySide6 翻译裁剪：96 个 `.qm` 文件裁到 6 个，PySide6 目录约减少 6.2 MB。
- `numpy.tests`、`pandas.tests`、`pytest` 等 excludes 已补齐，但当前 bundle 本来没有打入这些目录，所以实际减重有限。
- Playwright selective collect 跳过：离线 one-folder 运行仍依赖 Playwright driver，贸然过滤可能破坏 KR/UK/US HTML 渲染。
- UPX 跳过：本机未安装 `upx`，且 release 阶段不引入可能触发 Defender 误报的新变量。

### 14.4 不做的项（明确）

- 不用 WebView2 替代 Chromium：会影响 KR/UK/US HTML-to-PDF 渲染路径，未来单独评估。
- 不替换 akshare：会引入更高维护成本。
- 不为 SG 增加 Key 配置：当前 SGXNet 公网路径不需要 Key。
- 不扩大到财务数据抓取、指标计算或 Excel 输出。

验证基线：非 e2e `170 passed / 22 deselected`，benchmark opt-in `11 passed`，ruff 全绿；SG smoke 已覆盖 DBS/UOB/Singtel 年报、UOB 半年报与 `3407` IPO 招股书。

— Codex，2026-05-24（v1.0 release candidate）

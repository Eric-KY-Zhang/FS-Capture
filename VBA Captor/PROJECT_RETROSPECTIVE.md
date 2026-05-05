# 项目复盘:上市公司财务数据查询 v1.0

> **作者**: Eric Zhang(214978902@qq.com)
> **日期**: 2026-05-02 ~ 2026-05-05(4 天)
> **AI 协作**: Claude Code(planner/reviewer)+ Codex(generator)+ GPT 5.5 Pro(static auditor)
> **代码量**: VBA ~10,000 行 + Python ~2,500 行 + 文档 ~5,000 行
> **commit 数**: ~40+ commits / 12 个 phase(4b14 → 4n)

---

## 1. 项目目标

构建一个 Excel + VBA 桌面工具,让财务/审计专业人士在 1-3 分钟内完成 **A 股 / 美股 / 港股 / 韩股 4 市场上市公司同期可比财报抓取 + 跨市场对标**,替代手工复制粘贴 + 工作日多次切换数据源的工作流。

**核心交付**:
- 4 市场抓数(新浪 + EDGAR + 雪球 + stockanalysis 混合数据源)
- 18 项标准指标统一口径
- 跨市场指标合表(公式 cell-ref 实时联动)
- B6 toggle:原币 ↔ 统一 RMB 实时切换(汇率 UDF 实时刷新)
- 磁盘缓存 + retry/backoff + 数据质量自动 QA
- 8 个离线测试 + 7 张 frozen 回归驱动

---

## 2. 4 天 vibe coding 开发过程

### Day 0(项目背景)

- 起点是 V2.2 单市场(A 股新浪)的 VBA 工具,~20 公司容量
- 用户需求:扩展到 4 市场 + 加跨市场对比 + 给财务审计专业人士看的数据/UI 质量

### Day 1(2026-05-03)

**Phase 4b14 / 4c / 4d / 4e** — 多市场扩张

- **4b14**:美股 EDGAR + 雪球 fallback 美股财报字段映射开源化(POM/HTT/BABA 等 20-F/中概股 hardcode → mapping JSON)
- **4c**:港股从雪球 HK API 接入,字段命中诊断 sheet 落地
- **4d**:韩股 stockanalysis.com KRX HTML 表格抓取(无 cookie 路径)
- **4e**:UX 第一轮(诊断 sheet 默认隐藏 + 样本池四市场分栏)

### Day 2(2026-05-04)

**Phase 4f / 4g / 4h / 4i / 4j** — 主线 + 抛光

- **4f**:RMB 换算 hook(汇率 sheet + B6 toggle + GetFxRate scaffold,WriteWideTable 接 RMB 换算 hook)
- **4g**:跨市场指标合表 + hide-tab 按钮 + POOL_DATA_START_ROW 迁移 10→11
- **4h**:5 件并行(B6 实时 toggle / 磁盘缓存 / stockanalysis 中概美股 fallback / 跨市场 BS-IS-CF / 4g 收账)— 单 phase 最大爆发(~14h Codex)
- **4i / 4i.1 / 4i.2**:UX 抛光(样本池布局 / 使用说明商务化 / 汇率说明区 / 单表按钮删除 / fallback 自动化)
- **4j / 4j.1**:简化跨市场对比策略(只留指标表,删 BS/IS/CF 合表)+ 抑制合并弹窗

### Day 3(2026-05-05 上午)

**Phase 4j.2 / 4j.3 / 4j.4** — 样本池视觉 1:1 还原

- 用户给目标 UI 截图,planner 写 plan 让 Codex 1:1 还原
- 第一版差距大 → 反复 patch(padding cols 缩窄 → 隐藏)
- 最终 4 张 brand-color 卡片 + 工具栏 + 使用提示 panel

### Day 3(2026-05-05 下午)

**Phase 4k / 4l / 4m / 4n** — 优化 4 sprints

GPT 5.5 Pro 静态审阅给 14 项 backlog,reviewer 选 12 项做(2 项 defer):

- **4k**:数据准确性(FX missing 不再 fallback 1)+ live FX UDF + AppStateGuard(1 入口示范)+ KR Score 日期化修复
- **4l**:HTTP/cache 诊断遥测(诊断 sheet 11→17 列)+ retry/backoff/SEC 限流 + CleanReleaseWorkbook 宏
- **4m**:8 个离线 fixture 测试 + 数据质量 QA(BS 平衡 / FX missing / 关键字段)+ cache 分源 TTL
- **4n**:AppStateGuard 扩到剩余 5 入口 + ARCHITECTURE.md + release notes

### Day 4(2026-05-05 晚)

**收尾** — v1.0 release 版构建 + 项目复盘 + 风险登记表 + 文件夹整理。

---

## 3. 最有效的 AI 协作方式

### 3.1 Triangle 模式(实测有效)

```
        Claude Code (planner)
              ↓ 写 PHASE_*.md plan + Codex prompt
              ↓
        Codex (generator)
              ↓ 端到端实现 + 装表 + 跑回归
              ↓
        Claude Code (reviewer)
              ↓ 独立验证 + 给 verdict
              ↓
        commit & next phase
```

每个 phase 的循环:
1. **Planner** 写 PHASE_*.md(目标 / 子任务 / 严禁动 / 触发 Planner 条件),起 Codex 端到端 prompt
2. **Codex** 跑端到端,1 个 commit + 自报 commit hash + 回归结果
3. **Reviewer** 独立跑 frozen 回归 + spot-check 关键 code + 给 ✅/⚠️ verdict
4. 循环

### 3.2 关键技巧

**Plan 写作**:
- "严禁动" 列表明确告诉 Codex 哪些 frozen,避免 scope creep
- "联系 Planner 触发条件" 列出 Codex 应该 stop-before-commit 的边界(eg "回归任一退化")
- "state-bound inspect 同步规则" — 区分真 frozen(`test_fx_live.py`)vs state-bound(`inspect_phase4*_state.py`)
- 每个 step 给具体 code template / cell range / 颜色 hex,Codex 不需要自由发挥

**Codex prompt**:
- 项目语境 anchor(强调本期是 X 已有数据源的延续,降低安全分类器误报)
- 必跑回归列表(任一退化立即停下)
- 完成后报告格式(commit hash + 实测数据 + 截图行号)

**Reviewer 验证**:
- 不光看 Codex 自报,**独立跑 4-8 张 frozen 驱动**确认
- spot-check 关键代码段(eg Step 4 的 UDF 公式 / Step 1 的 NumberFormat)
- 看实测数据(eg FX UDF 0.13s 远超 5s budget,SEC 限流 468ms 满足 110ms 要求)

### 3.3 GPT 5.5 Pro 作 static auditor

- 一次性把整个 codebase + xlsm 给 GPT 5.5 Pro 静态审阅
- 输出 14 项 backlog 排序(P0/P1/P2)+ 每项 实施细节 + 验收标准
- Reviewer(我)逐条评估同意/部分同意/不同意,选 ROI 最高的进 sprint
- **GPT 5.5 Pro 的优势**:不疲劳,会全面扫描;**reviewer 的优势**:知道项目实际场景,能砍掉 over-engineering

---

## 4. 最容易出错的地方

### 4.1 Excel COM 卡死(高频)

- `Range("E5").Value = ""` 在合并区域(E5:M5)会触发模态弹窗 → COM 调用 hang 永不返回
- 缓解:`Application.DisplayAlerts = False` + 先 UnMerge 再写值 + Re-Merge
- Phase 4j.1 / 4n release_v1_clean 都中过这个坑

### 4.2 win32com 跨进程 COM 状态污染

- 上一次 inspect script crash 留下 hung Excel.exe → 下一次 inspect 一开 dispatch 立刻 0x800ac472
- 缓解:`tasklist | grep excel` 检查 + 手工 kill stale process(沙箱可能 block)
- 表现:连续跑 frozen 回归时,第 2-3 张突然全失败,实际是 Excel 状态污染不是代码 bug

### 4.3 字段映射跨市场对不齐(中美 GAAP / IFRS / K-IFRS 差异)

- Phase 4h Step 2 跨市场 BS/IS/CF 合表:实测**字段交集 = 0**(完全没法对齐)
- P2 并集方案:117/49/107 行,A 股堆上面,港美韩堆下面 — **用户不可读**
- 决策:Phase 4j.1 干脆放弃 BS/IS/CF 合表,**只保留 18 项标准指标的合表**
- 教训:**别假设跨市场字段能 1:1 mapping**,标准化要在生成 18 项指标这一步搞定

### 4.4 frozen rule 语义不严密

- Phase 4j.1 把跨市场 BS/IS/CF + 字段映射 sheet 删了,但 inspect_phase4h_state.py 还在检查这些
- Codex stop-before-commit 等 Planner 决策:plan frozen 跟需求冲突
- 教训:**state-bound inspect** 必须显式允许"跟随主线 state 变更同步"
- 修正:后续 plan 把 inspect 同步规则写到 §⚠️

### 4.5 用户视角 vs 开发者视角 doc

- 早期使用说明 / 注释充满"WinHttp / cache / fallback / FX rate / shell-out / OERN"等技术黑话
- 用户(财务审计)看不懂
- Phase 4j 用户视角重写:全部翻译成业务话(eg "WinHttp" → "本地脚本拉取",fallback → "备用数据源")
- 教训:**写 doc 前先想读者**,不是写完才考虑

---

## 5. 技术债

| 项 | 说明 | 严重性 |
|---|---|---|
| 模块_工具函数.bas 3000+ 行 god module | HTTP/cache/FX/QA/write/UI 全混 | 中(P2-01,defer)|
| OERN 84 处全收敛 | Phase 4l 已收敛部分,剩残留 | 低(P1-07,defer)|
| MarketAdapter 抽象 | 4 市场各 ~500 行 duplicate logic | 低(无新市场需求)|
| AAPL fiscal year fuzzy match ±31 天 | Phase 4g 80% 兜住,fiscal metadata struct 全实现没做 | 低(P1-04,defer)|
| Cookie 仍存样本池 cell | 用户便利性优先,CleanReleaseWorkbook 宏分享前清 | 低(用户接受)|
| stockanalysis fallback 仅 BABA/JD/PDD 白名单 | 其他中概股不承诺 | 低(用户当前足够)|
| PowerShell 汇率 fetch shell-out | 个人 Windows 安全策略下没问题 | 低(P0-04,defer)|

---

## 6. 下一阶段应避免什么

### 6.1 Plan 阶段

- ❌ 不要把 inspect 列入 frozen 而不区分"真 frozen vs state-bound"
- ❌ 不要在一个 Phase 塞 5+ 件并行(Phase 4h 14h 工作量 review 负担过重,虽然成功但风险大)
- ❌ 不要在 plan 里说"Codex 看着办" — 要给具体 cell range / 颜色 hex / 行高 spec

### 6.2 Codex 协作

- ❌ 不要让 Codex 跑 macro 时 sandbox 同时启动其他 Excel COM(stale state 污染)
- ❌ 不要 review 时只看 Codex 自报 — 必须独立跑回归
- ❌ 不要在 plan 没明确"端到端 vs 多 round"时让 Codex 选(Codex 倾向端到端,但有些 phase 需要 mid-stream review)

### 6.3 视觉/UX 阶段

- ❌ 不要让 Codex 自由发挥视觉 — 必须给目标截图 + 像素级 spec(列宽、行高、颜色 hex)
- ❌ 不要在视觉 phase 改业务逻辑(eg Phase 4j 的样本池视觉重构跟跨市场对比简化分开做)

### 6.4 数据准确性

- ❌ 不要让汇率 / 单位 / 货币 缺失时 fallback 到 1(必须显式 missing 状态 + 诊断行)
- ❌ 不要只测 happy path — 加离线 fixture 测异常路径(429 / malformed / missing fields)

---

## 7. 后续类似项目模板化经验

### 7.1 Phase 节奏模板

```
Day 0:  目标 + 决策(AskUserQuestion 锁 4-5 个核心决策)
Day 1:  数据源接入(每个市场 / 每个数据源 1 phase,顺序串)
Day 2:  跨数据源整合 + 用户感知功能(toggle / 合表 / UX)
Day 3:  视觉抛光 + 用户视角 doc 重写
Day 4:  优化(GPT static audit 给 backlog → reviewer 选 ROI 高的做)+ release 收尾
```

### 7.2 frozen / state-bound 区分原则

| 文件类型 | 分类 | 规则 |
|---|---|---|
| 业务核心 helper(`fetch_*` / `cache_*` / `Get*Rate` 签名)| 真 frozen | 跨 phase 不动,跨 reviewer 守 |
| 集成测试驱动(`test_fx_live.py` / `diff_*`)| 真 frozen | 测核心 invariant,改它=测试基线变了 |
| state-bound inspect(`inspect_phase4*_state.py`)| state-bound | 当被检查的 sheet/button/cell 被明确变更时,必须同步 |
| `STATUS.md` / `PHASE_*_PLAN.md`| 历史追溯 | append-only,不重写已闭环段 |
| `README.md` / `ARCHITECTURE.md`| 用户面向 | 跟随主线变化更新 |

### 7.3 Plan 模板(每 phase 必须包含)

```markdown
> 版本 / 状态 / 作者 / 背景

## 项目语境(给 Generator 的 anchor 段)
[强调本期 = 已有数据源延续 / 不引入新风险]

## 用户决策(AskUserQuestion 锁定)
| 决策点 | 选项 |

## Step 总览
| Step | 内容 | 估时 | 阻塞依赖 |

## Step 详细(每 step)
- 背景
- 实施 spec(具体到 cell / 颜色 / 列宽)
- 验证
- Generator 不要做(明确禁止项)

## ⚠️ 全 Phase 严禁动
| 文件/区域 | 原因 |

## ⚠️ 联系 Planner 触发条件
- 具体到"任一回归退化" / "性能 > X 秒" / "命中率 < Y%"

## State-bound inspect 同步规则
[明确 inspect 跟随主线变更的边界]
```

### 7.4 GPT static audit 应用

- 当 codebase 到 5000+ 行 / 10+ phases 后,做一次 GPT 5.5 Pro 静态审阅
- Reviewer 评估 backlog 14 项,选 ROI 最高的 9-10 项分 sprint 做
- 不全做 — 个人/小团队工具不需要 enterprise-grade(eg secrets 管理 / MarketAdapter 抽象 都 over-engineering)

### 7.5 测试 pyramid(VBA 项目)

```
        ▲  集成测试 (test_fx_live + diff_*) - 联网真打,慢但 high-fidelity
       ▲▲  state inspect (inspect_phase*_state) - 每 phase 1 张,验 sheet/button/cell
      ▲▲▲  离线 fixture 测试 (Test_Offline_*) - 解析 / 缓存 / FX / QA, 不发 HTTP
```

每张测试覆盖一类(集成测真 HTTP,state inspect 测 layout,fixture 测 parse 逻辑)。

---

## 8. 总结

| 维度 | 评估 |
|---|---|
| 范围扩张 | 单市场 → 4 市场 + 跨市场合表 + 优化 4 sprints |
| 时间 | 4 天 vibe coding |
| 代码增量 | ~10K VBA + ~2.5K Python + ~5K 文档 |
| 测试覆盖 | 7 张 frozen 回归 + 8 离线 fixture |
| 视觉质量 | 财务审计专业人士可用(brand color 卡片 + 商务文档)|
| 数据准确性 | FX missing 不 fallback 1 + 数据质量 QA + 离线测试 |
| 用户体验 | B6 实时 toggle 0.13s + cache 4x 提速 + 一键 X 股 自动展开 |
| 可维护性 | ARCHITECTURE.md + STATUS.md + PHASE_*.md plans 完整追溯 |

**项目正式发布 v1.0**(release: `release/上市公司财务数据查询v1.0_release.xlsm`,source: `release/上市公司财务数据查询v1.0_source.xlsm`)。

后续维护:
- 实际使用反馈触发 patch sprint
- 数据源页面变化时,跑 `tools/run_offline_tests.py` 第一时间发现
- 雪球 cookie 失效时,用户主动更新 B5,fallback 路径自动接管

— Eric Zhang(214978902@qq.com),2026-05-05

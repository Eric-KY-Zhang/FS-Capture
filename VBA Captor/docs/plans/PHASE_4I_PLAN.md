# Phase 4i: UX 抛光 — 样本池重组 + 使用说明商务化 + 汇率说明区

> **版本**: v2(2026-05-04,Phase 4i 交付收口)
> **状态**: ✅ Phase 4i 全期闭环
> **作者**: Claude(planner) + Codex(generator)
> **背景**: Phase 4h 把 5 件主线 backlog 全清完后,样本池累积了 12 个按钮 (Q/S 列散布)、B8 fallback toggle 挤在 A 股按钮旁 — 视觉混乱;使用说明 tab 还是 Phase 4b 的纯文本 dump;汇率 tab 数据下方留白没有说明取数逻辑。本期纯 UX 抛光,不动任何业务逻辑。

## 项目语境(给 Generator 的 anchor 段)

**全部本地工作,零网络改动**。本期只重排样本池按钮位置 + 重写两张文档 sheet 内容 + 在汇率 sheet 数据下方追加说明区。不改任何 fetch / write / cache 路径,不动任何 commit hash 是 e06aaa2(Phase 4g)/ 286c28b(Phase 4h)的 frozen 文件。

## 用户 3 项核心要求(2026-05-04 用户截图反馈)

| # | 痛点 | 期望 |
|---|---|---|
| 1 | 样本池 12 个按钮 + 配置 cell + B8 toggle 视觉混乱 | 按功能分区 + 加 section label + 简要使用提示 |
| 2 | 使用说明 tab 太水(纯单列文本) | 商务/正式风:封面 + TOC + section 大标题 + 表格化 |
| 3 | 汇率 tab 数据区下面留白 | 加『数据源与取数逻辑』说明区 |

## Step 总览

| Step | 内容 | 估时 | 阻塞依赖 |
|---|---|---|---|
| 1 | 样本池重组(按钮分区 + section 标签 + B8 移位 + 内联使用提示) | 2h | 无 |
| 2 | 使用说明 tab 商务化(封面 + TOC + 7 section 重组 + 排版升级) | 2h | 无 |
| 3 | 汇率 tab 加『数据源与取数逻辑』说明区 | 0.8h | 无 |
| 4 | README 同步小修(反映新样本池布局) | 0.3h | 1 |
| 5 | 回归 + 验证(4 张 frozen 驱动 + 1 张新 inspect) | 0.5h | 1-4 |

**Codex 工作流建议(1 round 端到端)**:
- 全部 5 step 一轮交付,commit `Phase 4i UX polish (sample pool layout + intro sheet + fx legend)`,完成后停下报告 commit hash。这 5 step 全部 UX/doc,scope 小到不需要分 round。

---

## Step 1 — 样本池重组

### 1A. 当前布局问题(用户截图直接观察)

```
Q 列                                 S 列
Q1:Q3   一键全抓 4 市场               S1:S3   合并 4 张跨市场表
                                    
Q5:Q7   合并跨市场指标表              S5:S7   合并跨市场资产负债表
                                    S8:S10  合并跨市场利润表
Q8:Q10  切换所有分市场 tabs 显隐      S11:S13 合并跨市场现金流量表
                                    
Q14     清空缓存(orphan,no group)
```

**问题**:
- 合表按钮 5 个分布 Q5/S1/S5/S8/S11 — 不在同一列
- 切换 tabs / 清缓存 跟一键全抓挤在同一 Q 列,层级混乱
- B8 fallback toggle 挤在 F8(美股一键按钮旁),视觉割裂

### 1B. 新布局(目标)

**右侧按钮区按操作类型分 3 列**:

```
P 列(section 标签,加粗灰底)        Q 列(主操作)                R-S 列(合表细分)
                                                                
"主操作"                            Q1:Q3   一键全抓 4 市场       
                                                                
"合表"                              Q5:Q7   合并 4 张跨市场表    R5:S7   合并 跨市场_指标表
                                                               R8:S10  合并 跨市场_资产负债表
                                                               R11:S13 合并 跨市场_利润表
                                                               R14:S16 合并 跨市场_现金流量表
                                                                
"显示"                              Q18:Q20 切换所有分市场 tabs   
                                                                
"工具"                              Q22:Q24 清空缓存             
```

**或者更紧凑(用户偏好)**:

```
Q 列(主)                            S 列(合表细分)
Q1:Q3   一键全抓 4 市场             S1:S3   合并 4 张跨市场表
                                    
Q5:Q7   切换所有分市场 tabs         S5:S7   合并 跨市场_指标表
Q9:Q11  清空缓存                    S8:S10  合并 跨市场_资产负债表
                                    S11:S13 合并 跨市场_利润表
                                    S14:S16 合并 跨市场_现金流量表
```

**Codex 实现时选**:推荐紧凑版(避免占太多列宽),按上方紧凑布局摆放。

### 1C. 子任务

**1C.1** 在 `tools/install_modules.py` `BUTTONS` list 调整:
- 把 `BtnBuildCrossInd` 从 `Q5:Q7` 移到 `S5:S7`
- 把 `BtnBuildCrossBS` 从 `S5:S7` 移到 `S8:S10`
- 把 `BtnBuildCrossIS` 从 `S8:S10` 移到 `S11:S13`
- 把 `BtnBuildCrossCF` 从 `S11:S13` 移到 `S14:S16`
- 把 `BtnHideAll` 从 `Q8:Q10` 移到 `Q5:Q7`
- 把 `BtnClearCache` 从 `Q14:Q15` 移到 `Q9:Q11`

**1C.2** 在 `layout_sample_pool` (`install_modules.py`) 给 4 个 group 的最上面一行加 section label cell(浅灰底 + 加粗黑字 + 居中):
- `Q0`(实际 `R1` 上方:用 `O1` 或 `P1` 都行,选不冲突的位置)
  - 选 `P1:P3` "操作" + `R1:R3` "合表"
  - 或者用 mini header 行(`Q4` = "切换 / 工具",`S4` = "单独合表")— **Codex 选择最不挤压的方案**

**1C.3** B8 toggle 当前位置(F8?)— 实际是 `B8`(根据 STATUS §W.3"B8 开关为了兼容现有样本池按钮布局,位于 A 股一键按钮右侧"和 inspect 显示 B8 fallback toggle='关')。

把 B8 fallback toggle 从当前位置(`B8` 一键 A 股按钮 row 下方左侧?screenshot 显示 F8 = 开)迁移到顶部配置区:
- B5 = cookie(已有)
- B6 = 显示币种(已有)
- B7 = "中概美股 stockanalysis fallback 开关"(label)
- B8 = "开 / 关"(下拉,默认 关) — 跟 cookie / currency 同列同区

**注意**:`B8` cell 是 Phase 4h 新增,移位时:
- VBA `ReadStockAnalysisFallbackEnabled()` 函数(grep 一下找位置)需要更新读取的 cell 地址
- `install_modules.py` 装表时初始化的 cell 也要改
- 数据验证下拉选项 (`关 / 开`) 同步迁移
- 旧数据迁移:用户从 Phase 4h 升级时,如果 B8 以外位置有数据先保留,否则默认空

**1C.4** 在样本池 `O1:Q5` 区域(右侧空白区,按钮组左边)加 1 个内联"使用提示"框,内容简洁:

```
【使用提示】
1. A2 / A4 选年度 + 季度
2. B5 填雪球 cookie(港股 / 部分美股 fallback 必填)
3. B6 切原币 / 统一RMB(切换立即重算,无需重抓)
4. B8 默认关;雪球失效时改"开"启用 stockanalysis 中概美股 fallback
5. 第 11 行起按市场录入公司,Q/S 列点对应按钮跑数 + 合表
6. 跑数后 4 张『跨市场_*』sheet 自动展示对标视图
```

cell 格式:浅黄底 + 自动换行 + 字号 9 + 字体 微软雅黑

### 1D. 验证

- 装完打开 xlsm,样本池右侧 Q 列只有 3 个按钮(一键全抓 / 切换 tabs / 清空缓存),S 列只有 5 个合表按钮(合并 4 张 + indicator + BS + IS + CF)
- B8 fallback toggle 在 cookie / currency 下方,跟配置区同栏
- O1:Q5 内联使用提示框可见,文字不溢出

### 1E. Generator 不要做

- ❌ 不要改任何按钮的 OnAction(只动 cell 位置)
- ❌ 不要删除现有按钮(只改位置 + 新增 section label cell)
- ❌ 不要动样本池数据区(Row 11+ 用户数据)
- ❌ 不要动样本池左侧 4 市场分栏布局(Phase 4e 已 frozen,只动右侧操作区)

---

## Step 2 — 使用说明 tab 商务化

### 2A. 当前问题

`使用说明` sheet 53 行,全是单列 A 列纯文本(R1 标题 / R3-R49 散落内容),没有视觉层级。

### 2B. 新结构

**Sheet 重新规划成 7 个 section,加 TOC + 商务排版**:

```
┌─────────────────────────────────────────────────────┐
│  R1-R3   封面区:                                     │
│          - 大标题"上市公司财务数据查询"(字号 24 加粗) │
│          - 副标题"Multi-Market Financial Data Tool"  │
│          - 版本号 + 修订日期(右上对齐)               │
│          - 作者 / 联系邮箱                            │
├─────────────────────────────────────────────────────┤
│  R5-R12  目录(TOC):                                  │
│          - § 1 项目概览       → R14                  │
│          - § 2 快速开始       → R20                  │
│          - § 3 输出 sheet 说明 → R30                 │
│          - § 4 数据源声明      → R40                 │
│          - § 5 汇率换算说明    → R48                 │
│          - § 6 常见问题        → R56                 │
│          - § 7 版本历史        → R64                 │
├─────────────────────────────────────────────────────┤
│  R14+  正文 (各 section):                            │
│          每 section 顶部 section header:              │
│            - 编号 + 标题(字号 14 加粗 + 深蓝底白字) │
│            - 占 A:H 列宽合并                         │
│          内容根据 section 用 paragraph / table / list │
└─────────────────────────────────────────────────────┘
```

### 2C. 各 section 内容(Codex 实现时根据现有 README + 53 行 内容 + STATUS §W 总结)

**§ 1 项目概览** (R14-R18)
- 工具用途(1 段简介)
- 当前支持市场表(4 市场 + 数据源):3 列表格(市场 / 状态 / 数据源)

**§ 2 快速开始** (R20-R28)
- 6 步使用流程(编号列表),每步左侧 step 编号 cell + 右侧描述 cell

**§ 3 输出 sheet 说明** (R30-R38)
- Tab 顺序总览(2 列表格:tab 名 / 用途)
- 包括 4 张跨市场表 + 4 市场 × 4 张分市场表 + 3 张诊断 sheet + 汇率 + 样本池

**§ 4 数据源声明** (R40-R46)
- A股 / 美股 / 港股 / 韩股 / 汇率 各自数据源 + URL 模式(展示用,不超链)
- 隐私 / cookie 提示

**§ 5 汇率换算说明** (R48-R54)
- B6 toggle 行为(原币 / 统一RMB 两种模式)
- 期末汇率 vs 期间均值 应用规则(BS 用期末 / IS+CF 用均值)
- RMB 短路逻辑(reporting currency = RMB / CNY 直接 =1.0)
- 汇率手改不会反向刷,需要重跑

**§ 6 常见问题** (R56-R62)
- Q1: cookie 失效怎么办?
- Q2: 切换 B6 后数据变了一半?
- Q3: stockanalysis fallback 怎么启用?
- Q4: 老版样本池升级数据丢了?
- Q5: 跨市场合表行项目跨市场对不齐?

每 Q 占 1-2 行(Q 加粗 + A 普通)

**§ 7 版本历史** (R64-R72)
- Phase 4b → 4i 简表(版本 / 日期 / 主线交付),用 3 列表格

### 2D. 排版细节

- **配色**:封面深蓝(`#1F3864`)+ section header 中蓝(`#4472C4`)+ 表格 header 浅蓝(`#D9E1F2`)
- **字体**:微软雅黑 全局,封面 24/14/10 三级,section header 14,正文 11,table 10
- **行高**:封面区 R1=36/R2=22/R3=18;section header 22;正文 18;表格 16
- **列宽**:A=4(序号)/B=18(label)/C=42(描述)/D=15(状态)/E-H 按需
- **边框**:section 内表格加细边框,正文 paragraph 不加
- **冻结**:R4(让 TOC 始终可见)— 可选,Codex 看效果决定

### 2E. 实现路径

`tools/install_modules.py` 现有 `update_intro_sheet(wb)` 函数(grep 确认)— 重写其内容,从单列文本变成 multi-section 排版。**保留**函数签名 + 调用点不变。

### 2F. Generator 不要做

- ❌ 不要新增 sheet(只重写现有 `使用说明`)
- ❌ 不要 hardcode 个人联系方式之外的内容(版本号 / 修订日期可硬编码)
- ❌ 不要在使用说明 sheet 里塞图片 / 超链外部 URL(本地工具,无需引流)

---

## Step 3 — 汇率 tab 加『数据源与取数逻辑』说明区

### 3A. 当前问题

汇率 sheet R1 表头 + R2-R3 数据(USD 2024 / 2023 + HKD 2024 / 2023) + R4-R∞ 全空白。没有任何说明,新用户看不懂"USDCNY期末"和"USDCNY期均"什么时候用哪个。

### 3B. 新增说明区(J10+)

**J10**: section header `数据源与取数逻辑`(深蓝底白字 + 字号 14 + 合并 J:Q)

> v2 实现注:为避免污染汇率缓存 A 列 `End(xlUp)` 追加逻辑,实际说明区落在 J10+;A:H 继续留给 R1-R3 及未来新增报告期数据。

**R11-R20**: 内容,2 列布局(label / 描述):

```
A11: 数据源
B11: 雪球 K 线接口(单线程顺序请求 + 1s 间隔)
A12:   USDCNY
B12:   https://stock.xueqiu.com/v5/stock/chart/kline.json?symbol=USDCNYC ...
A13:   HKDCNY
B13:   https://stock.xueqiu.com/v5/stock/chart/kline.json?symbol=HKDCNYC ...
A14:   KRWCNY
B14:   https://stock.xueqiu.com/v5/stock/chart/kline.json?symbol=KRWCNYC ...

A16: 字段定义
B16:
A17:   USDCNY期末
B17:   报告期最后一天的 close 价(BS 用,反映时点资产负债)
A18:   USDCNY期均
B18:   报告期内每日 close 算术平均(IS / CF 用,反映期间损益 / 现金流)

A20: 应用规则
B20:
A21:   B6 = "原币"
B21:   非 RMB 报告币种公司直接显示原币数值,不调用汇率
A22:   B6 = "统一RMB"
B22:   按报告币种 + 报表类型查表,期末汇率乘以 BS,期间均值乘以 IS / CF
A23:   reporting currency = RMB / CNY
B23:   短路 = 1.0,跳过查表 + 不写汇率 sheet
A24:   汇率 sheet override 列 (H 列)
B24:   留作未来手填备注;手改本表 B-G 列数值不会反向刷写表(需要重跑写表按钮)

A26: 缓存策略
B26:
A27:   .cache/ 24h TTL
B27:   首次抓数 落地 ~30 个 JSON;24h 内重复跑同公司同期免 HTTP
A28:   清空缓存
B28:   样本池 Q9:Q11 按钮;清完下次跑数会重新拉所有汇率 + 财报数据

A30: 注意事项
B30:
A31:   汇率手改不会反向刷
B31:   写表时 FX 值 baked into 公式,要重新换算需点对应『一键 X 股』或『合并 4 张跨市场表』
A32:   B5 cookie 失效
B32:   汇率本身不需要 cookie(雪球 K 线公开),但 B5 雪球 cookie 同时被港股 + 美股 fallback 使用,失效时这两路径会失败
```

### 3C. 排版

- A 列宽 18(label),B 列宽 60(描述),自动换行
- A 列字体 微软雅黑 10 加粗,B 列字体 11 普通
- section header(J10)合并 J:Q,字号 14 加粗,深蓝底白字
- 子节标题(数据源 / 字段定义 / 应用规则 / 缓存策略 / 注意事项)字号 11 加粗,左对齐

### 3D. 实现路径

`tools/install_modules.py` 现有 `_make_fx_sheet(wb, name)` 函数 — 在其末尾追加调用 `_write_fx_legend(ws_fx)` 新 helper。

### 3E. Generator 不要做

- ❌ 不要动 R1-R3 现有数据 + 表头
- ❌ 不要新增数据列(8 列 frozen)
- ❌ 不要把 URL 写成超链(纯文本展示)

---

## Step 4 — README 同步小修

### 4A. 子任务

**4A.1** README.md "## 使用方式" 段更新按钮位置:
- "顶部按钮" 子段:`合并 4 张跨市场表` 描述句不变,但补一句"合表按钮统一在 S 列分组(BS / IS / CF / Indicator + 合并 4 张)"
- B8 改写:从"位于 A 股一键按钮右侧"改成"在 cookie / currency 配置区下方(B7 label + B8 toggle)"

**4A.2** "## 注意事项" 段保持不变(STATUS §W.3 已覆盖)

### 4B. Generator 不要做

- ❌ 不要重写 README 整体结构(只改 2 段说明)
- ❌ 不要移除任何现有 caveat

---

## Step 5 — 回归 + 验证

### 5A. 跑现有 4 张 frozen 驱动

```bash
cd "VBA Captor"
py tools/test_fx_live.py --skip-install      # Phase 4f Step 2 frozen
py -u tools/diff_phase4f_step3_lite.py       # Phase 4f Step 3-5 frozen
py -u tools/inspect_phase4g_state.py         # Phase 4g frozen
py -u tools/inspect_phase4h_state.py         # Phase 4h frozen
```

期望:**全 4 PASS**。Phase 4i 是纯 UX,任一退化说明动了不该动的代码。

### 5B. Phase 4i 专项 manual 检查(Codex 自测,无需新写 inspect 脚本)

打开 xlsm:
- **样本池**: 右侧 Q 列只有 3 个按钮 + S 列只有 5 个按钮,B8 在配置区,O1:Q5 内联使用提示可见
- **使用说明**: 看到封面 + TOC + 7 section header,排版商务风
- **汇率**: J10 起有"数据源与取数逻辑"说明区,A10 保持空白以不影响汇率缓存追加

### 5C. PHASE_4I_PLAN.md v1 → v2,STATUS §X 收口

**5C.1** Plan 状态行 → ✅ Phase 4i 全期闭环

**5C.2** STATUS.md 追加 §X(模仿 §W 格式):

```markdown
## X. Phase 4i UX 抛光: 样本池重组 + 使用说明商务化 + 汇率说明区

执行依据: `PHASE_4I_PLAN.md` v1。状态: ✅ Codex 已实现并通过 4 张 frozen 回归 + 手工 UX 检查。

### X.1 本阶段已完成
- [Step 1] 样本池: Q 列主操作 / S 列合表细分 / B8 toggle 移到配置区 / O1:Q5 内联使用提示
- [Step 2] 使用说明: 封面 + TOC + 7 section + 商务排版
- [Step 3] 汇率 sheet 增"数据源与取数逻辑"说明区(J10+,保留 A:H 数据区)
- [Step 4] README 按钮位置同步小修

### X.2 验证结果
- 4 张 frozen 驱动 全 PASS
- 手工检查 3 张 sheet 视觉效果符合 plan §1B/§2B/§3B

### X.3 已知边界
- Phase 4i 纯 UX, 不动业务逻辑;后续如需要继续优化样本池可作为 Phase 4i.1
- 使用说明 sheet 内容 hardcode 在 install_modules.py update_intro_sheet, 修改需要重装
```

---

## ⚠️ 全 Phase 严禁动的东西

| 文件/区域 | 原因 |
|---|---|
| `modules/*.bas` 全部 | Phase 4i 是纯 UX, **不动任何 VBA 业务逻辑**;唯一可能改的 VBA 是 `ReadStockAnalysisFallbackEnabled()` 读取 cell 地址(B8 移位需要),其他全 frozen |
| 现有 4 张回归驱动 + 1 张 inspect_phase4h_state | frozen |
| 4 市场 fetch 模块 + WriteWideTable + cache 层 | 全部业务逻辑 frozen |
| 跨市场表 4 张 sheet 内容 | Phase 4h 已交付,frozen |
| 样本池 Row 11+ 用户数据 | 数据安全,严禁清空 / 重写 |
| 汇率 sheet R1-R3 原有数据 | frozen |

## ⚠️ 联系 Planner 触发条件

- Step 1 B8 toggle 移位后任何 fetch 模块的 `ReadStockAnalysisFallbackEnabled()` 等 helper 在 grep 时找不到对应 cell 引用(说明 cell 名 hardcode 太分散,需要重新决策迁移策略)
- Step 2 `update_intro_sheet` 重写后任一 frozen 回归驱动退化(应该不会,但触发就停)
- Step 3 汇率 sheet `_make_fx_sheet` 重写后 `test_fx_live.py` 退化(说明动了不该动的数据区)
- Step 5 任一 frozen 驱动 PASS → FAIL

# Codex 接管: Phase 4f 后续(本地开发模式)

> **生效日期**: 2026-05-03
> **接管者**: Codex
> **协调者**: Claude Code(reviewer + planner,**不再执行代码**)
> **关键约束**: **Codex 从现在起只做本地代码开发,不做任何联网取数**

---

## 1. 项目当前状态 snapshot

### 1.1 已闭环(不再动)

| Phase | 内容 | 状态 |
|---|---|---|
| 1-3 | A 股(新浪)4 张表 + 一键全抓 + 圆角按钮 | ✅ |
| 4b-1 → 4b-14a | 美股(EDGAR + 雪球)4 张表 + 三级递进 + 诊断 sheet | ✅ |
| 4c | 港股(雪球 HK)4 张表 + 港股诊断 sheet | ✅ |
| 4d | 韩股(stockanalysis.com)4 张表 + 韩股诊断 sheet | ✅ |
| 4e | UX 优化(诊断 hidden + 样本池 4 列分市场 + 4 个市场专用一键) | ✅ |
| **4f Step 1** | **雪球 FX API + 傲基代码 probe** | ✅ **PASS_WITH_NOTES** |
| 4f Step 2 | 汇率 sheet + 模块_抓汇率.bas + ReadDisplayCurrency/GetFxRate helpers | 📋 task list 已详化, **Codex 接** |
| 4f Step 3-7 | 各市场 Run* hook 接入 + 实测 + 收口 | ⏳ |

### 1.2 文件 / 数据 ground truth(Codex 不需重新 probe)

**已 dump 到 `samples/` 的 Step 1 探查结果**:
- `xueqiu_quote_currency.json` — 雪球 quote 实时汇率 dump
- `xueqiu_kline_USDCNY.json` — 雪球 K 线历史汇率 dump(完整 365d 数据 + count 上限实测)
- `aoji_probe_summary.json` + 5 个 raw probe dump — 傲基股份 = 02519.HK 双重确认
- `RMB_FX_PROBE.md` — Step 1 探查报告(13.4 KB,6 节齐全)
- `KR_API_PROBE.md` / `HK_API_PROBE.md` — Phase 4c/4d 留存的 probe 报告

**已可重跑的 probe 脚本** `tools/probe_*.py`(用户 / Claude Code 后续可跑,**Codex 别跑**):
- `probe_1a_quote_fx.py`
- `probe_1b_kline_fx.py`
- `probe_1c_aoji.py`

### 1.3 Phase 4f Step 1 的 7 条 must-fix(已写入 plan v2 头部)

| # | 实测纠正 | 来源 |
|---|---|---|
| 1 | FX symbol 必带 `.FX`(`USDCNY.FX/HKDCNY.FX/KRWCNY.FX`) | quote symbol_variants 对照 |
| 2 | FX 端 `Accept-Encoding: gzip, deflate`(NOT identity) | plan_call vs working_call body_len 对照 |
| 3 | 进程启动需 GET `https://xueqiu.com/hq` 暖 session | warmup 拿 visitor cookie |
| 4 | quote 响应 `parsed["data"]` 直接 list;kline `parsed["data"]["item"]` 是 list 12 列 | dump 实测 |
| 5 | K 线 ts=item[0]=4-byte ms(VBA 1-based 索引 +1),close=item[5] | column 实测 |
| 6 | 时区 = CST 0 点 ms;day bar 落 04:00 UTC = 12:00 CST | first_item mod 86400000 = 14400000 |
| 7 | **不动** `模块_工具函数.bas` line 560/619/666 的 `Accept-Encoding: identity`(美股/港股/韩股已验证用 identity 工作);**FX 用独立 `XueqiuFxHttpGet`** | 风险隔离 |

---

## 2. Codex 接管范围 + 约束

### ✅ 可做(本地代码开发,完全离线)

- 写 VBA 模块(`modules/*.bas`)
- 改 Python 工具(`tools/*.py`)
- 改 plan / STATUS / handoff MD 文件
- 跑 `py tools/build_template.py`(生成空 xlsx,纯 openpyxl 离线)
- 跑 `py tools/install_modules.py`(本地 Excel COM 注入模块,**不联网**)
- 用 Excel 打开 .xlsm 做 VBA 编译验证(Alt+F11 → Debug → Compile,**离线**)
- 参考 `samples/` 已 dump 的样本文件做实现参考
- 参考 `PHASE_4F_RMB_PLAN.md` v2 / `PHASE_4F_STEP2_TASKS.md` 的 schema/伪码

### ❌ 不做(任何联网 / 实测)

- ❌ 不跑实际抓数(任何「一键全抓」/「一键 X 股」/「更新 X 表」按钮)
- ❌ 不跑 `EnsureFxRateCached("2024-12-31", "USD")` 立即窗口实测
- ❌ 不跑 `WarmupXueqiuSession()` 联网验证
- ❌ 不跑任何 `curl` / `requests` / `WinHttp` 联网请求
- ❌ 不跑 `tools/probe_*.py`(已有 dump 够用)
- ❌ 不跑 `tools/diff_xlsm.py` 跨工作簿对比(只读但触及 .xlsm 写入,留 Claude Code)

**实测联网验证由 Claude Code(我)/ 用户后续手动跑**,Codex 提交后等待我 review + 跑联网验证。

---

## 3. Phase 4f Step 2 立即可执行清单

完整 task 见 `PHASE_4F_STEP2_TASKS.md`。**Planner 推荐执行顺序**:

```
1. 2C.1-2C.2: 模块_工具函数.bas 加常量 FX_SHEET/FX_DATA_ROW + ReadDisplayCurrency()
2. 2A.1-2A.6: build_template.py + install_modules.py 加『汇率』sheet(8 列布局)
3. 2D.1-2D.6: 样本池 B6 toggle 「显示币种」(原币 / 统一RMB)
4. 2C.3+:   GetFxRate(...) — 先 stub return 0(后面 2C 回填)
5. 2B.1-2B.14: 模块_抓汇率.bas 全部(~250-300 行,核心模块)
6. 2C 回填: GetFxRate 真正连 EnsureFxRateCached
7. 2A.7-2A.8: 使用说明 sheet 加汇率段
```

### Codex 完成后必做的离线验证

| 验证 | 方法 | 预期 |
|---|---|---|
| 1. VBA 编译 | Excel 打开 .xlsm → Alt+F11 → Debug → Compile VBAProject | 0 error |
| 2. 模板生成 | `py tools/build_template.py` | xlsx 含 `汇率` sheet,Row 1 = 8 列表头 |
| 3. 安装 | `py tools/install_modules.py` | xlsm 含 `+ installed: 模块_抓汇率`,日志含 `+ sheet 新建: 汇率` |
| 4. Sheet 顺序 | 打开 xlsm 看 tab 列表末尾 | `汇率` 在最后(诊断 sheet 之后) |
| 5. B6 默认 | 打开样本池看 B6 cell | "原币",下拉两选 |
| 6. line 不变 | `grep "Accept-Encoding.*identity" modules/模块_工具函数.bas` | 仍 3 行(line 560/619/666 不动) |
| 7. FX 模块独立 | `grep "gzip, deflate" modules/模块_抓汇率.bas` | 出现 ≥ 2 次(warmup + FX HTTP) |
| 8. .FX symbol | `grep "USDCNY\\.FX" modules/模块_抓汇率.bas` | 至少 1 行 |

### Codex **不要**跑的(留给联网验证)

- ❌ 立即窗口跑 `?GetFxRate("USD", "2024-12-31", True)`
- ❌ 立即窗口跑 `?ReadDisplayCurrency()`
- ❌ 立即窗口跑 `?EnsureFxRateCached("2024-12-31", "USD")`
- ❌ 检查 `汇率` sheet 是否真的写入数据

这些**必须有真实联网 + 雪球 cookie**,Codex 跳过,Claude Code / 用户接手验证。

---

## 4. Phase 4f Step 3-7 路线图(Codex 完成 Step 2 后继续)

完整 plan 在 `PHASE_4F_RMB_PLAN.md`。摘要:

### Step 3 — 各市场 Run* hook 接入 GetFxRate

修改 4 个市场 Run* 主函数,在写 cell 前 hook:
```vba
Dim displayCurrency As String: displayCurrency = ReadDisplayCurrency()
Dim fxRate As Double: fxRate = 1#
If displayCurrency = "统一RMB" Then
    fxRate = GetFxRate(企业报告币种, 报告期, (报表类型 = "BalanceSheet"))
    If fxRate <= 0 Then fxRate = 1#  ' 失败兜底
End If
cellValue = cellValue * fxRate
```

诊断 sheet 加第 11 列 `FX_Rate`。**Codex 可做(纯代码),不联网**。

### Step 4 — 用户实际场景测试

样本池配置 5 A 股 + 02519 傲基 + Zinus 韩股,B6=统一RMB,跑一键全抓。
**联网验证,Codex 不做,Claude Code / 用户跑**。

### Step 5 — STATUS.md §T 收口

汇总 Step 2-4 完成情况,Phase 4f 闭环。**Codex 可做(纯文档)**。

### Step 6-7 (可选 #2 合表) — 推荐保留 16 张分市场表,本期不动

留 Phase 4g 决定。

---

## 5. Codex 工作流(本地开发版)

```
1. 读 PHASE_4F_RMB_PLAN.md(含 v2 修订纪要 7 条 must-fix)
2. 读 PHASE_4F_STEP2_TASKS.md(Generator 实施清单 + 验收 criteria)
3. 按 Planner 推荐顺序写代码:
   2C.1-2C.2 → 2A → 2D → 2C.3+ stub → 2B → 2C 回填 → 2A.7-2A.8
4. 每个子任务完成后:
   - VBA 编译验证(Alt+F11 → Compile)
   - `py tools/install_modules.py` 验证打包
   - 不跑联网验证
5. 全部 Step 2 完成后:
   - 写 STATUS.md §T 草案(Phase 4f Step 2 完成段)
   - 在 PHASE_4F_RMB_PLAN.md 头部加 "Step 2 ✅ Codex 实现完毕,等待联网验证 by Claude Code/用户"
6. 报告给协调者(Claude Code / 用户)
7. 等联网验证反馈,有 bug 再回炉
```

## 6. 文件清单(Codex Phase 4f Step 2 改动范围)

### 新增

- `modules/模块_抓汇率.bas` (~250-300 行)

### 改动(只追加,不删除现有内容)

- `modules/模块_工具函数.bas` (line 642 后追加 ~80 行,0 改动现有 line 552-676 区域)
- `tools/build_template.py` (+ ~50 行: build_fx_sheet + B6 toggle + main 调用)
- `tools/install_modules.py` (+ ~80 行: _make_fx_sheet + install_currency_toggle_cell + ensure_market_sheets + reorder + layout_sample_pool + update_intro_sheet)

### 不动

- `modules/模块_工具函数.bas` line 552-676 区域(`EdgarHttpGet` / `XueqiuHttpGet` / `EdgarHttpGetTickers` 三个 identity HTTP 客户端 — 美股/港股 4a-4e 已验证)
- 所有 `modules/模块_抓*.bas` 主流程(Step 3 才改)
- 所有现有 sheet(除『汇率』新增)
- `tools/probe_*.py`(Step 1 留存的探查脚本,**别跑别改**)
- `samples/` 全部内容(已是 ground truth)

---

## 7. 给 Codex 的开工 prompt

复制下面这段直接给 Codex:

> **Phase 4f Step 2 — Codex 接管,本地开发模式**
>
> 完整 plan: `PHASE_4F_RMB_PLAN.md`(v2 头部有 7 条 must-fix)
> 详细 task: `PHASE_4F_STEP2_TASKS.md`(Planner 已细化 4 子任务 2A/2B/2C/2D,含验收 criteria)
> 接管约束: `CODEX_HANDOFF_PHASE_4F.md`(本文)
>
> **硬约束**:
> 1. 只做本地代码开发,**不联网 / 不跑 probe / 不跑实际抓数 / 不跑立即窗口实测**
> 2. 实测联网验证由 Claude Code(我)/ 用户后续接手
> 3. 沿用 Step 1 已 dump 的 `samples/RMB_FX_PROBE.md` 等 ground truth,不要重新 probe
> 4. 7 条 must-fix 必须严格遵守(尤其 #7: 不动 line 560/619/666)
>
> **执行顺序**(Planner 推荐):
> 1. 2C.1-2C.2: 工具函数加常量 + ReadDisplayCurrency
> 2. 2A: build_template + install 加汇率 sheet
> 3. 2D: 样本池 B6 toggle "原币/统一RMB"
> 4. 2C.3+ stub: GetFxRate(返 0)
> 5. 2B: 模块_抓汇率.bas (~250-300 行)
> 6. 2C 回填: GetFxRate 接 EnsureFxRateCached
> 7. 2A.7-2A.8: 使用说明
>
> **完成后离线验证**(see `CODEX_HANDOFF_PHASE_4F.md` §3 验收清单):
> - VBA 编译 0 error
> - `py tools/install_modules.py` 成功
> - tab 列表末尾 `汇率`,Row 1 = 8 列表头
> - 样本池 B6 默认 "原币" + 下拉两选
> - grep `identity` 工具函数仍 3 行(line 不变)
> - grep `gzip, deflate` 抓汇率出现 ≥ 2 次
> - grep `USDCNY.FX` 抓汇率 ≥ 1 行
>
> **报告**:
> - 4 子任务每个的 status(✅ / ⚠️ / ❌)
> - VBA 编译输出
> - install 输出最后 20 行
> - 文件清单(改动 + 新增)
> - 任何偏离 task 计划的地方
> - **明确标注「未跑联网验证,等 Claude Code 接手」**
>
> 完成后 commit + 通知 Claude Code 做联网验证 + 最终 review。

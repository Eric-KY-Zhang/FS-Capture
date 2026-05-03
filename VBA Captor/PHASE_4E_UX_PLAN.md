# Phase 4e: UX 优化 — 诊断隐藏 + 样本池 4 列分市场

> **版本**: v1(2026-05-03)
> **状态**: ✅ 已由 Codex 实现并通过本地验证,等待 Claude Code review
> **作者**: Claude(planner)+ Codex(executor)
> **背景**: Phase 4d 韩股闭环,4 市场全部跑通。但工作簿 tab 数膨胀到 21 个,样本池单列混 4 市场 → UX 拥挤。本期先收快胜(#1+#3),其余 #2/#4/#5 留 Phase 4f+。

## 已锁定的决策(用户已确认)

| 决策 | 选项 |
|---|---|
| 优先级 | **快胜组合 #1 (诊断隐藏) + #3 (样本池 4 列分市场)** |
| 诊断 sheet 隐藏方式 | `xlSheetHidden`(用户右键能 unhide,默认不污染 tab 列表)— 不用 `xlSheetVeryHidden` |
| 样本池布局 | **4 列分市场 + 各市场顶部专用一键 + 全局「一键全抓 4 市场」保留**(详见下方布局图) |
| 旧样本池迁移 | **install_modules.py 自动迁移**(读旧 A-C 三列数据,按 `ResolveMarket` 推断分到 4 个市场新列;过程透明,有 print 日志) |
| 4 市场内部「一键 X」 | 各市场新增专用 macro `Sub 一键A股() / 一键美股() / 一键港股() / 一键韩股()`,调本市场 4 张表 + 触发本市场诊断刷新 |

## 执行收口(2026-05-03)

| 项目 | 状态 | 结果 |
|---|---|---|
| #1 诊断隐藏 | ✅ 完成 | VBA + install 双保险设 `xlSheetHidden`,3 张诊断表默认隐藏 |
| #3 样本池四市场分栏 | ✅ 完成 | A:B / E:F / I:J / M:N 分别对应 A股/美股/港股/韩股 |
| 旧样本池迁移 | ✅ 完成 | 17 家公司迁移结果 A:4 / US:4 / HK:4 / KR:5 |
| 4 个市场专用一键 | ✅ 完成 | `一键A股/美股/港股/韩股` 均已跑通 |
| 全局一键 | ✅ 完成 | `一键全抓 4 市场` 已跑通,四张指标表公式错误 0 |
| #5 RMB 策略 | 📝 备忘 | 已写入 `STATUS.md §S.3`,本期不实施 |

## 项目 #5 (统一 RMB) 汇率策略锁定(留 Phase 4g 实施)

用户已锁:**期末/期间平均混合**(会计准则推荐)
- BS(时点数)→ 期末汇率
- IS / CF(期间数)→ 期间平均汇率
- 数据源:雪球 quote API(`USDCNY` / `HKDCNY` / `KRWCNY`)+ 用户在新 sheet `汇率` 手动 override 兜底
- **本期不做**,只在 STATUS.md 留备忘

---

## 项目 #1 — 诊断 sheet 隐藏

### 实施位置(双保险)

**A. VBA 端**(`模块_工具函数.bas` `EnsureDiagnosticSheet`):
```vba
' 在创建/获取 ws 后, 末尾追加
On Error Resume Next
ws.Visible = xlSheetHidden    ' 0
Err.Clear
On Error GoTo 0
```
注意: `Visible = xlSheetHidden` 时如果 ws 是当前活动 sheet 会报错,所以要 `On Error Resume Next` 保护。

**B. install_modules.py 端**(`ensure_market_sheets` 创建诊断 sheet 后):
```python
diag_sheet.Visible = 0    # xlSheetHidden
```

**C. reorder_report_sheets 兼容性**:
- 隐藏的 sheet 仍能被 `Move` 操作影响顺序(测试一下)
- 如果 reorder 失败,临时 `.Visible = -1` 操作完再设回 0

### 验收
- 安装后打开 V3.xlsm,**tab 栏只见 18 个 sheet**(使用说明 + 样本池 + 16 张正式表),诊断 3 张不可见
- 右键任意 sheet tab → "取消隐藏" → 列表里能选中 3 张诊断 sheet
- VBA 跑数后诊断仍正常写入(隐藏不影响写入,只影响显示)

---

## 项目 #3 — 样本池 4 列分市场

### 新布局图(target)

```
       A      B       C   D       E      F        G    H      I       J     K    L      M       N      O    P
Row 1:  [年份]
Row 2:  [年份值=2025]
Row 3:  [季度]
Row 4:  [季度值=全部]
Row 5:  [雪球Cookie]   [....cookie value (B5:F5 合并).............................................]
Row 6:
Row 7:  ┌────────── A 股 (新浪) ─────────┐  ┌─────── 美股 (EDGAR + 雪球) ────────┐  ┌── 港股 (雪球 HK) ─┐  ┌── 韩股 (stockanalysis) ─┐  ┌─ 全局 ─┐
Row 8:  │  [一键 A 股]                   │  │  [一键 美股]                       │  │  [一键 港股]      │  │  [一键 韩股]            │  │ [一键全抓 4 市场] │
Row 9:  │ 代码  │ 简称 │                 │  │ 代码 │ 简称 │                     │  │ 代码 │ 简称 │      │  │ 代码 │ 简称 │           │
Row 10+:│ 300866│ 安克 │                 │  │ AAPL │ Apple │                    │  │ 00700│ 腾讯 │      │  │ 005930│ 三星 │         │
        └────────────────────────────────┘  └────────────────────────────────────┘  └──────────────────┘  └─────────────────────────┘

A 股 列: A=代码 B=简称
美股 列: E=代码 F=简称
港股 列: I=代码 J=简称
韩股 列: M=代码 N=简称
全局   : Q1 「一键全抓 4 市场」
间隔   : C/D / G/H / K/L / O/P 各 spacer
```

### 列宽

| 列 | 宽度 | 用途 |
|---|---:|---|
| A | 11 | A 股代码 |
| B | 16 | A 股简称 |
| C | 2 | spacer |
| D | 2 | spacer |
| E | 8 | 美股代码 |
| F | 18 | 美股简称 |
| G | 2 | spacer |
| H | 2 | spacer |
| I | 7 | 港股代码 |
| J | 14 | 港股简称 |
| K | 2 | spacer |
| L | 2 | spacer |
| M | 8 | 韩股代码 |
| N | 16 | 韩股简称 |
| O | 2 | spacer |
| P | 2 | spacer |
| Q | 22 | 全局按钮区 |

### Row 7-8 设计

**Row 7**: 市场标题(合并单元格 + 市场色填充)
- A7:B7 合并 = "A 股(新浪)" + 蓝底 `#4472C4` 白字
- E7:F7 合并 = "美股(EDGAR+雪球)" + 红底 `#C00000` 白字
- I7:J7 合并 = "港股(雪球 HK)" + 绿底 `#548235` 白字
- M7:N7 合并 = "韩股(stockanalysis)" + 紫底 `#7030A0` 白字

**Row 8**: 各市场专用「一键 X」按钮
- 按钮覆盖 A8:B8 / E8:F8 / I8:J8 / M8:N8(各占 2 cell 宽)
- 颜色与标题对应

**Row 9**: 数据列头
- A9 `代码` / B9 `简称` / E9 `代码` / F9 `简称` / 等等
- 不需要 C 列「市场」(每列固定市场,redundant)

**Row 10+**: 公司数据(各市场独立录入)
- A 股: A10-A1000 / B10-B1000
- 美股: E10-E1000 / F10-F1000
- 港股: I10-I1000 / J10-J1000
- 韩股: M10-M1000 / N10-N1000

**Q1**: 全局「一键全抓 4 市场」按钮(深蓝主按钮)

### 配置区合并

- A1:A2 → 年份(label + value 各占一行)
- A3:A4 → 季度
- A5 → cookie label,B5:F5 合并 → cookie 长输入框

或者更紧凑:Row 1-5 配置区跨 4 市场列上方(因为是全局配置)

### VBA 改动

**新增常量**(模块_工具函数.bas):
```vba
Public Const POOL_A_CODE_COL As Long = 1     ' A
Public Const POOL_A_NAME_COL As Long = 2     ' B
Public Const POOL_US_CODE_COL As Long = 5    ' E
Public Const POOL_US_NAME_COL As Long = 6    ' F
Public Const POOL_HK_CODE_COL As Long = 9    ' I
Public Const POOL_HK_NAME_COL As Long = 10   ' J
Public Const POOL_KR_CODE_COL As Long = 13   ' M
Public Const POOL_KR_NAME_COL As Long = 14   ' N
Public Const POOL_DATA_START_ROW As Long = 10   ' 从 Row 8 改成 Row 10
```

**4 个 Run*Statement 读样本池逻辑**:
- 旧: `arrPool = wsPool.Range("A" & POOL_DATA_START_ROW & ":H" & lngRow).Value` + `ResolveMarket` 过滤
- 新: 各 Run 自己读固定列对(A/B 或 E/F 或 I/J 或 M/N)
- 例如 `RunUSStatement` 读:
  ```vba
  Dim lastRow As Long
  lastRow = wsPool.Cells(wsPool.Rows.Count, POOL_US_CODE_COL).End(xlUp).Row
  arrPool = wsPool.Range(wsPool.Cells(POOL_DATA_START_ROW, POOL_US_CODE_COL), _
                          wsPool.Cells(lastRow, POOL_US_NAME_COL)).Value
  ```
- 不再需要 `ResolveMarket` 调用

**4 个市场专用入口 macro**(模块_总入口.bas):
```vba
Public Sub 一键A股()
    g_silentMode = True
    模块_抓资产负债表.Main
    模块_抓利润表.Main
    模块_抓现金流量表.Main
    模块_抓指标表.Main
    g_silentMode = False
    ' 单独一键不弹全局汇总;各 Sub 自己的 silent flag 控制
End Sub

Public Sub 一键美股()
    g_silentMode = True
    g_diagnosticSheetName = "美股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticAppendOnly = True
    模块_抓美股资产负债表.Main
    模块_抓美股利润表.Main
    模块_抓美股现金流量表.Main
    模块_抓美股指标表.Main
    g_diagnosticAppendOnly = False
    g_silentMode = False
End Sub

Public Sub 一键港股()
    g_silentMode = True
    g_diagnosticSheetName = "港股_抓取诊断"
    ClearDiagnosticSheet
    g_diagnosticAppendOnly = True
    模块_抓港股资产负债表.Main
    ' ...
End Sub

Public Sub 一键韩股()
    ' 类似
End Sub

' 全局「一键全抓 4 市场」原 一键全抓() Sub 保持不变
```

### install_modules.py 改动

**新 BUTTONS 列表**:
```python
BUTTONS = [
    # 全局
    ("BtnRunAll",     "一键全抓 4 市场",  "模块_总入口.一键全抓",   42, PRIMARY_FILL,  PRIMARY_FG,  13, True),
    # 4 市场专用一键
    ("BtnRunA",       "一键 A 股",        "模块_总入口.一键A股",    36, "4472C4",      "FFFFFF",    12, False),
    ("BtnRunUS",      "一键 美股",        "模块_总入口.一键美股",   36, "C00000",      "FFFFFF",    12, False),
    ("BtnRunHK",      "一键 港股",        "模块_总入口.一键港股",   36, "548235",      "FFFFFF",    12, False),
    ("BtnRunKR",      "一键 韩股",        "模块_总入口.一键韩股",   36, "7030A0",      "FFFFFF",    12, False),
    # 16 个单表按钮可保留也可移到独立 sheet,用户角度按钮过多反而干扰
    # 推荐:保留但折叠到 row 30+(滚动可见,不挤主区)
    # ... 16 个单表按钮 ...
]
```

**按钮位置**:
- 各市场顶部按钮(`BtnRunA/US/HK/KR`)放 Row 8 各市场列对(A8:B8 / E8:F8 / ...)
- 全局「一键全抓 4 市场」放 Q 列顶部(Q1 起,42pt 高)
- 16 个单表按钮:折叠到 Row 30+ 区域(辅助按钮区,不主推)

**install_quarter_cell + install_xueqiu_cookie_cell**:
- A1-A4 配置区保持(全局)
- A5 cookie label / B5:F5 合并 cookie 值(跨 4 市场列上方)

**自动迁移逻辑**(`migrate_old_sample_pool`):
```python
def migrate_old_sample_pool(ws_pool):
    """
    旧布局: A=代码 / B=简称 / C=市场 (Row 8+)
    新布局: A:B=A股 / E:F=美股 / I:J=港股 / M:N=韩股 (Row 10+)

    过程:
    1. 读旧 A8:C{lastRow}
    2. 按 C 列市场分到 4 个新列对
    3. 清空旧 A8:H1000 区
    4. 重新布局表头 + Row 7-9 配置
    5. print 迁移摘要
    """
    # 读旧
    last_row = ws_pool.Range("A" + str(ws_pool.Rows.Count)).End(xlUp).Row
    if last_row < 8:
        return  # 空样本池,无需迁移
    old_data = ws_pool.Range(f"A8:C{last_row}").Value

    # 分类
    by_market = {"A": [], "US": [], "HK": [], "KR": []}
    for row in old_data:
        code, name, market = row[0], row[1], row[2] or ""
        if not code: continue
        market = market.upper().strip()
        if market not in by_market:
            # 自动判断
            if str(code).isalpha(): market = "US"
            elif len(str(code)) == 5: market = "HK"
            elif len(str(code)) == 6 and not market: market = "A"  # 6 位无注 → A 股
            else: continue
        by_market[market].append((code, name))

    # 清空旧
    ws_pool.Range("A8:H1000").Clear()

    # 写新 (Row 10 起)
    col_map = {"A": (1, 2), "US": (5, 6), "HK": (9, 10), "KR": (13, 14)}
    for market, companies in by_market.items():
        code_col, name_col = col_map[market]
        for i, (code, name) in enumerate(companies):
            ws_pool.Cells(10 + i, code_col).Value = code
            ws_pool.Cells(10 + i, name_col).Value = name

    # print 迁移摘要
    summary = " / ".join(f"{m}: {len(v)}" for m, v in by_market.items())
    print(f"  + 样本池迁移: {summary}")
```

**新版 build_sample_pool 模板**(`tools/build_template.py`):
- 完全重写
- Row 1-5 全局配置(年份/季度/cookie)
- Row 7-8 4 市场标题 + 按钮位
- Row 9 数据列头
- Row 10 起公司数据(初始可放示例 4 家:300866 安克 / AAPL / 00700 / 005930)

### 验收

#### Test 1 — 新模板生成 + install
- `py tools/build_template.py` → 新 .xlsx 含 4 市场列
- `py tools/install_modules.py` → V3.xlsm 装好按钮(4 市场顶部 + 全局)
- 视觉确认:Row 7 4 个市场标题色对应,Row 8 4 个按钮色对应

#### Test 2 — 旧版迁移
- 临时把样本池改回旧布局(A:C 三列混市场,300866/AAPL/00700/005930 都在 A 列)
- 跑 install_modules.py
- **期望**: 自动迁移到新布局,4 市场各列正确分布

#### Test 3 — 跑数验证
- 配置 A2=2024 A4=Q4
- 点 4 个市场专用「一键」分别跑
- 点全局「一键全抓 4 市场」
- **期望**: 跟 Phase 4d 验收等价(16 表 + 3 诊断隐藏 sheet 数据齐全)

#### Test 4 — 诊断 hidden 验证
- tab 栏只见 18 个 sheet
- 右键 sheet tab → "取消隐藏" → 能选 3 张诊断 sheet → 显示数据正常
- 重新隐藏(右键 → "隐藏")也正常

## 文件改动清单

| 文件 | 改动 | 责任 |
|---|---|---|
| `modules/模块_工具函数.bas` | 新增 4 市场列常量 + `EnsureDiagnosticSheet` 末尾设 hidden | Codex |
| `modules/模块_抓资产负债表.bas` 等 4 个 A 股 Main | 改读 A:B 列 | Codex(改 RunOneStatement 接收 column args 或通过常量) |
| `modules/模块_抓美股财报.bas` `RunUSStatement` | 改读 E:F 列 | Codex |
| `modules/模块_抓港股财报.bas` `RunHKStatement` | 改读 I:J 列 | Codex |
| `modules/模块_抓韩股财报.bas` `RunKRStatement` | 改读 M:N 列 | Codex |
| `modules/模块_总入口.bas` | 新增 `一键A股 / 一键美股 / 一键港股 / 一键韩股` 4 个 Sub | Codex |
| `tools/build_template.py` `build_sample_pool` | 完全重写新版布局 | Codex |
| `tools/install_modules.py` | `BUTTONS` 重构 + `migrate_old_sample_pool` 新增 + 配置区/按钮位置全改 | Codex |

## 实施顺序建议(Codex)

```
Day 1:
  Side: Phase #1 诊断 sheet hidden (1 小时, 双保险 VBA + Python)
  Step 1: 4 市场列常量 + 4 个 Run*Statement 读列改造 (3-4 小时)
  Step 2: 4 个市场专用一键 macro (1 小时)

Day 2:
  Step 3: build_template.py 新版 build_sample_pool (3-4 小时)
  Step 4: install_modules.py 大重构 (4-6 小时, 含 migrate 逻辑)

Day 3:
  Test 1/2/3/4 + 修 bug
  STATUS.md §S 收口
```

## 风险点

1. **旧样本池迁移稳定性** — 用户可能手填了不规范市场列(中文"美国"/"港股"等),`migrate` 要兜底自动判断
2. **按钮位置碰撞** — Row 8 4 个市场按钮宽度需精确算,避免重叠;每个按钮宽 ≈ 2 列宽 = `Range("A8:B8").Width`
3. **`ResolveMarket` 是否还需要** — 4 列固定市场后,`ResolveMarket` 在新流程不调用,但保留作为 helper 不删(防止其它地方依赖)
4. **现有用户工作流** — 用户已有的 4 公司 A 股(300866 等)+ 4 公司美股 + 4 港股 + 5 韩股 共 17 条,迁移后必须自动正确分布
5. **诊断 hidden 跟 reorder_report_sheets 兼容** — 隐藏的 sheet 能否被 `Move` 操作顺序? Codex 实测确认

---

## 给 Codex 的执行指南

1. **诊断 hidden 立即做**(双保险 VBA + Python),1 小时
2. **样本池重构** 是大改,先在脑子里完整过 layout,再动手
3. **`migrate_old_sample_pool`** 必做,否则用户从老 V3.xlsm 升级会丢数据
4. **测试公司样本** 用之前 4 家 A 股 + 4 家美股 + 4 家港股 + 5 家韩股 共 17 条,跑通 4 个市场专用一键 + 全局一键
5. **STATUS.md §S 收口** 完成后写

---

## Phase 4f+ 待办(future, 不在本期)

- **#4 stockanalysis 替代雪球**(probe 美股/港股 → 决定是否切换,2-3 天)
- **#5 统一 RMB 跨市场**(汇率 API + 期末/期间平均策略 + UX toggle,3-5 天)
- **#2 合表 4 市场同表**(必须等 #5 完成才有意义,2-3 天)

汇率策略已锁(用户确认):**期末/期间平均混合(会计准则推荐)**
- BS 用期末汇率 / IS+CF 用期间平均汇率
- 数据源候选:雪球 quote API + 用户手动 override 兜底
- 详细方案 Phase 4g 启动时再细化

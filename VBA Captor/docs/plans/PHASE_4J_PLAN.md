# Phase 4j: 跨市场表精简 + 用户视角 doc + 视觉规范 + tab 行为修正

> **版本**: v2(2026-05-04,scope 缩减后重写)
> **状态**: 🚧 READY FOR GENERATOR
> **作者**: Claude(planner) + Codex(generator)
> **背景**: 用户实际使用 Phase 4i.2 后反馈:跨市场 BS/IS/CF 3 张表(Phase 4h 交付)实际上意义不大(行项目跨市场对不齐 + 信息密度太低),只留 18 标准指标合表的『跨市场_指标表』就够用。同时 Phase 4i.2 的样本池布局 + 一键按钮还有 4 项小改进。

## 项目语境

最终用户是财务/审计专业人士,**不是开发者**。本期 5 项改动全部围绕"让审计人员一打开 xlsm 就能看懂、能用"展开:

1. 删掉跨市场 BS/IS/CF 3 张冗余表,只留指标表
2. 样本池布局精简(取消"跨市场对比"独立 column,3 个全局按钮整齐摆同列)
3. 用户视角 doc 重写(去掉所有技术黑话)
4. 4 张重点 sheet 视觉规范统一(样本池 / 使用说明 / 跨市场_指标表 / 汇率)
5. tab 行为修正(诊断永隐 + 一键 X 股自动展开 + 抑制合并单元格弹窗 + 跨市场指标表独立 visibility)

零网络改动,纯本地工作。

## 用户已锁定的决策(2026-05-04 收集)

| # | 决策 |
|---|---|
| 跨市场对比表保留范围 | **只保留 跨市场_指标表 1 张**;BS/IS/CF 3 张删除 |
| 跨市场指标表 visibility | **独立**:不跟分市场 toggle,但跟"显示/隐藏 所有市场数据"全局 toggle 联动 |
| 跨市场指标表自动刷新 | **一键全抓 4 市场** 末尾自动刷;不需要单独的"一键跨市场对比"按钮 |
| 全局按钮集中 | 一键全抓 + 显示/隐藏 所有市场 + 清空 HTTP 缓存 **三个全部放在同一列**(Q 列)|
| Doc + 视觉范围 | 只改用户面向 doc + 视觉重点 4 张 |

## Step 总览

| Step | 内容 | 估时 |
|---|---|---|
| 1 | 删除跨市场 BS/IS/CF 3 张表 + 相关 VBA + 按钮 + 样本池布局精简 | 2h |
| 2 | 抑制合并单元格弹窗 + 一键 X 股自动展开 + 诊断永隐 + 跨市场指标表 visibility 独立 | 1.5h |
| 3 | 用户视角 doc 重写(使用说明 sheet + 汇率说明区 + Excel cell comments + README『使用方式』段)| 2h |
| 4 | 视觉规范 design system + apply 4 张重点 sheet(样本池 / 使用说明 / 跨市场_指标表 / 汇率)| 2h |
| 5 | 回归 + 新增 `inspect_phase4j_state.py` + STATUS §Y 收口 | 0.8h |

**总耗时 ~8h Codex 工作 + ~1h Reviewer**。1 round 端到端,commit `Phase 4j cross-market simplification + auditor docs + visual + tab behavior`。

---

## Step 1 — 跨市场表精简 + 样本池布局

### 1A. 删除跨市场 BS/IS/CF 3 张表

**1A.1** 从 xlsm 中删除 sheet:
- `跨市场_资产负债表`
- `跨市场_利润表`
- `跨市场_现金流量表`

**保留** `跨市场_指标表`(Phase 4g 交付,继续是核心跨市场对比 sheet)。

**1A.2** `tools/install_modules.py` 修改:
- `_make_cross_market_statement_sheet` helper 删除(只保留 `_make_cross_market_indicator_sheet`)
- `ensure_market_sheets` 删除 3 张表的 idempotent install
- `reorder_report_sheets` 的 `desired_order` 删除 3 张表
- 升级老版 xlsm 时,如果工作簿里还有这 3 张 sheet,**自动删除**(避免遗留数据混淆)

**1A.3** `tools/build_template.py` 修改:
- 删除 `build_cross_market_statement_sheet` 调用
- `main()` 末尾只 create `跨市场_指标表`(已有)

**1A.4** `modules/模块_工具函数.bas` VBA 改造:
- **保留** `BuildCrossMarketIndicatorSheet`(Phase 4g 主体,frozen)
- **删除** `BuildCrossMarketStatementSheet`(generic 通用 Sub)
- **删除** `BuildCrossMarketBalanceSheetWrapper` / `BuildCrossMarketIncomeWrapper` / `BuildCrossMarketCashFlowWrapper` 3 个 wrapper
- **删除** `BuildAllCrossMarketSheets`(改成在 `BuildCrossMarketIndicatorSheet` 之外不做其他事;或者简化为 alias)
- **保留** `MarketStatementSheetName` / `CollectCompaniesFromIndicatorSheet` 等 helper 如果还有使用方;否则删除

**1A.5** `模块_总入口.bas` 一键全抓 末尾的 `BuildAllCrossMarketSheets` 调用改成 `BuildCrossMarketIndicatorSheet`(直接调指标表合表 Sub,不走 wrapper)。

### 1B. 样本池布局精简

**1B.1** Q 列摆 3 个全局按钮(整齐同列):

```
Q1:Q3   一键全抓 4 市场                (蓝, PRIMARY_FILL)
Q5:Q7   显示/隐藏 所有市场数据         (浅蓝, SECONDARY_FILL)
Q9:Q11  清空 HTTP 缓存                 (浅蓝, SECONDARY_FILL)
```

**1B.2** 删除以下 BUTTONS 项(它们都对应已删的跨市场细分 / 独立按钮):
- `BtnBuildCrossInd`(Phase 4g 一键合并指标表 — 不需要,因为 一键全抓 自动刷)
- `BtnBuildCrossAll`(一键跨市场对比)
- `BtnHideCrossMarket`(显示/隐藏 跨市场对比)
- 任何 Phase 4i.1/4i.2 留下的跨市场细分按钮(BS/IS/CF wrapper)

**1B.3** 保留的样本池布局(完整):

```
列     A:B          E:F          I:J          M:N          Q 列(全局)
Row 1-6 配置区(年度/季度/cookie/币种/使用提示)
Row 7  A股(蓝)     美股(红)     港股(绿)     韩股(紫)
Row 8  一键 A 股   一键 美股    一键 港股    一键 韩股    Q1:Q3 一键全抓 4 市场
Row 9  显示/隐藏   显示/隐藏    显示/隐藏    显示/隐藏    Q5:Q7 显示/隐藏 所有市场数据
       A 股数据    美股数据     港股数据     韩股数据
Row 10 表头(代码/简称) ...                                Q9:Q11 清空 HTTP 缓存
Row 11+ 数据 ...

内联使用提示框: 保留在 O1:P5(Q 列被占,用 O:P 双列宽够装)
```

**1B.4** S / T / U 列全部清空(之前 Phase 4i.1/4i.2 用过的位置):
- 删除 S 列任何 leftover Shape
- 删除 T / U 列任何 leftover Shape

### 1C. 验证

- 装完打开 xlsm,sheet 数量 = `使用说明 / 样本池 / 4 市场 × 4 张 / 3 张诊断 / 跨市场_指标表 / 汇率 / 字段映射(Phase 4i 残留?)` 共 ~21 张(不再有跨市场 BS/IS/CF 3 张)
- 样本池 Q 列 3 个全局按钮整齐排列
- S/T/U 列 干净无 Shape leftover
- 跑一键全抓,跨市场_指标表 自动刷新

### 1D. Generator 不要做

- ❌ 不要删除 `跨市场_指标表`(Phase 4g 核心交付)
- ❌ 不要删除 `BuildCrossMarketIndicatorSheet` 主体逻辑
- ❌ 不要清掉 18 张分市场 sheet 的数据
- ❌ 不要动 4 个 一键 X 股 主入口

---

## Step 2 — tab 行为修正 + 抑制弹窗

### 2A. 抑制合并单元格弹窗

**2A.1** 用户截图显示生成指标表时弹出"合并单元格时,仅保留左上角的值,而放弃其他值"对话框。这是 Excel `Application.DisplayAlerts = True`(默认)在 cell merge 时触发的警告。

**2A.2** 修复:`模块_工具函数.bas` 中所有调用 `.Range(...).Merge` 或 `.Merge()` 的地方,前后包 `Application.DisplayAlerts = False/True`。常见位置:
- `BuildCrossMarketIndicatorSheet`(R1 公司名跨期合并)
- `WriteWideTable`(R1 公司名 跨期合并)
- `BuildStandardIndicatorSheet`(R1 公司名跨期合并)
- 任何写表 Sub 涉及合并单元格的位置

**2A.3** 推荐 pattern(避免漏抑制):

```vba
Public Sub <写表 Sub>(...)
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False    ' 抑制合并 / 删除 sheet / 覆盖等所有警告
    On Error GoTo CleanUp
    
    ' ... 现有写表逻辑(含 Merge / Clear / etc.)
    
CleanUp:
    Application.DisplayAlerts = prevAlerts    ' 还原, 不影响其他模块
End Sub
```

**2A.4** 注意:不要 globally 把 `Application.DisplayAlerts = False` 设到永远 — 用户在 Excel 里手动操作时还是要看到正常警告。只在 写表 Sub 范围内抑制。

### 2B. 诊断 sheet 永远不参与 toggle 显隐

**2B.1** `模块_总入口.bas` 的 `ToggleMarketTabsVisibility(market)` 函数加诊断 sheet 排除条件:

```vba
Private Sub ToggleMarketTabsVisibility(ByVal market As String)
    Dim prefix As String
    Select Case UCase$(Trim$(market))
        Case "A":  prefix = "A股_"
        Case "US": prefix = "美股_"
        Case "HK": prefix = "港股_"
        Case "KR": prefix = "韩股_"
        Case Else: Exit Sub
    End Select

    ' Phase 4j Step 2: 排除诊断 sheet (诊断 sheet 永远 hidden, 排查时手工右键 unhide)
    Dim newVisible As Long: newVisible = -1
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(prefix)) = prefix _
           And InStr(ws.Name, "抓取诊断") = 0 Then    ' ← 新增排除条件
            If ws.Visible = -1 Then newVisible = 0
            Exit For
        End If
    Next ws

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(prefix)) = prefix _
           And InStr(ws.Name, "抓取诊断") = 0 Then    ' ← 新增排除条件
            ws.Visible = newVisible
        End If
    Next ws
    Err.Clear
    On Error GoTo 0
End Sub
```

### 2C. 一键 X 股 抓数末尾自动 unhide 对应市场 4 张正式表

**2C.1** `模块_总入口.bas` 4 个一键 Sub(`一键A股` / `一键美股` / `一键港股` / `一键韩股`)末尾追加调用 `UnhideMarketTabs(market)`:

```vba
Public Sub 一键A股()
    ' ... 现有抓数逻辑不动
    
    ' Phase 4j Step 2: 抓完自动 unhide A 股 4 张正式表 (诊断 sheet 保持 hidden)
    UnhideMarketTabs "A"
    
    ' ... 现有 ShowMarketRunSummary 等也不动
End Sub
```

**2C.2** 新增 helper `Public Sub UnhideMarketTabs(ByVal market As String)`(放在 `ToggleMarketTabsVisibility` 旁边):

```vba
Public Sub UnhideMarketTabs(ByVal market As String)
    Dim prefix As String
    Select Case UCase$(Trim$(market))
        Case "A":  prefix = "A股_"
        Case "US": prefix = "美股_"
        Case "HK": prefix = "港股_"
        Case "KR": prefix = "韩股_"
        Case Else: Exit Sub
    End Select

    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(prefix)) = prefix _
           And InStr(ws.Name, "抓取诊断") = 0 Then    ' 同样排除诊断
            ws.Visible = -1    ' xlSheetVisible
        End If
    Next ws
    Err.Clear
    On Error GoTo 0
End Sub
```

### 2D. 跨市场_指标表 visibility 模型

**关键决策(用户明确)**:
- 单市场 toggle(显示/隐藏 X 股数据)→ **不动** 跨市场_指标表
- 全局 toggle(显示/隐藏 所有市场数据)→ **包含** 跨市场_指标表
- 一键 X 股(单市场)→ **不动** 跨市场_指标表 visibility(单市场抓数完成后,跨市场对比可能尚未刷新,不主动展示)
- 一键全抓 4 市场 → **自动 unhide** 跨市场_指标表 + auto refresh(已是 Phase 4g 行为)

**2D.1** 修改 `模块_总入口.bas` 的 `切换所有分市场tabs` Sub:

```vba
Public Sub 切换所有分市场tabs()
    ' Phase 4j Step 2: 全局 toggle 含跨市场_指标表
    Dim m As Variant
    For Each m In Array("A", "US", "HK", "KR")
        ToggleMarketTabsVisibility CStr(m)
    Next m
    
    ' 跨市场_指标表 也跟着 toggle (用一个独立 helper)
    ToggleCrossMarketIndicatorVisibility
End Sub

Private Sub ToggleCrossMarketIndicatorVisibility()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("跨市场_指标表")
    If ws Is Nothing Then Exit Sub
    
    ' 跟随分市场最新状态: 如果 A股_资产负债表 是 hidden,跨市场也 hidden;反之亦然
    Dim probeWs As Worksheet: Set probeWs = ThisWorkbook.Sheets("A股_资产负债表")
    If probeWs Is Nothing Then Exit Sub
    ws.Visible = probeWs.Visible
    Err.Clear
    On Error GoTo 0
End Sub
```

**2D.2** 修改 `模块_总入口.bas` 一键全抓末尾(已有 BuildCrossMarketIndicatorSheet 调用):

```vba
Public Sub 一键全抓(...)
    ' ... 4 市场抓数逻辑
    
    ' Phase 4j Step 2: 抓完自动 unhide 4 市场 16 张正式表 + 跨市场_指标表
    UnhideMarketTabs "A"
    UnhideMarketTabs "US"
    UnhideMarketTabs "HK"
    UnhideMarketTabs "KR"
    On Error Resume Next
    ThisWorkbook.Sheets("跨市场_指标表").Visible = -1
    Err.Clear
    On Error GoTo CleanUp
    
    ' Phase 4g 已有: 自动刷新跨市场指标表
    On Error Resume Next
    BuildCrossMarketIndicatorSheet
    Err.Clear
    On Error GoTo CleanUp
    
    ' ...
End Sub
```

### 2E. 验证

- 装完后 3 张诊断 sheet 默认 hidden ✓
- 点『显示/隐藏 美股数据』,只 toggle 美股 4 张正式表(BS/IS/CF/Indicator),`美股_抓取诊断` 始终 hidden,跨市场_指标表 不动
- 点『显示/隐藏 所有市场数据』,toggle 16 张正式表 + 跨市场_指标表(共 17 张)
- 点『一键 美股』,抓数完成后 美股 4 张正式表自动 visible(若之前 hidden),跨市场_指标表不动
- 点『一键全抓 4 市场』,抓数完成后 16+1 张正式表自动 visible + 跨市场_指标表自动刷新
- 跑任意一键 X 股,**不弹**"合并单元格仅保留左上角"对话框

### 2F. Generator 不要做

- ❌ 不要 globally `Application.DisplayAlerts = False` (只在写表 Sub 范围内)
- ❌ 不要把诊断 sheet 装表时 default visibility 改成 visible
- ❌ 不要让单市场一键 X 股影响跨市场_指标表 visibility(跟用户决策矛盾)

---

## Step 3 — 用户视角 doc 重写

### 3A. 重写原则

**目标读者**:有 5+ 年经验的财务/审计人员,熟悉财务报表口径,**不熟悉**:
- 编程术语(API / cache / fallback / fuzzy match / shell-out / hook / instance)
- 网络术语(HTTP / cookie / TLS / encoding / throttle)
- VBA 术语(Sub / Wrapper / OnAction / Shape)

**翻译规则**:
| 技术语 → 业务语 |
|---|
| API / endpoint → 数据接口 |
| cache / 缓存 → 本地暂存(下次抓数会复用,加快速度)|
| fallback → 备用数据源 |
| fuzzy match → 容差匹配(应对周末漂移、财年口径微调)|
| shell-out → (不出现,改成"通过本地脚本拉取")|
| FX rate / 汇率换算 → "原币换算成人民币的折算"|
| RMB 短路 → "本身就是人民币口径,无需折算"|
| WinHttp → (不出现,改成"通过本地脚本拉取")|
| toggle → 开关 / 切换 |
| Worksheet_Change → (不出现,改成"修改后立即生效")|

### 3B. 改动范围

**3B.1 使用说明 sheet**(完整重写 7 section):
- § 1 项目概览 → "对比 4 个证券市场上市公司财务表现的桌面工具,3 分钟内自动完成所有公司同期数据的抓取、单位换算、跨表对齐"
- § 2 快速开始 → 6 步流程
- § 3 输出 sheet 说明 → 把每张 sheet 用 1 句"它是干嘛的"描述,不展开技术细节
- § 4 数据来源声明 → 4 市场各 1 句话,不说技术 endpoint
- § 5 汇率换算说明 → "本工具会自动从公开渠道获取 4 大币种(USD/HKD/KRW/CNY)历史汇率,按报告期换算。资产负债表用期末汇率,利润表/现金流量表用期间均值汇率(符合审计准则)"
- § 6 常见问题 → 5 个常见 Q 用业务话改写
- § 7 版本历史 → 简表

**3B.2 汇率说明区**(R10+):删掉所有技术词,改成业务话

**3B.3 Excel cell comments**:
- 样本池 关键 cell(B5 cookie / B6 currency / 一键按钮 / 全局按钮)comments 全部审核
- 跨市场_指标表 A1 注释 → 用业务话(eg "本表合并 4 个市场 18 项标准财务指标的对比视图;每次跑数后由『一键全抓 4 市场』自动刷新")
- 3 张诊断 sheet A1 注释保留技术细节(诊断本来就是排查用的)

**3B.4 README.md『使用方式』段**:同 3B.1 重写

### 3C. Generator 不要做

- ❌ 不要改 STATUS.md / PHASE_*_PLAN.md(开发文档,保留追溯价值)
- ❌ 不要改 commit message / VBA 注释 / Python docstring(开发面向)
- ❌ 不要因为追求"业务话"而失去技术准确性(eg 不能把"期末汇率"翻译成"那一天的汇率")

---

## Step 4 — 视觉规范 + 4 张重点 sheet

### 4A. Design System(全工作簿统一)

**配色**:
- 主色 / 标题底色:`#1F3864`(深蓝,封面 + 大标题)
- 次色 / section header:`#4472C4`(中蓝,所有 section 标题底色)
- 强调色 / 表格 header:`#D9E1F2`(浅蓝,表格表头)
- 数据正区:`#FFFFFF`(白)
- 数据替代行:`#F2F2F2`(浅灰,zebra striping,可选)
- 警告 / 备注:`#FFF2CC`(浅黄,内联使用提示框)
- 4 市场标识色保留(蓝/红/绿/紫)
- 跨市场_指标表 sheet tab 颜色 = `PRIMARY_FILL` 蓝(跟 A 股同色,因为也是综合视图)— **本期不改 Phase 4i.1 染色**

**字体**:微软雅黑全局,封面 24/14/10,section header 14,正文 11,表格 10,备注 9
**行高**:封面 36 / 二级标题 24 / 三级 20 / 正文 18 / 表格 16
**边框**:表格内部 weight 2,外部 weight 3
**对齐**:表头居中,数字右对齐,文字左对齐,label 加粗

### 4B. 4 张重点 sheet apply 范围

**4B.1 样本池**:
- 配置区(R1-R6):统一字体 + 配色 + 行高
- 4 市场分栏(R7-R10):4 市场底色保留
- 数据区(R11+):字体统一,zebra striping 可选
- 操作区(Q 列):section label 浅灰底 + 加粗黑字 + 居中
- 内联使用提示框(O1:P5):浅黄底 + 微软雅黑 9
- 整体冻结 SplitRow=10(已有)

**4B.2 使用说明**:Step 3 已重写,本 step 只 cross-check 视觉一致性

**4B.3 跨市场_指标表**(原计划是 4 张,现在只有 1 张):
- 表头 R1-R2:深蓝底白字 + 公司名加粗
- 数据区(R3-R20,18 标准指标):正区背景 + 字体 11 + 数字右对齐
- 列宽:A=12(指标类型)/ B=18(指标名)/ C=24(英文指标名)/ D+=15.875(数据列)
- 冻结:R3,D3
- A1 注释:Step 3 已改

**4B.4 汇率**:
- R1-R3 数据区:统一字体 + 行高
- R10+ 说明区(Step 3 已改业务话):section header 配色 + 子节标题 12 加粗 + 正文 11
- 列宽:A=18 / B=60 / C-H 自动

### 4C. Generator 不要做

- ❌ 不要 apply 视觉规范到 16 张分市场 sheet(超范围)
- ❌ 不要破坏 Phase 4i.1 sheet tab 染色(4 市场色 + 跨市场蓝保留)

---

## Step 5 — 回归 + 验证 + STATUS §Y 收口

### 5A. 跑 frozen 回归

```bash
cd "VBA Captor"
py tools/test_fx_live.py --skip-install
py -u tools/diff_phase4f_step3_lite.py
py -u tools/inspect_phase4g_state.py
py -u tools/inspect_phase4h_state.py
```

期望 4/4 PASS。Phase 4j 不动 fetch/cache/write 业务逻辑,任一退化是 bug。

**注意**:Phase 4h inspect 之前可能验证过跨市场 BS/IS/CF 3 张表存在,Step 1 删了之后这个断言要更新 — 在 inspect 文件加兼容逻辑(`if sheet exists then check, else skip`),不要硬退化。

### 5B. 新增 `tools/inspect_phase4j_state.py`

检查项:
- 跨市场 BS/IS/CF 3 张表 **不存在**(被 Step 1 删除)
- 跨市场_指标表 仍存在 + R1 公司 + R2 报告期 + R3-R20 18 行标准指标
- 样本池 Q 列只有 3 个全局按钮(BtnRunAll / BtnHideAll / BtnClearCache),S/T/U 列无 Shape
- 一键 X 股 自动 unhide 该市场 4 张表(模拟 unhide → 跑 一键 A 股 → 检查 A股_* 4 张 visible)
- 显示/隐藏 所有市场数据 toggle 含跨市场_指标表(toggle 一次 17 张全 hidden,再 toggle 17 张全 visible)
- 诊断 sheet 任何 toggle 后保持 hidden
- 写表过程中无弹窗(Application.DisplayAlerts 在写表 Sub 范围内 = False)
- 视觉规范:抽样 5 个 cell 验证字体 / 字号 / 配色

### 5C. STATUS §Y 收口

```markdown
## Y. Phase 4j 收口: 跨市场表精简 + 用户视角 doc + 视觉规范 + tab 行为修正

执行依据: PHASE_4J_PLAN.md v2。状态: ✅ Codex 已实现并通过 4 张 frozen 回归 + 新增 inspect。

### Y.1 已完成
- [Step 1] 删除跨市场 BS/IS/CF 3 张表 + 相关 VBA wrapper + 样本池 Q 列精简到 3 个全局按钮
- [Step 2] 抑制合并单元格弹窗(Application.DisplayAlerts 范围内 = False) + 诊断永隐 + 一键 X 股 自动展开 + 跨市场指标表独立 visibility
- [Step 3] 用户视角 doc 重写: 使用说明 sheet 7 section + 汇率说明区 + 关键 cell comments + README 使用方式
- [Step 4] 视觉规范 design system + apply 样本池 / 使用说明 / 跨市场_指标表 / 汇率 = 4 张重点 sheet

### Y.2 验证结果
[5A 回归 + 5B inspect 结果 + 手工抽查 cell 视觉]

### Y.3 已知边界
- 跨市场 BS/IS/CF 行项目对齐工作 永久 defer(Phase 4h §W.3 提出 Phase 4h.1/4i 候选 mapping 工作 撤销, 用户决策不再做)
- 视觉规范本期只 apply 4 张重点 sheet;16 张分市场报表保留 Phase 4i.1 现状
- doc 重写不含 STATUS / PHASE_*_PLAN(保留开发追溯价值)
- 跨市场指标表 visibility 独立: 不跟单市场 toggle, 跟全局 toggle 联动, 一键全抓自动 unhide
```

PHASE_4J_PLAN.md v1 → v2,标记 ✅。

---

## ⚠️ 全 Phase 严禁动的东西

| 文件/区域 | 原因 |
|---|---|
| 任何 fetch / cache / WriteWideTable VBA 业务逻辑 | Phase 4j 是 sheet 精简 + doc + 视觉 + tab 行为 |
| `BuildCrossMarketIndicatorSheet` 主体 | Phase 4g frozen,Step 1 只删 BS/IS/CF 通用化的 wrapper |
| 4 张回归驱动 + Phase 4h inspect 核心断言 | frozen(Phase 4h inspect 关于跨市场 BS/IS/CF 的断言加 if-exists 兼容) |
| Phase 4d 韩股 stockanalysis 解析 | frozen |
| 16 张分市场 BS/IS/CF/Indicator 表 视觉 | 超范围 |
| STATUS.md / PHASE_*_PLAN.md 内容 | 开发文档 |
| 4 个 一键 X 股 主入口的抓数逻辑 | 只在末尾追加 unhide 调用,不动主体 |

## ⚠️ 联系 Planner 触发条件

- Step 1 删除跨市场 BS/IS/CF 3 张表后,任何残留 VBA 调用(grep 找不全)→ 标 `[BLOCKED]` 报告
- Step 2 抑制弹窗后,某 写表 Sub 在出错 path 没还原 `DisplayAlerts`(globally 永远 = False)→ 立即修
- Step 4 视觉规范 apply 后任一公式被误删 → 立即停下回滚
- Step 5 4 张 frozen 回归任一退化(Phase 4h inspect 关于 BS/IS/CF 3 张的断言更新不算退化,要在 commit msg 说明)

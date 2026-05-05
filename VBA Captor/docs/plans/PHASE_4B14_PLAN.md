# Phase 4b-14a: 美股字段映射 — 高覆盖 + 优雅降级 + 明确诊断

> **版本**: v3(2026-05-03 EOD,经 Codex 两轮 review 收敛)
> **状态**: ✅ **已实现并通过 Codex 端到端验证**(待 Claude Code review) — **本文件是 4b-14 唯一 source of truth**(STATUS.md §N 已废弃,见 §文档关系)
> **作者**: Claude(初稿) + Codex(review)
> **背景文档**: `STATUS.md` §M(Phase 4b-13 收口);§N 已废弃
> **重要变更**: v2/v3 修订纪要见下方

## 进度追踪(2026-05-03 工作中)

| 工作项 | 状态 | 责任人 | 备注 |
|---|---|---|---|
| `MapEntry*` helpers (5 个) | ✅ 已实现 | Codex | `MapEntryUsGaapConcepts/Ifrs/Unit/Scale/FuzzyHint` |
| `MapEntryCategory` / `MapEntryLabel` | ✅ 已实现 | Claude | 补 Codex 漏定义,在 工具函数模块 |
| `CoreLabelsForKind` | ✅ 已实现 | Codex | BS=Total assets / IS=Revenue+Net income / CF=Cash from operations+Cash at end |
| `EnsureDiagnosticSheet` / `ClearDiagnosticSheet` / `DeleteDiagnosticRowsForKind` / `AddDiagnosticRow` / `WriteDiagnosticForKind` | ✅ 已实现 | Codex | 工具函数模块,带表头初始化 + 10 列布局 |
| `AddFuzzyDiagnosticCandidates` + 10 个 fuzzy helpers (Tokenize/Score/HasNegative 等) | ✅ 已实现 | Codex | 工具函数模块, top-3 候选,score >=5 才推荐 |
| `AddMissingDiagnosticsForCompany` | ✅ 已实现 | Claude | Public 版本(Codex 删了重复 Private 版本) |
| **三级递进 fetch 主流程** (`FetchAndAccumulateUSCompany` 重写) | ✅ **已验证** | Codex | tempData 暂存 + Tier 1/2/3 + 核心字段判断 + 原子 commit |
| `AccumulateEdgarTaxonomy` / `CanonicalEdgarEntries` / `HasAnyCoreLabel` / `CommitTempUSData` / `AppendDiagnosticsForConceptMap` / `AppendNonUsdDiagnostics` | ✅ 已实现 | Codex | fetch 模块 Private helper |
| `FetchUSFromXueqiu` 加 `collDiagRows` 参数 + 雪球命中 emit `OK_XUEQIU` 行 | ✅ 已实现 | Codex | |
| `RunUSStatement` 创建 collDiagRows + WriteDiagnosticForKind | ✅ 已实现 | Codex | |
| `一键全抓` 设 g_diagnosticAppendOnly + ClearDiagnosticSheet | ✅ 已实现 | Codex | |
| **诊断框架(端到端)** | ✅ **已验证** | Codex+Claude | 一键全抓追加、单表按报表重写、10 列诊断均通过 |
| `tools/install_modules.py`: `_make_diagnostic_sheet` + `ensure_us_sheets` 加 `美股_抓取诊断` + `reorder_report_sheets` 加 sheet | ✅ 已实现 | Claude | 列宽/字色与 VBA 端 EnsureDiagnosticSheet 三处保持一致 |
| `tools/build_template.py`: `build_diagnostic_sheet` + 模板 sheet 列表加 `美股_抓取诊断` | ✅ 已实现 | Claude | |
| **Layer 1 concept 候选扩充** (BS/IS/CF + 雪球字段) | ✅ **已验证** | Codex | 保持原 primary 第一位;收窄高风险候选;补 BABA 财年 Q4 雪球匹配 |
| 编译/打包测试 + AAPL/AMZN/POM/HTT 回归 | ✅ 已完成 | Codex | install_modules.py 跑通;Test 1/Test 2/非法 ticker 均通过 |

**当前状态**: Codex 执行完成,等待 Claude Code code review。验证摘要见 `STATUS.md` §O。

## Context

**项目当前状态**: V3.xlsm 在 Phase 4b-13 收口,8 张表 + 一键全抓 + 备份稳定版都已交付。AAPL/AMZN/POM/HTT 4 家测试公司 100% 通过(STATUS.md §M)。

**这一阶段要解决什么**: 美股字段映射目前**完全靠 hardcode**:
- `模块_抓美股资产负债表.bas/GetBSConcepts()` — 27 个 us-gaap concept,3 元素 tuple,**单一 concept** 无候选(早期实现,落后于 CF 表)
- `模块_抓美股利润表.bas/GetISConcepts()` — 14 个 concept,**单一 concept** 无候选
- `模块_抓美股现金流量表.bas/GetCFConcepts()` — 34 个 concept,Phase 4b-10 已升级为**多候选** CSV 格式 ✅
- `模块_抓美股财报.bas/XueqiuFieldMapForKind()` — 雪球 fallback 字段表 BS 27 + IS 12 + CF 12 + Indicator 3,**全部多候选 CSV**

**真实场景的失败模式**: 用户随便抓一只新美股,可能命中 hardcode 命中率 < 100% 的两个洞:
1. **概念名变体**: AAPL 单家就有 4 个 Revenue 变体 (`Revenues` / `RevenueFromContractWithCustomerExcludingAssessedTax` / `SalesRevenueNet` / `SalesRevenueServicesGross`),不同财年用不同的。
2. **20-F filer 用 ifrs-full**: 部分 20-F 公司在 EDGAR 的 companyfacts 里走 `ifrs-full` 而非 `us-gaap`,我们目前**完全不查 ifrs-full**。

**用户纠正的方向(2026-05-03)**: 不是把映射拆成 JSON 让别人维护,而是**让我们的代码自己适应**。

## v2 修订纪要(Codex 第一轮 review 收敛点)

| # | Codex 原始批评 | 修订动作 |
|---|---|---|
| 1 | "开箱即用任意 ticker"过度承诺 | 目标改成"**高覆盖 + 优雅降级 + 明确诊断**",删除"100% 自适应"措辞 |
| 2 | 完全 hardcode 堵死开源后社区维护 | 本阶段维持 hardcode(用户原话);**新增 future work 一节**,标注"开源化时再上隐藏 sheet 或可选 JSON" |
| 3 | fuzzy 自动入账与勾稽校验脱节,矛盾 | **fuzzy 只写诊断推荐,不入账财报**;勾稽校验整体推迟 Phase 4b-14b |
| 4 | 7-tuple 主流程到处 UBound 难维护 | **封装 5 个 helper**:`MapEntryUsGaapConcepts/MapEntryIfrsConcepts/MapEntryUnit/MapEntryScale/MapEntryFuzzyHint`,主流程零 UBound |
| 5 | BABA 非 USD 强转风险 | 单位非 USD → `MISSING_NON_USD` 状态,不强写;触发雪球 fallback |
| 6 | 诊断表 7 列不够 | **加 3 列变 10 列**:Taxonomy / Unit / Score |

## v3 修订纪要(Codex 第二轮 review 收敛点)

| # | Codex 第二轮批评 | 修订动作 |
|---|---|---|
| 1 | 诊断表清空/追加机制不明 — 一键全抓 4 张表会被最后一张覆盖 | **一键全抓开头清空一次,4 张美股表追加;单独点按钮按 strKind 过滤删行重写**(详见 §诊断 sheet 写入规则) |
| 2 | fallback 触发语义未写死 — `matchedCount=0` 太粗暴,应该看核心字段 | **核心字段清单写死**:BS=Total assets / IS=Revenue+Net income / CF=Cash from operations+Cash at end。核心字段全无或全非 USD 才整表 fallback |
| 3 | us-gaap 缺失时直接去雪球,跳过 ifrs-full | 改成**三级递进**:us-gaap → ifrs-full(单位严格 USD) → 雪球 |
| 4 | Test 1 要求"诊断全部 OK"过严 — POM 现金流字段稀疏 MISSING 合理 | 验收改成:**已输出字段必须 OK / OK_XUEQIU,未输出字段允许 MISSING(只要不报错或中断)** |
| 5 | `XYZ` 错代码不稳定(可能是某真公司) | 改用 `ZZZINVALID123` 这种明显非法值 |
| a | J 列备注信息不够 — 区分不出"concept 命中本期无值"和"完全没命中" | J 列追加 `periods_written=N` 元信息 |
| b | Layer 1 扩候选可能改写候选优先级,导致 AAPL/AMZN 回归不一致 | **Layer 1 写死约束:现有 concept 保持第一位**,新增别名追加在后 |
| c | STATUS.md §N 与 PHASE_4B14_PLAN.md 关系不清 | **本文件是 4b-14 唯一 source of truth**,STATUS.md §N 已废弃(实施完成后会在 STATUS.md §N 头部加废弃标 + §O 收口) |

## 设计目标(v2)

| 目标 | 说明 |
|---|---|
| **高覆盖** | 常见 us-gaap 别名 + ifrs-full 兜底,主流公司核心字段尽力抓全 |
| **优雅降级** | 单家公司部分字段缺失不影响其它字段;一张表全空触发雪球 fallback;雪球也无 → 弹窗失败但不卡死 |
| **明确诊断** | 每次抓数都把"用了哪个 concept、怎么匹配上的、单位、taxonomy、评分"写到 `美股_抓取诊断` sheet |
| **无侵入** | 不破坏 Phase 4b-13 已验证的 AAPL/AMZN/POM/HTT 输出 |
| **保守** | **fuzzy 命中只推荐,不入账**;勾稽校验推迟 4b-14b;非 USD 不强转 |
| **不依赖外部文件** | 全部逻辑在 VBA + .bas,V3.xlsm 仍是独立可分发文件(开源时另议) |

## 实施方案(分 3 层,本阶段全做)

### Layer 1 — 扩充已知候选(EDGAR + 雪球双侧)

**EDGAR 侧**: 把 `GetBSConcepts()` / `GetISConcepts()` 的 3-tuple 全部升级为 Phase 4b-10 同款多候选 CSV。基于业内常见的 us-gaap 别名,补全候选列表。

**v3 强约束(Codex 修订 #b)**: **现有 concept 必须保持第一位**,新增别名追加在 CSV 后面。这样 AAPL/AMZN 等已验证公司的命中结果完全不变,候选只在 primary 没命中时才生效。

样例 (Revenue):
```vba
' 旧
a(0) = Array("Revenue", "RevenueFromContractWithCustomerExcludingAssessedTax")
' 新 — 现有 concept 仍在第一位
a(0) = Array("Revenue", _
    "RevenueFromContractWithCustomerExcludingAssessedTax," & _   ' ← 保持第一位 (Phase 4b-13 primary)
    "Revenues," & _
    "SalesRevenueNet," & _
    "RevenueFromContractWithCustomerIncludingAssessedTax," & _
    "SalesRevenueGoodsNet")
```

需要扩的核心字段(按使用频率):
- BS: Cash & equivalents, Inventory, Total assets, Total liabilities, Total stockholders' equity, Goodwill, PP&E, Long-term debt
- IS: Revenue, COGS, Gross profit, Operating income, Net income, EPS

**雪球侧**: `XueqiuFieldMapForKind` 已经多候选了,但基于历史观察补几个高频字段。同样规则:**现有字段保持第一位**。

**预期收益**: 不改 schema、不动主流程。已有多候选机制(`FetchAndAccumulateUSCompany` 282-357 行 + `FetchUSFromXueqiu` 782-805 行)直接消费。

### Layer 2 — us-gaap → ifrs-full → 雪球 三级递进(v3 严格定义)

**v3 修订 #2/#3**: fallback 触发语义写死,不再"facts 缺 us-gaap 就直接雪球"。

**核心字段清单** (`Function CoreLabelsForKind(strKind) As Variant`,新增于 `模块_工具函数.bas`):
```vba
Select Case strKind
    Case "BalanceSheet": CoreLabelsForKind = Array("Total assets")
    Case "Income":       CoreLabelsForKind = Array("Revenue", "Net income")
    Case "CashFlow":     CoreLabelsForKind = Array("Cash from operations", "Cash at end of period")
    Case "Indicator":    CoreLabelsForKind = Array()  ' 4b-9 起指标表纯标准指标,不进 fallback 体系
End Select
```

**三级递进流程**(替换原 `FetchAndAccumulateUSCompany` 220-252 段):

```
EDGAR HTTP 200
  ↓
[Tier 1] 试 facts.us-gaap 全候选 (per concept) → 写入 dictData
  ↓
[Tier 2] 检查核心字段是否至少 1 个被写入
  ├ YES → 完成,补全的字段后续按 concept 循环正常处理 (per-field MISSING 写诊断)
  └ NO  → 进入 Tier 2 ifrs-full
            ↓
            试 facts.ifrs-full 全候选 (per concept,严格 unit=USD)
            ↓
            非 USD 命中 → 诊断 MISSING_NON_USD,不入账
            ↓
            重新检查核心字段
              ├ YES → 完成 (诊断标 OK_IFRS)
              └ NO  → Tier 3 雪球 fallback
                        ↓
                        FetchUSFromXueqiu (现有逻辑)

EDGAR HTTP 404 / facts 缺失 → 直接 Tier 3 雪球
```

**关键设计点**:
- **不做 per-field 混合填充**(Codex 修订 #2):同一公司同一报表要么走 EDGAR(可能 us-gaap+ifrs-full 混合),要么走雪球,不混源
- 进入 Tier 3 雪球前,**清空当前公司当前报表已写入的 dictData 数据**(避免 EDGAR 半截数据 + 雪球数据矛盾)
- ifrs-full 的 unit 严格只取 USD,不强转
- conceptMap 升级到 6/7-tuple,新增 `ifrs_concepts` 字段(可选,缺省时复用 us-gaap concept 名)

```vba
' 不指定 ifrs 时, 系统自动用 us-gaap concept 名也试一次 ifrs (多数 IFRS taxonomy 同名)
a(12) = Array(_
    "",                              ' category
    "Total assets",                  ' label
    "Assets",                        ' us-gaap concepts CSV (Layer 1)
    "USD",                           ' unit (5-tuple 既有)
    1000000#)                        ' scale (5-tuple 既有)

' 指定不同 ifrs 名时显式 6-tuple
a(12) = Array(_
    "",
    "Total assets",
    "Assets",
    "USD",
    1000000#,
    "Assets,AssetsTotal")            ' index 5: ifrs-full concepts CSV (Layer 2)
```

### Layer 3 — 自动发现(只推荐到诊断,不入账)

**Codex 修订 #3**: 原计划 fuzzy 命中后写入财报。本版改为:**fuzzy 仅推荐,不写财报**。

**新增 sheet** `美股_抓取诊断`,**10 列**(Codex v2 修订 #6 加 Taxonomy/Unit/Score):

| 列 | 字段 | 示例 |
|---|---|---|
| A | 公司 | `AAPL` |
| B | 报表 | `BalanceSheet` |
| C | 输出指标 | `Total assets` |
| D | 状态 | `OK` / `OK_IFRS` / `OK_XUEQIU` / `MISSING` / `MISSING_NON_USD` / `RECOMMEND_FUZZY` |
| E | 数据源 | `EDGAR us-gaap` / `EDGAR ifrs-full` / `Xueqiu` / `—` |
| F | **Taxonomy** | `us-gaap` / `ifrs-full` / `xueqiu` / `—` |
| G | 命中字段 | `Assets` (实际用的 concept,fuzzy 推荐时也写) |
| H | **Unit** | `USD` / `EUR` / `CNY` / `—` |
| I | **Score** | `100` (hardcoded primary) / `85` (hardcoded alt) / `7.5` (fuzzy candidate) / `—` |
| J | 匹配方式+备注 | `hardcoded_primary; periods_written=4` / `hardcoded_alt[2/5]; periods_written=2` / `fuzzy_candidate (人肉确认后回填到 hardcode)` |

**v3 修订 #a — J 列必须含 `periods_written=N`**: 区分"concept 命中但本期没值" vs "完全没命中"。所有 OK/OK_IFRS/OK_XUEQIU 状态都必须在 J 列尾部带 `; periods_written=N`(N = 该 concept 在本次抓数实际写入的报告期数);MISSING 状态 J 列写原因不写 periods_written。

### 诊断 sheet 写入规则(v3 修订 #1 — 关键)

避免一键全抓 4 张表彼此覆盖,定义两个写入模式:

**模式 A — 全清重写(单表运行)**:
- 触发: 用户单独点『更新美股资产负债表』/『更新美股利润表』/...
- 行为: 在 `RunUSStatement` 写诊断时,先 ClearContents 数据区中 `B 列 == strKind` 的所有行(其它表的诊断保留),再 append 本次结果

**模式 B — 累积追加(一键全抓)**:
- 触发: 用户点『一键全抓』
- 行为:
  - 一键全抓入口 `Sub 一键全抓()` **开头清空整张诊断 sheet 一次** (R3:J1000)
  - 后续 4 张美股表运行时,**纯 append 不清空** (跳过模式 A 的过滤删行步骤)

**实现**(`模块_工具函数.bas` 新增 module-level flag):
```vba
Public g_diagnosticAppendOnly As Boolean    ' True = 模式 B; False (default) = 模式 A
```

`一键全抓` 入口:
```vba
Sub 一键全抓()
    g_diagnosticAppendOnly = True
    ClearDiagnosticSheet                    ' 全清一次
    ' ... 调 4 张美股表 Main ...
    g_diagnosticAppendOnly = False           ' 复位
End Sub
```

`WriteDiagnosticForKind(strKind, dictRows)` 内部:
```vba
If Not g_diagnosticAppendOnly Then
    DeleteDiagnosticRowsForKind strKind     ' 模式 A: 按 B 列过滤删行
End If
AppendDiagnosticRows dictRows                ' 都执行
```

**自动发现算法** (`Function FuzzyMatchConceptCandidates(taxonomy, label, fuzzyHint, expectedUnit) As Variant`):

1. 以 label 为种子,按规则 split 出关键词集 (e.g. "Total assets" → `["Total", "Assets"]`,过滤停用词)
2. 遍历 taxonomy 全部 keys,对每个 concept 名打分:
   - **关键词命中**: 包含全部关键词 +5,部分命中 +1/词
   - **单位匹配**: 单位 = expectedUnit +3
   - **排除规则**: 命中 fuzzyHint 里负向关键词 → 直接放弃
3. **不取 top 1 写入财报**;返回 top 3 candidates(score, concept, unit)写到诊断 sheet
4. **status = `RECOMMEND_FUZZY`**(明确表示这是推荐不是已采用)

**用户工作流**:
- 跑完看诊断 sheet
- 发现某指标 `MISSING` 但有 `RECOMMEND_FUZZY` 推荐
- 人肉确认推荐合理 → 把 concept 添加到对应 .bas 的 hardcode 列表 → 下次跑会变 `OK`(从 `RECOMMEND_FUZZY` 升级)

conceptMap 7-tuple 最终形态(可选):
```vba
a(12) = Array(_
    "",
    "Total assets",
    "Assets",
    "USD",
    1000000#,
    "Assets,AssetsTotal",                                    ' index 5: ifrs-full
    "Asset|Total~Current,Noncurrent,Intangible,Goodwill")    ' index 6: fuzzy hint (positive|positive~negative,negative)
```

### MapEntry helper 封装(Codex 修订 #4)

新增于 `模块_工具函数.bas`:
```vba
Public Function MapEntryUsGaapConcepts(ByVal entry As Variant) As String
    MapEntryUsGaapConcepts = CStr(entry(2))
End Function

Public Function MapEntryIfrsConcepts(ByVal entry As Variant) As String
    ' 6-tuple 起有 ifrs;否则复用 us-gaap (多数同名)
    If UBound(entry) >= 5 Then
        MapEntryIfrsConcepts = CStr(entry(5))
    Else
        MapEntryIfrsConcepts = CStr(entry(2))
    End If
End Function

Public Function MapEntryUnit(ByVal entry As Variant) As String
    ' 5-tuple 起有 unit
    If UBound(entry) >= 4 Then
        MapEntryUnit = CStr(entry(3))
    Else
        MapEntryUnit = "USD"
    End If
End Function

Public Function MapEntryScale(ByVal entry As Variant) As Double
    If UBound(entry) >= 4 Then
        MapEntryScale = CDbl(entry(4))
    Else
        MapEntryScale = 1000000#
    End If
End Function

Public Function MapEntryFuzzyHint(ByVal entry As Variant) As String
    If UBound(entry) >= 6 Then
        MapEntryFuzzyHint = CStr(entry(6))
    Else
        MapEntryFuzzyHint = ""
    End If
End Function
```

主流程改造:
```vba
' 旧
strConcept = CStr(mapEntry(2))
If UBound(mapEntry) >= 4 Then
    strUnit = CStr(mapEntry(3))
    dblScale = CDbl(mapEntry(4))
Else
    strUnit = "USD": dblScale = 1000000#
End If

' 新
strConcept = MapEntryUsGaapConcepts(mapEntry)
strIfrs = MapEntryIfrsConcepts(mapEntry)
strUnit = MapEntryUnit(mapEntry)
dblScale = MapEntryScale(mapEntry)
strFuzzy = MapEntryFuzzyHint(mapEntry)
```

## 不做的事(明确边界)

- ❌ **不做外部 JSON / mapping 文件**(本阶段;Codex 标 future work,见 §future work)
- ❌ **不做勾稽校验**(整体推到 Phase 4b-14b)
- ❌ **fuzzy 不写入财报**(只写诊断推荐;Codex 修订 #3)
- ❌ **不做 derive_from(Total = sum of children)的自动派生**(留 4b-14b)
- ❌ **不做 SEC filing presentation linkbase 解析**(VBA 不适合,留作未来 Python helper)
- ❌ **不做雪球字段的 fuzzy 自动发现**(Layer 1 扩充候选已经够,且 fuzzy 在 snake_case 上反而风险大)
- ❌ **非 USD 单位不强转**(Codex 修订 #5;明确 MISSING_NON_USD,触发雪球 fallback)

## 文件改动清单

| 文件 | 改动类型 | 责任人 |
|---|---|---|
| `modules/模块_抓美股资产负债表.bas` | 重写 GetBSConcepts | **Codex (Layer 1)** |
| `modules/模块_抓美股利润表.bas` | 重写 GetISConcepts | **Codex (Layer 1)** |
| `modules/模块_抓美股现金流量表.bas` | 升级 GetCFConcepts | **Codex (Layer 1)** |
| `modules/模块_抓美股财报.bas` (`XueqiuFieldMapForKind`) | 雪球字段补 | **Codex (Layer 1)** |
| `modules/模块_工具函数.bas` (新增 `MapEntry*`) | 新增 helpers | **Claude (Layer 2/3)** |
| `modules/模块_抓美股财报.bas` (`FetchAndAccumulateUSCompany`) | ifrs-full + fuzzy 调用 | **Claude (Layer 2/3)** |
| `modules/模块_工具函数.bas` (新增 fuzzy + diagnostic) | 新增 fuzzy match 函数 + 诊断 writer | **Claude (Layer 2/3)** |
| `tools/build_template.py` | 加 sheet + 10 列表头 | **Claude (Layer 3)** |
| `tools/install_modules.py` | 加 sheet 维护 | **Claude (Layer 3)** |

**并行不冲突保证**: Codex 只动 Layer 1 的 conceptMap entries(扩 us-gaap 候选 + 可选加 ifrs CSV);Claude 只动主流程逻辑 + helper 函数 + 工具脚本。Layer 2/3 通过 helper 封装兼容所有 tuple shape,所以 Codex 的 Layer 1 可以渐进 commit。

## 关键代码复用(不重新发明)

- **多候选 CSV 解析**: 已有 `Split(strConcept, ",")` 模式 (`模块_抓美股财报.bas:284`),直接复用
- **canonical entry 选择** (max end + min start): 已有 (`模块_抓美股财报.bas:298-326`),保持
- **JsonConverter.ParseJson**: 已稳定使用,不变
- **ResolveMarket / ReadXueqiuCookie / EdgarHttpGet / XueqiuHttpGet**: 全部不动
- **WriteWideTable**: 不动,诊断 sheet 用专门的 writer

## 验收方案(端到端测试)

### Test 1 — 已有公司回归(v3 修订 #4)
样本池: AAPL, AMZN, POM, HTT
配置: `A2=2025, A4=全部`
跑: 一键全抓
**期望**(放宽:不再要求"诊断全部 OK"):
- ✅ 8 张正式输出表数据列填充与 Phase 4b-13 稳定版备份**完全一致**(diff = 0)— 这是硬指标
- ✅ 已输出字段诊断状态必须是 `OK` / `OK_IFRS` / `OK_XUEQIU` 之一,不能是 `MISSING`
- ⚠️ **未输出字段允许 `MISSING`**(POM 现金流字段稀疏属正常),只要不报错或中断
- ✅ 不出现任何公式错误或弹窗失败 > 0

### Test 2 — 新公司探索(验证自适应能力)
样本池: 加入 5 只之前没测过的 — `MSFT 微软`, `GOOGL 谷歌`, `TSLA 特斯拉`, `NVDA 英伟达`, `BABA 阿里巴巴 (20-F)`
配置: `A2=2024, A4=Q4`
跑: 一键全抓
**期望**(Codex 修订 #5: 不再要求 BABA 数据填充,只观察诊断行为):
- MSFT/GOOGL/TSLA/NVDA: 美股 BS 至少 Total assets 抓到(可能命中 hardcode 主候选 OR 备选)
- BABA(20-F):
  - 走 ifrs-full 命中且单位是 USD → 数据填充 + 诊断显示 `OK_IFRS`
  - 走 ifrs-full 命中但单位是 CNY → 诊断显示 `MISSING_NON_USD`,触发雪球 fallback;诊断额外行显示 `OK_XUEQIU` 或 `MISSING`
  - **任何一种结果都算"通过",验收点是诊断 sheet 真实反映了发生了什么**
- 任何 fuzzy 推荐的 concept 名都记在诊断 G 列(`RECOMMEND_FUZZY` 状态),供我们事后回填到 hardcode

### Test 3 — 边界 case(v3 修订 #5)
- 故意填一个明显非法代码 `ZZZINVALID123` → 失败但不中断 + 诊断 sheet 整行 `MISSING`
- 一家 EDGAR 完全空 + 雪球也无的 → 优雅降级,诊断 sheet 全 `MISSING` 行,弹窗 "失败 N 条"

## 验证执行(必须做的命令)

```bash
# 1. 预览代码变更后,运行打包
cd "E:/Claude+CODEX Project/FS Capture/VBA Captor"
py tools/install_modules.py

# 2. 用户手动测试:
#    - 关 V3.xlsm 重开
#    - 样本池配置 Test 1 / 2 / 3 的公司清单
#    - 点『一键全抓』
#    - 检查 8 张表数据 + 美股_抓取诊断

# 3. 抽样校验(用 Python 对比 EDGAR raw):
py -c "
import json
with open('samples/AAPL_edgar.json') as f: d = json.load(f)
print('AAPL Revenue raw:', d['facts']['us-gaap']['RevenueFromContractWithCustomerExcludingAssessedTax']['units']['USD'][-3:])
"
```

## 已踩过的坑(继承 STATUS.md §B.6,新增本阶段相关)

1. **VBA Variant 数组从 3-tuple 升 7-tuple**: 通过 `MapEntry*` helper 封装,主流程零 UBound 判断(Codex 修订 #4)
2. **fuzzy 评分阈值不能太低**: AAPL 503 个 concept,关键词 "Cash" 会命中 30+ 个 — 必须靠负向关键词 + 单位过滤
3. **ifrs-full taxonomy 单位偶尔不是 USD**: Codex 修订 #5 — 必须严格单位过滤,非 USD → MISSING_NON_USD,不强转
4. **诊断 sheet 必须先 ClearContents**: 多次抓数追加会爆行数,每次跑前清空数据区(R3:J1000,10 列)
5. **Tim Hall JsonConverter 的 Dictionary `.Keys` 返回 Array 不是 Collection**: 遍历时用 `For Each k In dict.Keys`,不是 `For i = 1 To dict.Count`

## Future work(开源化时再做)

Codex 修订 #2 留口:本阶段保持 hardcode,但**未来开源时**应做:

- **隐藏 sheet 配置**: V3.xlsm 内嵌一张隐藏的 `_配置_美股字段映射` sheet,用户可以在不改 .bas 的情况下追加候选 concept
- **可选外部 JSON 导入**: 启动时如发现同目录 `mappings/us_overrides.json`,merge 到 hardcode 之上(用户先于 hardcode)
- **勾稽校验**: BS Assets ≈ Liabilities + Equity / IS GP ≈ Revenue - COGS / CF Cash 收支平衡 — 校验失败的 fuzzy 推荐降级
- **derive_from**: 缺 Total assets 时从 AssetsCurrent + AssetsNoncurrent 推导
- **SEC filing presentation linkbase 解析**: Python helper,不在 VBA

这些都属于 Phase 4b-14b 或开源准备阶段。

## 工作量估算(收敛后)

| 任务 | 责任人 | 时间 |
|---|---|---|
| **Layer 1**: 扩充 BS/IS hardcode 候选 + 雪球补字段 | Codex | 60 min |
| **Layer 2**: ifrs-full 兜底 + 单位严格检查 | Claude | 60 min |
| **Layer 3a**: MapEntry helpers 封装 | Claude | 30 min |
| **Layer 3b**: fuzzy match engine + 评分函数(只推荐) | Claude | 60 min |
| **Layer 3c**: 诊断 sheet writer + collector + template/install 同步 | Claude | 60 min |
| Test 1 回归 + Test 2 新公司 5 家测试 + 修 bug | Claude + Codex | 60-90 min |
| **合计** | 双人并行 | **5-6 小时** |

## 文档关系

- **本文件 `PHASE_4B14_PLAN.md` 是 4b-14 唯一 source of truth**(v3 修订 #c)
- `STATUS.md` §A-§M 是 Phase 1 → 4b-13 已完成阶段的实施记录,不动
- `STATUS.md` §N 是 Codex 早期的开源化路线草案,**已被本文件替代**;实施完成后,我会在 STATUS.md §N 头部加入"⚠️ 已废弃,实际方案见 PHASE_4B14_PLAN.md"标记,并新增 §O 作为 4b-14 收口记录

## 给 Codex 的回执(第二轮)

第一轮 6 条意见全部纳入 v2;第二轮 5 条必改 + 3 条小建议全部纳入 v3。修订摘要见 §"v2/v3 修订纪要"。

如果你已经在跑 Layer 1,可以并行启动;我即将开始 Layer 2/3。你的 conceptMap entries 只要遵守:
1. 不超过 6-tuple 形态(category, label, gaap_csv [, unit, scale, ifrs_csv])
2. **现有 concept 保持第一位**(v3 修订 #b)
3. 别名追加在 CSV 后

我的主流程改造就不会和你冲突。如果你想在 Layer 1 里同时加 ifrs-full CSV(6-tuple),也欢迎。

如还有进一步意见,直接追加到这个文件末尾,我开工前最后过一眼。

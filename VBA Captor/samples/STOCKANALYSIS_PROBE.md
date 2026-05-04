# Phase 4g Step 4: stockanalysis 港股 + 中概美股覆盖度调研

执行时间: 2026-05-04

本次只保存 6 个候选 URL 的 HTML sample,用于 Phase 4h 决策参考;未接入 VBA,未扩大抓取范围。HTML sample 已按 `samples/stockanalysis_*.html` 留在本地并加入 `.gitignore`,不提交 git。

## 表 1: URL 可达性与财报覆盖

| Ticker | 市场 | URL | HTTP | 本页是否含完整财报数据 | 观察 |
|---|---|---|---:|---|---|
| 00700 | 港股 | `https://stockanalysis.com/quote/hkg/00700/financials/` | 404 | 否 | 页面标题为 404,未出现财务表格字段 |
| 02519 | 港股 | `https://stockanalysis.com/quote/hkg/02519/financials/` | 404 | 否 | 页面标题为 404,未出现财务表格字段 |
| 09988 | 港股 | `https://stockanalysis.com/quote/hkg/09988/financials/` | 404 | 否 | 页面标题为 404,未出现财务表格字段 |
| BABA | 中概美股 | `https://stockanalysis.com/stocks/baba/financials/` | 200 | 部分 | 收入表字段完整;本 URL 是 income statement,BS/CF 在子链接 |
| JD | 中概美股 | `https://stockanalysis.com/stocks/jd/financials/` | 200 | 部分 | 收入表字段完整;本 URL 是 income statement,BS/CF 在子链接 |
| PDD | 中概美股 | `https://stockanalysis.com/stocks/pdd/financials/` | 200 | 部分 | 收入表字段完整;本 URL 是 income statement,BS/CF 在子链接 |

覆盖率: 3/6 HTTP 200,3/6 HTTP 404。港股候选路径覆盖为 0/3,中概美股候选路径覆盖为 3/3。

## 表 2: 港股 stockanalysis vs 雪球 HK 字段映射

当前 3 个港股候选 URL 都是 404,因此无法形成可执行的 stockanalysis HK 字段映射。现有雪球 HK API 字段仍是港股主路径。

| 标准指标 | 当前雪球 HK 字段 | stockanalysis 港股候选 | 结论 |
|---|---|---|---|
| Revenue | `tto` | 不可用 | 继续用雪球 |
| Total Assets | `ta` | 不可用 | 继续用雪球 |
| Total Liabilities | `tlia` | 不可用 | 继续用雪球 |
| Net Income | `ploashh` | 不可用 | 继续用雪球 |
| Cash from Operations | `nocf` | 不可用 | 继续用雪球 |

## 表 3: 中概美股 stockanalysis vs EDGAR + 雪球 fallback

| 维度 | stockanalysis 中概美股 | 当前 EDGAR + 雪球 fallback | 评价 |
|---|---|---|---|
| 覆盖公司 | BABA/JD/PDD 3/3 可达 | EDGAR 主路径覆盖美股;POM/HTT 等 20-F 特殊公司走雪球 fallback | stockanalysis 可作为中概美股备用来源候选 |
| 本轮 URL 内容 | income statement 可用,含 Revenue/Gross Profit/Operating Income/Net Income 等字段 | 当前系统已覆盖 BS/IS/CF + 指标生成 | stockanalysis 单个 `/financials/` URL 不够,还需 BS/CF 子页面 |
| 三张表覆盖 | 本页不含 BS/CF,但页面内存在 balance-sheet / cash-flow 子链接 | 已落地并通过回归 | 若切换,需新增每家公司多 URL 解析 |
| 币种/单位 | 页面口径需另做单位和货币确认 | 当前已有单位诊断 + Phase 4f RMB hook | 切换前需补单位审计 |
| 实施成本 | 中等:复用韩股 HTML 表格解析思路,但要新增 US stock path 和字段映射 | 已稳定 | 不建议 Phase 4g 切换 |

## 结论

Phase 4h 不建议直接切换到 stockanalysis:

- 港股候选路径 3/3 为 404,不能替代雪球 HK API。
- 中概美股 3/3 可达,但本轮 URL 只覆盖收入表;要替代 EDGAR/雪球 fallback,还要新增 balance-sheet 与 cash-flow 子页面抓样、单位/币种审计、字段映射和回归。
- 更合理的 Phase 4h 策略:仅在雪球 cookie 失效或 EDGAR/雪球 fallback 受阻时,把 stockanalysis 作为中概美股备用路径研究;港股暂不切。

实施成本估算:

- 中概美股 stockanalysis fallback:约 0.5-1 天,前提是补抓 BS/CF 样本并确认单位。
- 港股 stockanalysis fallback:当前候选路径不可用,需要重新找 URL 规则或放弃;不建议投入实现。

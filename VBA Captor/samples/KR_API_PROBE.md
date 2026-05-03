# KR API Probe (Phase 4d Step 1)

Date: 2026-05-03

## Summary

结论:建议韩股 Phase 4d 主数据源选 **stockanalysis.com**。

理由:
- `stockanalysis.com` 对 `KRX:005930` 三星财报页直接 HTTP 200,无登录、无 cookie、无 Cloudflare 拦截,HTML 里已有可解析表格。
- 字段名为英文,与现有美股/港股输出标签更接近,字段覆盖明显优于雪球 KR probe。
- 支持 KOSPI 与 KOSDAQ: `005930` 三星和 `013890` Zinus 的 KRX 页面都可访问。
- 支持年度和季度:年度 URL 直接访问,季度通过 `?p=quarterly` 访问,可覆盖后续 Q1/Q3 测试。

雪球 KR 本轮不建议作为主源:尝试 `/kr/`、`/kospi/`、`/kosdaq/`、`/korea/` 等路径均未返回有效 `data.list`;退到 `/us/` 虽 HTTP 200 但 `list_len=0`,不能用于正式抓数。

## 1A. Xueqiu KR Probe

测试 ticker:
- `005930` 三星电子(KOSPI)
- `013890` Zinus(KOSDAQ)

测试 endpoint kind:
- `balance`
- `income`
- `cash_flow`
- `indicator`

测试路径/代码格式:
- path: `kr`, `kospi`, `kosdaq`, `korea`, `hk`, `us`
- symbol: `005930`, `KR005930`, `005930.KS`, `KRX005930`, `A005930`;Zinus 同理

结果摘要:

| Ticker | Kind | 最佳响应 | HTTP | error_code | data.currency | data.list |
|---|---|---|---:|---:|---|---:|
| 005930 | balance | `/us/balance.json?symbol=A005930` | 200 | 0 | CNY | 0 |
| 005930 | income | `/us/income.json?symbol=A005930` | 200 | 0 | CNY | 0 |
| 005930 | cash_flow | `/us/cash_flow.json?symbol=A005930` | 200 | 0 | CNY | 0 |
| 005930 | indicator | `/us/indicator.json?symbol=A005930` | 200 | 0 | CNY | 0 |
| 013890 | balance | `/us/balance.json?symbol=A013890` | 200 | 0 | CNY | 0 |
| 013890 | income | `/us/income.json?symbol=A013890` | 200 | 0 | CNY | 0 |
| 013890 | cash_flow | `/us/cash_flow.json?symbol=A013890` | 200 | 0 | CNY | 0 |
| 013890 | indicator | `/us/indicator.json?symbol=A013890` | 200 | 0 | CNY | 0 |

说明:
- `/kr/`, `/kospi/`, `/kosdaq/`, `/korea/` 路径均返回 404。
- `/us/` 路径只是接口存在,但对韩股 symbol 返回空列表。
- 本轮未找到雪球韩股可用 finance endpoint。

Dump:
- `samples/xueqiu_KR_005930_balance.json`
- `samples/xueqiu_KR_005930_income.json`
- `samples/xueqiu_KR_005930_cash_flow.json`
- `samples/xueqiu_KR_005930_indicator.json`
- `samples/xueqiu_KR_013890_balance.json`
- `samples/xueqiu_KR_013890_income.json`
- `samples/xueqiu_KR_013890_cash_flow.json`
- `samples/xueqiu_KR_013890_indicator.json`

## 1B. stockanalysis.com Probe

测试 URL:
- `https://stockanalysis.com/quote/krx/005930/financials/`
- `https://stockanalysis.com/quote/krx/005930/financials/balance-sheet/`
- `https://stockanalysis.com/quote/krx/005930/financials/cash-flow-statement/`
- `https://stockanalysis.com/quote/krx/005930/financials/ratios/`
- KOSDAQ 覆盖补测: `https://stockanalysis.com/quote/krx/013890/financials/balance-sheet/`

结果摘要:

| URL kind | HTTP | Cloudflare | HTML table | 数据单位 | 备注 |
|---|---:|---|---:|---|---|
| 005930 income | 200 | No | 1 | millions KRW | 英文字段完整 |
| 005930 balance | 200 | No | 1 | millions KRW | 英文字段完整 |
| 005930 cash_flow | 200 | No | 1 | millions KRW | 英文字段完整 |
| 005930 ratios | 200 | No | 1 | N/A | 比率/估值指标 |
| 013890 balance | 200 | No | 1 | millions KRW | KOSDAQ 覆盖可用 |

年度字段样例:

Balance Sheet:
- `Cash & Equivalents`
- `Short-Term Investments`
- `Accounts Receivable`
- `Inventory`
- `Total Current Assets`
- `Property, Plant & Equipment`
- `Total Assets`
- `Accounts Payable`
- `Short-Term Debt`
- `Total Current Liabilities`
- `Long-Term Debt`
- `Total Liabilities`
- `Shareholders' Equity`
- `Total Liabilities & Equity`

Income Statement:
- `Revenue`
- `Cost of Revenue`
- `Gross Profit`
- `Selling, General & Admin`
- `Research & Development`
- `Operating Expenses`
- `Operating Income`
- `Interest Expense`
- `Pretax Income`
- `Net Income`
- `EPS (Basic)`
- `EPS (Diluted)`

Cash Flow:
- `Net Income`
- `Depreciation & Amortization`
- `Change in Accounts Receivable`
- `Change in Inventory`
- `Change in Accounts Payable`
- `Operating Cash Flow`
- `Capital Expenditures`
- `Investing Cash Flow`
- `Financing Cash Flow`
- `Net Cash Flow`
- `Free Cash Flow`

Ratios:
- `Current Ratio`
- `Quick Ratio`
- `Debt / Equity Ratio`
- `Asset Turnover`
- `Inventory Turnover`
- `Return on Equity (ROE)`
- `Return on Assets (ROA)`

季度可用性:
- `?p=quarterly` 可返回季度表。
- `balance-sheet/?p=quarterly` 示例列: `Q1 2026`, `Q4 2025`, `Q3 2025`, `Q2 2025`, `Q1 2025`, `Q4 2024`, `Q3 2024`。
- `financials/?p=quarterly` 与 `cash-flow-statement/?p=quarterly` 同样可用。

Dump:
- `samples/stockanalysis_KR_005930_income.html`
- `samples/stockanalysis_KR_005930_balance.html`
- `samples/stockanalysis_KR_005930_cash_flow.html`
- `samples/stockanalysis_KR_005930_indicator.html`
- `samples/stockanalysis_KR_005930_income_quarterly.html`
- `samples/stockanalysis_KR_005930_balance_quarterly.html`
- `samples/stockanalysis_KR_005930_cash_flow_quarterly.html`

## Scale Decision

stockanalysis HTML 表格显示单位为 `millions KRW`。三星 FY2024:

| Field | stockanalysis 表格值 | 实际含义 |
|---|---:|---:|
| Total Assets | 514,531,948 | 514,531,948 million KRW = 514.532 trillion KRW |
| Revenue | 300,870,903 | 300,870,903 million KRW = 300.871 trillion KRW |
| Net Income | 33,621,363 | 33,621,363 million KRW = 33.621 trillion KRW |

建议正式表单位采用 **十亿韩元(KRW billions)**:
- 如果解析 HTML 表格的 `millions KRW`,写表时除以 `1,000`。
- 如果解析 inline JS 的原始 KRW 数值,写表时除以 `1,000,000,000`。

理由:
- 百万韩元单位下,三星 Total Assets 会显示 `514,531,948`,可读性差。
- 兆韩元单位下,小公司如 Zinus 会太粗。
- 十亿韩元单位下,三星 Total Assets 显示 `514,531.948`,Zinus 也仍有足够精度。

## Recommendation

Phase 4d Step 2 推荐方案:

- 主源: `stockanalysis.com`
- 新主流程: `RunKRStatement`,独立于 US/HK。
- 获取方式: HTTP GET HTML,解析 `<table>`;年度页面直接访问,季度页面追加 `?p=quarterly`。
- 诊断状态建议: `OK_STOCKANALYSIS` / `MISSING`。
- 单位: `KRW billions`。
- KOSPI/KOSDAQ:统一 URL 形态 `/quote/krx/{ticker}/...`,不需要区分板块。

暂不建议:
- 暂不接雪球 KR,除非后续发现正式 endpoint。
- 暂不依赖 cookie。
- 暂不解析 ratios 页生成正式指标表;指标表继续沿用 18 个标准指标公式。

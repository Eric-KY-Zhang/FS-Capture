# Filings Atlas / 全球披露图谱

[English](#english) | [中文](#chinese)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/Eric-KY-Zhang/FS-Capture)](https://github.com/Eric-KY-Zhang/FS-Capture/releases/latest)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue.svg)]()

---

## English

Filings Atlas is a Windows desktop tool for one-click downloading of official disclosure PDFs across 7 markets. It focuses on the original filing files only: no financial statement extraction, no ratio calculation, and no Excel workbook generation.

![Filings Atlas English UI](docs/screenshots/main_window_en.png)

### Supported Markets

| Market | Source | Reports | Key required |
|---|---|---|---|
| A-Share | CNINFO + akshare | Annual, Q1, Interim, Q3, IPO prospectus | No |
| Hong Kong | HKEXnews + Eastmoney | Annual, audit report, IPO prospectus | No |
| United States | SEC EDGAR | Annual, quarterly, IPO prospectus | No |
| Korea | DART | Annual, Q1, Interim, Q3 | Optional DART API key |
| Taiwan | TWSE + MOPS | Annual, Q1, Interim, Q3, IPO prospectus | No |
| Japan | EDINET | Annual, Q1, Interim, Q3 | Strongly recommended EDINET Subscription-Key |
| United Kingdom | FCA NSM | Annual, interim and trading updates where available | No |

### Quick Start

1. Download `FilingsAtlas-v1.0.0-windows.zip` from [Releases](https://github.com/Eric-KY-Zhang/FS-Capture/releases/latest).
2. Extract the zip to any writable folder.
3. Double-click `Filings Atlas.exe`.
4. Select one or more markets, enter ticker codes, and click **Confirm** to resolve company names.
5. Choose years and report types, then click **Download Reports**.
6. PDFs are saved flat under `output/` with names like `UK_ULVR_Unilever PLC_2024_年报.pdf`.

### Ticker Examples

| Market | Examples |
|---|---|
| A-Share | `600519`, `000001` |
| Hong Kong | `00700`, `09988` |
| United States | `AAPL`, `BRK.B` |
| Korea | `005930`, `000660` |
| Taiwan | `2330`, `2317` |
| Japan | `7203`, `6758`, `9984.T` |
| United Kingdom | `ULVR`, `HSBA.L`, `AZN` |

### Optional Keys

Korea works without a key by using the public DART disclosure pages. Adding a free DART API key in **Settings** can make Korea faster and more stable.

Japan supports a public fallback, but v1.0 strongly recommends an EDINET Subscription-Key. The official EDINET API v2 uses `Subscription-Key` for document list and document download requests. Register for a free key through [EDINET](https://disclosure2.edinet-fsa.go.jp/) / [EDINET API key registration](https://api.edinet-fsa.go.jp/api/auth/index.aspx), then paste it in **Settings**. Without a key, Japanese downloads may fail when the API endpoint rejects unauthenticated requests.

United Kingdom uses the FCA National Storage Mechanism and does not require a key.

### Privacy And Scope

- Uses public disclosure endpoints only.
- Does not bypass login, scrape private data, or access unauthorized systems.
- Keeps API keys, cache, sidecars and downloaded files on your local machine.
- Does not provide investment advice or trading functionality.

### Developer Notes

- Source code lives under `development/`.
- Architecture and extension guide: [ARCHITECTURE.md](ARCHITECTURE.md).
- Tests: `cd development && python -m pytest -m "not e2e" -v`.

---

## Chinese

Filings Atlas / 全球披露图谱 是一个 Windows 桌面工具，用于一键下载 7 个市场上市公司的官方披露 PDF。工具只解决“批量拿到原始披露文件”这一件事：不抓三大报表数字、不算财务指标、不生成 Excel 底稿。

![全球披露图谱中文界面](docs/screenshots/main_window_zh.png)

### 支持市场

| 市场 | 数据源 | 报告类型 | 是否需要 Key |
|---|---|---|---|
| A 股 | 巨潮资讯 + akshare | 年报、一季报、半年报、三季报、IPO 招股书 | 不需要 |
| 港股 | 披露易 + 东方财富 | 年报、审计报告、IPO 招股书 | 不需要 |
| 美股 | SEC EDGAR | 年报、季报、IPO 招股书 | 不需要 |
| 韩股 | DART | 年报、一季报、半年报、三季报 | DART API Key 可选 |
| 台股 | TWSE + MOPS | 年报、一季报、半年报、三季报、IPO 公开说明书 | 不需要 |
| 日股 | EDINET | 年报、一季报、半年报、三季报 | 强烈推荐 EDINET Subscription-Key |
| 英股 | FCA NSM | 年报、半年报及可用的交易更新 | 不需要 |

### 快速开始

1. 从 [Releases](https://github.com/Eric-KY-Zhang/FS-Capture/releases/latest) 下载 `FilingsAtlas-v1.0.0-windows.zip`。
2. 解压到任意可写文件夹。
3. 双击 `Filings Atlas.exe`。
4. 勾选市场，输入股票代码，点击“确认”识别公司名称。
5. 选择年份和报告类型，点击“抓报告”。
6. PDF 会平铺保存到 `output/`，文件名示例：`UK_ULVR_Unilever PLC_2024_年报.pdf`。

### 股票代码示例

| 市场 | 示例 |
|---|---|
| A 股 | `600519`, `000001` |
| 港股 | `00700`, `09988` |
| 美股 | `AAPL`, `BRK.B` |
| 韩股 | `005930`, `000660` |
| 台股 | `2330`, `2317` |
| 日股 | `7203`, `6758`, `9984.T` |
| 英股 | `ULVR`, `HSBA.L`, `AZN` |

### Key 配置

韩股不填 Key 也可以使用 DART 公网披露页；如需更快、更稳，可在设置中填入免费的 DART API Key。

日股 v1.0 强烈推荐配置 EDINET Subscription-Key。EDINET API v2 的书类列表和文件下载接口都使用 `Subscription-Key`；可通过 [EDINET](https://disclosure2.edinet-fsa.go.jp/) / [EDINET API Key 注册页](https://api.edinet-fsa.go.jp/api/auth/index.aspx) 免费注册并申请 Key，然后在设置中填写。没有 Key 时，工具只会尝试 EDINET 公网页兜底，稳定性不能等同于官方 API。

英股使用 FCA National Storage Mechanism 公网数据源，不需要 Key。

### 隐私与范围

- 只访问公开披露接口。
- 不绕过登录，不抓取私人数据，不访问未授权系统。
- API Key、缓存、sidecar 元数据和下载文件都只保存在本机。
- 不做交易、不抓实时行情、不提供投资建议。

### 开发者文档

- 源码目录：`development/`
- 架构与“如何加新市场”：[`ARCHITECTURE.md`](ARCHITECTURE.md)
- 测试命令：`cd development && python -m pytest -m "not e2e" -v`

## License

[MIT License](LICENSE) © 2026 Eric Zhang

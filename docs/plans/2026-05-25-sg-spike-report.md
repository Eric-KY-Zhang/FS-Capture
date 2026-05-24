# SG SGXNet Spike Report — 2026-05-25

## 端点

**主配置端点**：`https://www.sgx.com/config/appconfig.json?v=04c0b410`

**实测可用数据端点**：

- Announcements：`https://api.sgx.com/announcements/v1.1/`
- Financial reports：`https://api.sgx.com/financialreports/v1.0`
- IPO prospectus：`https://api.sgx.com/ipoprospectus/v1.0/`
- PDF/HTML attachment host：`https://links.sgx.com/`

**重要差异**：sprint 草案里的 `https://api.sgx.com/announcements/v1.0/search` 当前返回 `403 {"message":"Missing Authentication Token"}`，不应按该路径实现。当前网页配置指向 `announcements/v1.1/`，并要求从 CMS `we_chat_qr_validator` 生成 `authorizationToken`。

**curl 复现**：

```bash
curl -H "User-Agent: Mozilla/5.0" \
  "https://www.sgx.com/config/appconfig.json?v=04c0b410"

curl -H "User-Agent: Mozilla/5.0" \
  -H "Referer: https://www.sgx.com/securities/company-announcements" \
  -H "authorizationToken: <ROT13(qrValidator)>" \
  "https://api.sgx.com/announcements/v1.1/securitycode?value=D05&periodstart=20240101_000000&periodend=20241231_235959&pagestart=0&pagesize=5"
```

`qrValidator` 获取方式：

1. 读 `config/appconfig.json` 的 `CMS_VERSION` 和 `endpoints.CMS_API_URL`。
2. GET `{CMS_API_URL}/?queryId={CMS_VERSION}:we_chat_qr_validator`。
3. 对返回的 `data.qrValidator` 做字母 ROT13，作为 `authorizationToken` header。

## 三类 period 实测

| Period | Ticker | API 字面值 | PDF 直链字段 | 实跑 PDF 大小 |
|---|---|---|---|---|
| ANNUAL | D05 | `financialreports.title = "Annual Report"` | API `url` 是公告 HTML，需二次解析 `.pdf` href | 7,356,816 bytes |
| ANNUAL | U11 | `financialreports.title = "Annual Report"` | 同上 | 12,475,199 bytes |
| ANNUAL | Z74 | `financialreports.title = "Annual Report"` | 同上 | 7,527,630 bytes |
| H1 | U11 | `category_name = "Financial Statements"` + title `Financial Statements and Related Announcement::Half Yearly Results` | API `url` 是公告 HTML，需二次解析 `.pdf` href | 1,043,730 bytes |
| IPO | LION-CM EM ASIA INDEX ETF | `ipoprospectus` rows，无 announcement_type；页面类别为 IPO | API `url` 是 IPO HTML，需二次解析 `FileOpen/*.ashx?App=IPO&FileID=...` | 720,845 bytes |

所有 5 票下载响应均为 `200 application/pdf`，文件头均为 `%PDF-`。

## Ticker 格式

实测规范化规则成立：

| 输入 | 规范化 |
|---|---|
| `D05` | `D05` |
| `d05` | `D05` |
| `D05.SI` | `D05` |
| `d05.si` | `D05` |

查询 SGX `securitycode` 端点时使用无 `.SI` 的代码。

## 反爬

- 必要 headers：`User-Agent`、`Referer`、`authorizationToken`。
- Token：由公开 config + CMS query 生成，脚本实测长度 140。
- Cookie/session：不需要。
- 连续请求：`announcements/v1.1/securitycode/count` 连续 10 次、间隔 0.5s，状态全为 `200`。
- 缺 token：`announcements/v1.1/securitycode` 返回 `401`。
- Playwright：不需要。HTTP API + HTML 附件解析可拿到 PDF；Playwright headless 访问页面会遇到 SGX 页面级 Access Denied，但不影响 HTTP 实现路径。

## 已知坑

- `links.sgx.com` 的 API `url` 不是 PDF 直链，而是公告/IPO HTML 页面。
- 年报 HTML 可能有多份 PDF，例如 Annual Report + Letter to Shareholders；应优先文件名含 `Annual Report` 的 PDF。
- H1 页面可能有多份 PDF，例如 financial statement、news release、presentation；应优先文件名含 `Interim Financial Statement` 或同义财务报表文件。
- `financialreports/v1.0` 的过滤参数能力有限，网页前端也是批量拉取后前端过滤；插件可按页拉取后本地筛选，并缓存。
- `documentDate` 对年报更像报告年度日期；announcement `submission_date` 是披露日期。选片时需要像 UK 一样区分报告年和披露年。

## 建议 plugin 实现

- 数据源：JSON API 为主，HTML 附件解析为必要二跳；不需要 SPA/Playwright fallback。
- 名称解析：`announcements/v1.1/securitycode` 取 `issuers[].stock_code/security_name/issuer_name`，可做 24h 缓存。
- 年报：用 `financialreports/v1.0` 拉取 `id,companyName,documentDate,securityName,title,url`，本地过滤 `title == "Annual Report"`。
- H1：用 `announcements/v1.1/securitycode`，过滤 `category_name == "Financial Statements"` 且 title 含 `Half Yearly Results` 或 `Second Quarter and/ or Half Yearly Results`。
- IPO：用 `ipoprospectus/v1.0/` 拉取 `closing_date,name,id,modified_date,url,status`，二跳解析 `FileOpen` PDF。
- 限流：`sgxnet = 1.5` QPS 起步；spike 的 2 QPS 短 burst 未触发 429/403。
- 缓存：financial reports 全表和 IPO 列表建议 24h；token 可按进程缓存，失败时刷新。

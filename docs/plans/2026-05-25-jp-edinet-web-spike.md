# JP EDINET 公网爬虫 Spike Report — 2026-05-25

## 结论

建议回 Planner 走 **Plan B（Playwright 搜索列表 + httpx 直链下载 PDF）**。

原因：当前 EDINET `WEEE0030.aspx` 不是 addendum 假设的 ASP.NET WebForms `__VIEWSTATE` / `__EVENTVALIDATION` 表单，而是 GeneXus 页面。搜索按钮触发 3 段加密 `GXAjaxRequest` JSON POST，URL 中含 `GX_AJAX_KEY` / `GX_AJAX_IV` 派生 token。纯 httpx 要稳定复现，需要移植 GeneXus `gx.sec.rijndael` 加密与事件 payload 组合，实施风险高。

PDF 下载本身不需要 Playwright。搜索结果返回 `SHORUI_KANRI_NO` 后，可直接 GET：

```text
https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/{SHORUI_KANRI_NO}.pdf
```

## 端点

- 搜索页: `https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx`
- PDF 直链: `https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/{doc_id}.pdf`

复现命令：

```powershell
$env:PYTHONIOENCODING="utf-8"
python tmp_pytest_ok/jp_web_spike.py
```

PDF 直链复现：

```bash
curl -L -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" \
  "https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/S100TR7I.pdf" \
  -o S100TR7I.pdf
```

## 真实表单字段

| 字段 | 用途 | 必填 | 实测 |
|---|---|---:|---|
| `GXState` | GeneXus 页面状态，内含 `GX_AJAX_KEY` / `GX_AJAX_IV` / `AJAX_SECURITY_TOKEN` / hash | 是 | 有 |
| `W0018vD_KEYWORD` | 关键词，接受证券代码，如 `7203` | 是 | 有 |
| `W0018vCHKSYORUI1` | 有価証券報告書 / 半期報告書 / 四半期報告書 | 是 | 默认勾选 |
| `W0018vCHKSYORUI2` | 大量保有報告書 | 否 | 有 |
| `W0018vCHKSYORUI3` | その他の書類種別 | 否 | 有 |
| `W0018vCHKSYORUI4` | 臨時報告書 | 否 | 有 |
| `W0018vD_KIKAN` | 提出期間，`7` = 全期間 | 是 | 有 |
| `vD_HIDDEN_PARAM` | PDF/XBRL/CSV 点击时放 encrypted doc handle | PDF click 才需要 | 有 |
| `vD_HIDDENPAGE` | 页码 | 翻页才需要 | 有 |
| `__VIEWSTATE` | ASP.NET viewstate | 否 | **不存在** |
| `__EVENTVALIDATION` | ASP.NET event validation | 否 | **不存在** |

`D_Keyword` 是页面 label，不是真实 input name；真实 input name 是 `W0018vD_KEYWORD`。

## 搜索机制

浏览器点击 `W0018BTNBTN_SEARCH` 后产生 3 次 JSON POST：

1. `cmpCtx="W0018"`, event `"'DOBTN_SEARCH'"`
2. `cmpCtx="W0018"`, event `T(O(14,'WEEE0030_EVENTS'),72).SETPARAMETER`
3. `cmpCtx=""`, event `T(O(14,'WEEE0030_EVENTS'),72).BACKSETPARAMETER`

第 3 次响应的 `gxValues[0].AV125W_RESULT_LIST_JSON` 是结果数组。结果行包含：

- `TEISHUTSU_NICHIJI`
- `SHORUI_NAME`
- `SYORUI_SB_CD_ID`
- `SHORUI_KANRI_NO`
- `EDINET_CD`
- `TEISYUTUSYA_NAME`
- `PDFKBN`

纯 httpx 直打可见 URL query 的测试结果：

```text
WEEE0030.aspx?base64("mul=7203&...&pfs=7&ser=1&pag=1&sor=2")
status=200, total=0, rows=0
```

## PdfClick 解析

页面 JS：

```text
Weee0030PdfClick(shoruiKanriNo) -> vD_HIDDEN_PARAM = shoruiKanriNo -> BTNBTN_PDF.click()
```

实测类型：GeneXus AJAX postback。

但实现不必走 PdfClick：结果行直接提供 `SHORUI_KANRI_NO`（如 `S100TR7I`），PDF 直链无需 cookie、无需 `PLD_0001.aspx`、无需 token：

```text
https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/S100TR7I.pdf
```

Playwright 点击 PDF 时，服务端会返回 `PDFDISP.Target = https://disclosure2.edinet-fsa.go.jp/PLD_0001.aspx`，该页面再嵌入同一 `disclosure2dl` PDF URL。

## 翻页机制

- 精确 ticker 搜索本次 3 票均为 1 页，结果足够实现年报选择。
- 宽泛关键词 `トヨタ` 实测显示 `1 / 3 ページ - 全208 件`。
- 点击 `次へ` 后页面变为 `2 / 3 ページ`，未产生新的网络 POST；当前前端分页是客户端基于已返回结果 JSON 翻页。

## Session cookie

GET 搜索页设置：

```text
ApplicationGatewayAffinityCORS
ApplicationGatewayAffinity
EDINET_SessionId
GX_SESSION_ID
GX_CLIENT_ID
ASLBSA
ASLBSACORS
```

搜索阶段需要同一浏览器 session 执行 GeneXus JS；PDF 直链阶段不需要这些 cookie。

## 限流验证

本次 spike 过程中多次搜索与 PDF 下载未出现 429 / 403。脚本正式输出覆盖：

- 3 个 ticker 搜索，每票 3 次 GeneXus AJAX POST，共 9 次 POST
- 3 个 PDF 直链 GET
- 间隔 1s 下载下一票

建议实现默认 `edinet_web = 1.0 QPS`；如走 Playwright 搜索，每 ticker 一次搜索即可，避免按 365 天枚举。

## 三票验证

| Ticker | 公司 | Period | 搜到 | PDF 下载 | 文件大小 |
|---|---|---|---:|---:|---:|
| 7203 | Toyota | 2024 年提交年报 `S100TR7I` | 是 | 是 | 3,392,218 bytes |
| 6758 | Sony | 2024 年提交年报 `S100TS7P` | 是 | 是 | 1,430,846 bytes |
| 9984 | SoftBank Group | 2024 年提交年报 `S100TP3N` | 是 | 是 | 1,254,446 bytes |

落地文件：

```text
output/jp_web_spike/JP_7203_Toyota_2024_annual_spike.pdf
output/jp_web_spike/JP_6758_Sony_2024_annual_spike.pdf
output/jp_web_spike/JP_9984_SoftBank Group_2024_annual_spike.pdf
```

脚本已校验三个文件头均为 `%PDF`。

## 实现建议

- 搜索列表：走 Plan B，用 Playwright 执行 WEEE0030 搜索并捕获 `AV125W_RESULT_LIST_JSON`。
- PDF 下载：仍走 `default_client(source="edinet_web")` + `stream_to_file`，URL 为 `disclosure2dl/.../{doc_id}.pdf`。
- `reports.py` 建议采用按 ticker + 全期間一次查的方案，跳过无 key 模式下 365 天枚举。
- `_SELECTORS` 不应继续使用旧 stub 里的假设；至少包括 `W0018vD_KEYWORD`、`W0018vD_KIKAN`、`W0018vCHKSYORUI1` 与结果 JSON 捕获逻辑。

## 失败回退

如果 Planner 仍要求 Plan A（纯 httpx），需要新增独立 spike 移植 GeneXus `gx.sec.rijndael` 加密并复现 `GXAjaxRequest` URL token。该路线不是普通 BeautifulSoup 表单 POST，预计风险高于 Playwright 路线。

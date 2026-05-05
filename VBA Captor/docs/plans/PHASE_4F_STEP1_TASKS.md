# PHASE_4F_STEP1_TASKS.md

> Planner sub-agent 产出, Orchestrator 落盘, 供 Generator sub-agent 执行
> 路径相对 repo 根 `E:\Claude+CODEX Project\FS Capture\VBA Captor\`

## 全局执行约束(对 Generator 适用于所有子任务)

- **抓取统一用 Python `requests`**, 不用 PowerShell `Invoke-WebRequest`(`xueqiu_quote_currency.json` 和 `xueqiu_kline_USDCNY.json` 之前因 PS TLS 握手失败留下了 `"基础连接已经关闭: 接收时发生错误"` 残留)
- **每个 fetch 间隔 ≥ 1.0 s**(雪球反爬)
- **Headers 模拟 VBA 现有 `XueqiuHttpGet`**(参见 `modules/模块_工具函数.bas:608-633`):
  ```python
  HEADERS = {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
      "Accept": "application/json, text/plain, */*",
      "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
      "Accept-Encoding": "identity",
      "Referer": "https://xueqiu.com/",
      "Origin": "https://xueqiu.com",
  }
  ```
- **Cookie 来源**: 用户已登录的 xueqiu cookie 存在 `上市公司财务数据查询.xlsm` → `样本池!B5`。Generator 改读环境变量 `XUEQIU_COOKIE`(若未设置则不带 cookie 试一次, 失败后报告兜底)
- **所有 sample 文件用 UTF-8 写入**, 不用 PowerShell `Out-File`。Python `open(path, "w", encoding="utf-8")`
- **覆盖既有残留**: `samples/xueqiu_quote_currency.json` 和 `samples/xueqiu_kline_USDCNY.json` 当前是失败 dump, **必须覆盖重写**
- **执行顺序**: 1A → 1B → 1C 可并行, 1D 必须最后

---

## 子任务 1A — 雪球 quote 实时汇率 probe

### 1A.1 动作

```python
import requests, json, time, os, datetime

URL = "https://stock.xueqiu.com/v5/stock/realtime/quotec.json?symbol=USDCNY,HKDCNY,KRWCNY"
HEADERS = {  # 见全局约束
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "identity",
    "Referer": "https://xueqiu.com/",
    "Origin": "https://xueqiu.com",
}
cookie = os.environ.get("XUEQIU_COOKIE", "")

def fetch(use_cookie: bool):
    h = dict(HEADERS)
    if use_cookie and cookie:
        h["Cookie"] = cookie
    r = requests.get(URL, headers=h, timeout=30)
    return r.status_code, r.text

# 先无 cookie 试 (plan 假设 quote 公开)
status_nc, body_nc = fetch(False)
time.sleep(1.2)
# 有 cookie 再试一次作对照
status_c, body_c = fetch(True) if cookie else (None, None)

out = {
    "url": URL,
    "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
    "no_cookie": {"status": status_nc, "body_excerpt": body_nc[:600] if body_nc else None,
                  "parsed": json.loads(body_nc) if body_nc and body_nc.startswith("{") else None},
    "with_cookie": {"status": status_c, "body_excerpt": body_c[:600] if body_c else None,
                    "parsed": json.loads(body_c) if body_c and body_c.startswith("{") else None},
}
with open("samples/xueqiu_quote_currency.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
```

### 1A.2 预期产出

- `samples/xueqiu_quote_currency.json`(覆盖既有失败残留)
- 解析后 `parsed.data.items` 应是数组, 每项含 `symbol` ∈ {USDCNY, HKDCNY, KRWCNY}, 字段 `current` / `last_close` / `chg` / `percent` / `timestamp`

### 1A.3 验收 criteria

1. 文件存在、UTF-8 无 BOM、能被 `json.load` 解析
2. **至少 `no_cookie` OR `with_cookie` 一组** `status == 200` 且 `parsed.error_code == 0`
3. `parsed.data.items` 长度 == 3, 涵盖 `USDCNY` `HKDCNY` `KRWCNY`(`symbol` 字段)
4. 每个 item 都含 `current` 数值字段(浮点 > 0)

### 1A.4 风险 + 兜底

| 风险 | 兜底 |
|---|---|
| TLS 握手失败 | 已用 Python `requests`; 仍失败换 `httpx[http2]` 或 `verify=False` |
| 无 cookie 返回 401/403 | with_cookie 重试; 都失败则报告"必需 cookie" |
| symbol 大小写或后缀错误 | 若 `data.items` 为空, 补测 `symbol=USD/CNY,HKD/CNY,KRW/CNY` 和 `symbol=USDCNY.FX,...`, 全部 dump 到 `alt_attempts` |

### 1A.5 依赖

无(可与 1B、1C 并行)

---

## 子任务 1B — 雪球 K 线历史汇率 probe

### 1B.1 动作

```python
import requests, json, time, os, datetime

BASE = "https://stock.xueqiu.com/v5/stock/chart/kline.json"
HEADERS = { ... }  # 同 1A
cookie = os.environ.get("XUEQIU_COOKIE", "")

def fetch_kline(symbol, begin_ms, count, use_cookie=True):
    params = {
        "symbol": symbol,
        "begin": begin_ms,
        "period": "day",
        "type": "before",
        "count": count,  # 负数 = 向前取 N 根
    }
    h = dict(HEADERS)
    if use_cookie and cookie:
        h["Cookie"] = cookie
    r = requests.get(BASE, headers=h, params=params, timeout=30)
    return r.url, r.status_code, r.text

attempts = []

# 1) USDCNY 365 天 sample (cookie + no-cookie 各试)
for use_c in (False, True) if cookie else (False,):
    url, st, body = fetch_kline("USDCNY", 1704038400000, -365, use_cookie=use_c)
    parsed = json.loads(body) if body and body.startswith("{") else None
    attempts.append({
        "label": f"USDCNY_365d_cookie={use_c}",
        "url": url, "status": st,
        "parsed_meta": {
            "error_code": parsed.get("error_code") if parsed else None,
            "error_description": parsed.get("error_description") if parsed else None,
            "column_names": parsed.get("data", {}).get("column") if parsed else None,
            "item_count": len(parsed.get("data", {}).get("item", [])) if parsed else 0,
            "first_item": parsed.get("data", {}).get("item", [None])[0] if parsed else None,
            "last_item": parsed.get("data", {}).get("item", [None])[-1] if parsed else None,
        },
    })
    time.sleep(1.2)

# 2) USDCNY 拉到 2020-01-01 测试历史范围: begin=2020-01-01 UTC=1577836800000, count=-1500 (~6 yr)
url, st, body = fetch_kline("USDCNY", 1577836800000, -1500, use_cookie=bool(cookie))
parsed_long = json.loads(body) if body and body.startswith("{") else None
items_long = parsed_long.get("data", {}).get("item", []) if parsed_long else []
attempts.append({
    "label": "USDCNY_long_2020_to_now",
    "url": url, "status": st,
    "earliest_ts_ms": items_long[0][0] if items_long else None,
    "latest_ts_ms": items_long[-1][0] if items_long else None,
    "earliest_date_utc": datetime.datetime.utcfromtimestamp(items_long[0][0]/1000).isoformat() if items_long else None,
    "earliest_date_cst": datetime.datetime.utcfromtimestamp(items_long[0][0]/1000 + 8*3600).isoformat() if items_long else None,
    "item_count": len(items_long),
})
time.sleep(1.2)

# 3) HKDCNY + KRWCNY 各拉 365 天
for sym in ("HKDCNY", "KRWCNY"):
    url, st, body = fetch_kline(sym, 1704038400000, -365, use_cookie=bool(cookie))
    p = json.loads(body) if body and body.startswith("{") else None
    attempts.append({
        "label": f"{sym}_365d",
        "url": url, "status": st,
        "error_code": p.get("error_code") if p else None,
        "item_count": len(p.get("data", {}).get("item", [])) if p else 0,
        "first_item": p.get("data", {}).get("item", [None])[0] if p else None,
    })
    time.sleep(1.2)

out = {
    "primary_sample": {
        "symbol": "USDCNY", "period": "day", "count": -365,
        "begin_ms": 1704038400000,
        "begin_meaning_utc": "2024-01-01T00:00:00Z",
        "begin_meaning_cst": "2024-01-01T08:00:00+08:00",
    },
    "attempts": attempts,
    "full_USDCNY_365d_parsed": parsed,  # 含 column + item 完整数组
    "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
}
with open("samples/xueqiu_kline_USDCNY.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
```

### 1B.2 预期产出

- `samples/xueqiu_kline_USDCNY.json`(覆盖既有失败残留)
- `column` 数组通常为 `["timestamp", "volume", "open", "high", "low", "close", ...]` — 实测以输出为准
- 关键字段: `item[i][0]` = timestamp ms, `item[i][5]` = close

### 1B.3 验收 criteria

1. 文件可被 `json.load` 解析(UTF-8 无 BOM)
2. `attempts` 中至少一条 `status == 200` 且 `error_code == 0` 且 `item_count >= 240`
3. `full_USDCNY_365d_parsed.data.column` 含 `"timestamp"` 和 `"close"`
4. `attempts` 含 `USDCNY_long_2020_to_now`, `earliest_date_cst <= "2020-02-01"` — 证明能拉到 2020 起 6 年
5. `attempts` 含 `HKDCNY_365d` 和 `KRWCNY_365d`, 均 `error_code == 0` 且 `item_count > 200`
6. 1D 报告必须明确 begin 时区(UTC ms / CST ms): 对比 `attempts[0].first_item[0]` 转 UTC 日 与 CST 日

### 1B.4 风险 + 兜底

| 风险 | 兜底 |
|---|---|
| K 线需 cookie | attempts cookie + no-cookie 都试 |
| 拉不到 2020 历史 | 1D 写明 `earliest_date`, 推荐 fallback 数据源 |
| 时区误判 | 实测对照 begin 与 first_item 日期 |
| count 上限 | 实测上限, Step 3 拉 6 年要分段 |
| 反爬限速 | 已 sleep 1.2s; 429 时提到 2.0s |

### 1B.5 依赖

无

---

## 子任务 1C — 傲基股份代码 probe

### 重要前置发现

`samples/xueqiu_HK_02519_balance.json` **已存在并已经返回真数据**: `quote_name=傲基控股 / annual_settle_date=12-31 / currency=CNY`。**强烈暗示傲基股份 = 02519.HK**。1C 任务是"补齐 confirm + 清空 A 股竞争候选"。

### 1C.1 动作

```python
import requests, json, time, os, datetime, re

HEADERS = { ... }  # 同 1A
cookie = os.environ.get("XUEQIU_COOKIE", "")

candidates = [
    ("HK_02519_xueqiu_quote", "hk_quote", "02519.HK",
     "https://stock.xueqiu.com/v5/stock/quote.json?symbol=02519&extend=detail"),
    ("HK_02519_xueqiu_finance_balance", "hk_fin", "02519",
     "https://stock.xueqiu.com/v5/stock/finance/hk/balance.json?symbol=02519&type=all&is_detail=true&count=8"),
    ("HK_09927_xueqiu_quote", "hk_quote", "09927",
     "https://stock.xueqiu.com/v5/stock/quote.json?symbol=09927&extend=detail"),
    ("A_003031_xueqiu_quote", "a_quote", "SZ003031",
     "https://stock.xueqiu.com/v5/stock/quote.json?symbol=SZ003031&extend=detail"),
    ("A_003031_sina_balance", "a_sina", "sz003031",
     "https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/003031/ctrl/part/displaytype/4.phtml"),
]

results = []
for label, kind, sym, url in candidates:
    h = dict(HEADERS)
    if "xueqiu" in label.lower() and cookie:
        h["Cookie"] = cookie
    if "sina" in label.lower():
        h = {"User-Agent": HEADERS["User-Agent"],
             "Accept-Language": "zh-CN,zh;q=0.9",
             "Accept-Encoding": "identity"}
    try:
        r = requests.get(url, headers=h, timeout=30)
        if "sina" in label.lower():
            r.encoding = "gb18030"  # sina 财经页是 GB18030
        body = r.text
        st = r.status_code
    except Exception as e:
        body = f"EXCEPTION: {e}"
        st = None
    
    summary = {
        "label": label, "url": url, "status": st,
        "body_excerpt": body[:800] if isinstance(body, str) else None,
    }
    if "xueqiu" in label.lower() and body and body.startswith("{"):
        try:
            p = json.loads(body)
            summary["error_code"] = p.get("error_code")
            if "quote" in label.lower():
                d = p.get("data", {}).get("quote", {})
                summary["company_name"] = d.get("name")
                summary["currency"] = d.get("currency")
                summary["exchange"] = d.get("exchange")
                summary["type"] = d.get("type")
                summary["status"] = d.get("status")
                summary["isin"] = d.get("isin")
            elif "balance" in label.lower():
                d = p.get("data", {})
                summary["quote_name"] = d.get("quote_name")
                summary["currency"] = d.get("currency")
                summary["last_report_name"] = d.get("last_report_name")
                summary["list_len"] = len(d.get("list", []))
        except Exception as e:
            summary["parse_error"] = str(e)
    if "sina" in label.lower() and body:
        summary["html_contains_aoji"] = "傲基" in body or "Aoji" in body
        summary["html_contains_unknown"] = "未知" in body or "无效" in body
        m = re.search(r"<title>(.*?)</title>", body, re.IGNORECASE | re.DOTALL)
        summary["title"] = m.group(1).strip() if m else None
    
    fname_safe = label.replace("/", "_").replace(".", "_")
    ext = "html" if "sina" in label.lower() else "json"
    raw_path = f"samples/aoji_probe_{fname_safe}.{ext}"
    with open(raw_path, "w", encoding="utf-8") as f:
        f.write(body if isinstance(body, str) else "")
    summary["dump"] = raw_path
    
    results.append(summary)
    time.sleep(1.2)

with open("samples/aoji_probe_summary.json", "w", encoding="utf-8") as f:
    json.dump({
        "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "candidates": results,
        "preexisting_evidence": {
            "samples/xueqiu_HK_02519_balance.json": "Phase 4c HK probe 残留, "
                "header 已显示 quote_name=傲基控股, currency=CNY, "
                "annual_settle_date=12-31 → 强证据 02519.HK = 傲基股份"
        },
    }, f, ensure_ascii=False, indent=2)
```

### 1C.2 预期产出

- `samples/aoji_probe_summary.json` — 各候选归一化结果表
- `samples/aoji_probe_HK_02519_xueqiu_quote.json`
- `samples/aoji_probe_HK_02519_xueqiu_finance_balance.json`
- `samples/aoji_probe_HK_09927_xueqiu_quote.json`
- `samples/aoji_probe_A_003031_xueqiu_quote.json`
- `samples/aoji_probe_A_003031_sina_balance.html`

### 1C.3 验收 criteria

1. `aoji_probe_summary.json` 存在, 每个 candidate 都有结果
2. **至少一个港股候选明确命中"傲基"**:
   - `HK_02519_xueqiu_quote` 应返回 `company_name` 含"傲基", `exchange` ∈ {HKEX, SEHK, HK}
   - **OR** `HK_02519_xueqiu_finance_balance` 应返回 `quote_name` 含"傲基" 且 `list_len > 0`
3. A 股 `003031` 候选要明确 reject(不含傲基)
4. 1D 必须给出明确结论: **"傲基股份 = 02519.HK 傲基控股, 港股, 报告币种 = CNY"**

### 1C.4 风险 + 兜底

| 风险 | 兜底 |
|---|---|
| 02519 quote 端点需 cookie | 用 `xueqiu_HK_02519_balance.json` 已有证据兜底 |
| 09927 是占位 | 报告里列出"09927 = X 公司, 排除"|
| 003031 sina GBK 编码 | `r.encoding = "gb18030"` 显式解码 |
| 02519 在数据源未收录 | 优先信 xueqiu finance/hk 已有 8 报告期数据 |
| 用户口语 vs quote_name 全称 | 同时记 quote_name + plan 名"傲基股份"|

### 1C.5 依赖

无

---

## 子任务 1D — 写探查报告 `samples/RMB_FX_PROBE.md`

### 1D.1 动作

`1A`、`1B`、`1C` 全部完成后, 汇总成单一 Markdown 报告。**不要用模板自动生成**, 根据实测结果写**结论性表述**(参考 `samples/HK_API_PROBE.md` 和 `samples/KR_API_PROBE.md` 的风格)。

报告必含 6 节: TL;DR / 1A / 1B / 1C / 1D 期间平均算法 / 1E 给 Step 2/3 输入。

详细 schema 见 plan §1D.1。

### 1D.2 预期产出

- `samples/RMB_FX_PROBE.md` — 单一报告文件, UTF-8

### 1D.3 验收 criteria

1. 文件存在, UTF-8, markdown 头注 3 行齐全
2. **TL;DR 3 行**(quote 通否 / K线通否 / 傲基代码)给出明确"是/否"或具体值, 不允许 `<TBD>`
3. 1A 表格至少 USDCNY/HKDCNY/KRWCNY 三行 current 列填实数
4. 1B 章节明确写明: column index 表 + earliest 日期 + 时区结论(UTC 还是 CST)
5. 1C 章节"结论"那行不能含 `?`, 必须形如"傲基股份 = 02519.HK 傲基控股, 港股 SEHK, 报告币种 CNY"
6. 1D 章节有期间平均算法 pseudo + 至少一条"为什么不用月度加权"的取舍说明
7. 1E 章节列出 3 个"给 Step 2/3 的可执行输入"
8. 风险章节至少 3 条, 每条对应 Step 2/3 某个具体可调参数

### 1D.5 依赖

**1A + 1B + 1C 全部完成后才能开始**

---

## 给 Evaluator 的关注重点摘要(按风险高到低)

1. **TLS/编码污染再现**: `xueqiu_quote_currency.json` 之前 fail 过。第一件事跑 `python -c "import json; json.load(open('samples/xueqiu_quote_currency.json', encoding='utf-8'))"` 验通
2. **K 线时区决策**: Step 3 拉每报告期汇率全靠这个。1D 报告必须实测对照
3. **傲基代码确认强度**: 虽强证据已有, 1C 必须再 round-trip 拿到 `company_name` 字段
4. **K 线历史范围**: 若 < 2020 则 deal-breaker
5. **Cookie 依赖路径**: 1A/1B 双跑明确"是否 cookie 必需"
6. **003031 sina GB18030 编码**: 必须 `r.encoding = "gb18030"`

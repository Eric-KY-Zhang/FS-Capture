# SPRINT v0.6.1 — Patch 修复

**日期**：2026-05-23
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex（待委派）
**Reviewer**：Claude Code
**状态**：待 Codex 实施
**预计工作量**：1-2 天

---

## Context

v0.6（2026-05-17 发布）上线后，对 5 市场代码做了一轮系统审计（3 个并行 subagent + 交叉验证），找到 7 个确认存在的缺陷：3 个高危（用户可感知影响）、3 个中危（影响数据正确性但隐蔽）、1 个低危（潜在未来 bug）。本 sprint 专注 bug 修复 + 体验微调，**不引入新功能、不动架构**，目标是让 v0.6 用户在不学习新行为的前提下获得更高的数据正确性和 UX 反馈质量。

**不在范围内**：性能优化、KR 公网爬虫、新市场、Playwright 移除、ruff 全量清债——这些都进 v0.7/v0.8。

---

## 改动清单（按文件分组）

### 改动 #1 [🔴 高] `app/core/ratelimit.py` — TokenBucket 速率热更新

**现状**（`ratelimit.py:58-62`）：
```python
def get(self, source: str, rate: float) -> TokenBucket:
    with self._lock:
        if source not in self._buckets:
            self._buckets[source] = TokenBucket(rate)
        return self._buckets[source]   # ← 忽略新 rate
```

**问题**：用户在 Settings 中改限流后必须重启应用才生效。

**修复要求**：
- bucket 已存在时检测 `bucket.rate != rate`，差异时更新 `self.rate` 和 `self.capacity`；同时调用 `_refill()` 后 reset `_last` 时间戳避免速率跳变
- 更新动作需在 `self._lock` 内进行，保证线程安全
- 不修改公网 API 签名（`limiter(source, rate)` 调用方零改动）

**单元测试要求**（新增 `tests/test_ratelimit.py`）：
- `test_token_bucket_rate_hot_update`：创建 bucket(rate=2)，再 `get(source, rate=10)`，断言 `bucket.rate == 10` 且 `capacity == 10`
- `test_token_bucket_rate_unchanged_returns_same_instance`：相同 rate 两次 get 返回 `is` 同一对象

---

### 改动 #2 [🔴 高] `plugins/hk/reports.py` — PDF 验证失败应 drop 候选

**现状**（`hk/reports.py:411-438`）：
```python
scored: list[tuple[int, dict]] = [
    (_base_report_score(row, ticker, period), row) for row in candidates
]
if len(candidates) > 1:
    top = sorted(scored, ...)[:5]
    verified_ids: set[int] = set()
    for _score, row in top:
        if verify_pdf_year_and_kind(row.get("url") or "", period.year, verify_kind):
            verified_ids.add(id(row))
    scored = [(score + (10 if id(row) in verified_ids else 0), row) for score, row in scored]
```

**问题**：验证只用于 +10 加分，验证失败的候选仍可能因 base score 高被选中，导致用户拿到"标题含年份但内容是其他年份"的错文件（v0.6 已知风险，文档 `PROJECT_RETROSPECTIVE.md:170-176` 提及但未修）。

**修复策略**：
- 对进入 top 5 的候选**显式做 PDF 验证**，验证通过的进 `verified_ids`
- **新规则**：如果 `verified_ids` 非空（至少 1 个候选通过 PDF 验证），则**仅在 verified_ids 中**选最高分；如果全部失败，则降级回旧策略（按 base score 选最高），并 `logger.warning` 提示用户该结果未经 PDF 内容验证
- 加 `logger.info` 输出每个候选验证结果（带 url + 期望 year/kind + 实际识别），便于排查
- `verify_pdf_year_and_kind` 30s 超时不变；如果是网络问题导致验证全失败，降级路径保证不影响功能

**单元测试要求**（扩展 `tests/test_hk_selection.py`）：
- `test_pdf_verification_drops_wrong_year_candidate`：mock 2 个候选 PDF，verify 一个返回 True 一个 False，断言选中的是 True 那个，即使 False 那个 base score 更高
- `test_pdf_verification_all_fail_falls_back_with_warning`：mock 全部返回 False，断言仍返回 base score 最高的，并 `caplog` 捕获到 warning

---

### 改动 #3 [🔴 高] `plugins/hk/fiscal_year.py` — 补 8-10 家非 12 月财年港股

**现状**（`hk/fiscal_year.py:6-15`）：只覆盖 6 家（阿里、联想、旺旺、周大福、新世界、信和置业）。

**问题**：HK 选片在选 PDF 时按财年窗口判断 period.year 对应的披露日期范围，非 12 月财年公司若未在表中，默认按 12 月窗口选 → 选错。

**修复要求**：

**必加（验证后写入）**：
| code | 公司 | 财年终止月 | 来源 |
|---|---|---|---|
| 00005 | 汇丰控股 | 12（默认即可，**不用加**） | — |
| 01299 | 友邦保险 | 12（默认即可，**不用加**） | — |
| 00019 | 太古股份A | 12（默认即可，**不用加**） | — |
| 01038 | 长江基建 | 12（默认即可，**不用加**） | — |
| 00006 | 电能实业 | 12（默认即可，**不用加**） | — |

**❗ Codex 必须做的事**：上面这些热门股都是 12 月财年，**真正需要补的是非 12 月财年的**。请按以下名单逐个**实地查证**（HKEX 年报披露日 + 公司 IR 页），把确认非 12 月财年的写入：

| 候选 code | 公司 | 怀疑财年终止月 |
|---|---|---|
| 00001 | 长江和记实业 | 12（不用加） |
| 00002 | 中电控股 | 12（不用加） |
| 00003 | 香港中华煤气 | 12（不用加） |
| 00388 | 香港交易所 | 12（不用加） |
| 00688 | 中国海外发展 | 12 |
| 02888 | 渣打集团 | 12 |
| 06823 | 香港电讯 | 12 |
| 00066 | 港铁公司 | 12 |
| 01113 | 长实集团 | 12 |
| 02888 | 渣打 | 12 |
| 03988 | 中国银行 | 12 |
| 00939 | 建设银行 | 12 |
| 00011 | 恒生银行 | 12 |
| 00012 | 恒基地产 | 6（地产系常见 6 月） |
| 00016 | 新鸿基地产 | 6 |
| 00101 | 恒隆地产 | 12 |
| 00688 | 中国海外发展 | 12 |
| 00823 | 领展房产基金 | 3 |
| 00688 | 中国海外发展 | 12 |
| 01972 | 太古地产 | 12 |
| 00027 | 银河娱乐 | 12 |
| 01928 | 金沙中国 | 12 |
| 06030 | 中信证券 | 12 |
| 03968 | 招商银行 | 12 |
| 00386 | 中国石化 | 12 |
| 00857 | 中国石油 | 12 |
| 00883 | 中海油 | 12 |
| 01024 | 快手 | 12 |
| 09618 | 京东集团 | 12 |
| 09999 | 网易 | 12 |
| 03690 | 美团 | 12 |
| 09888 | 百度集团 | 12 |
| 00700 | 腾讯 | 12 |

**实际目标**：在港股大盘里筛出**真正非 12 月财年**的常见公司 8-10 家。已知的常见 3 月/6 月财年类别：

- **大陆赴港中概互联**：阿里（3）、联想（3）已在表
- **HK 房地产**：新鸿基地产 00016（6）、新世界发展 00017（6 已在表）、信和置业 00083（6 已在表）、恒基地产 00012（6）、九龙仓置业 01997（12）、九龙仓集团 00004（12）
- **REIT**：领展 00823（3）
- **零售消费**：周大福 01929（3 已在表）、旺旺 00151（3 已在表）

**最终交付清单**（Codex 实地查证后选 8-10 家纳入，下面是 Planner 建议名单）：

```python
NON_DEC_FISCAL_YEAR: dict[str, int] = {
    # === Existing ===
    "09988": 3,  # Alibaba
    "00992": 3,  # Lenovo
    "00151": 3,  # Want Want China
    "01929": 3,  # Chow Tai Fook Jewellery
    "00017": 6,  # New World Development
    "00083": 6,  # Sino Land
    # === New (verify each via HKEX/IR pages before commit) ===
    "00016": 6,  # Sun Hung Kai Properties — 港股最大地产之一
    "00012": 6,  # Henderson Land Development
    "00823": 3,  # Link REIT — REIT 普遍 3 月
    "01113": 12, # ❗ Codex 查证：长实集团；若是 12 不要加
    "00688": 12, # ❗ Codex 查证：中国海外发展；若是 12 不要加
    # 如有其他 3/6 月财年的 港股蓝筹 / 中概，按需加
}
```

**Codex 验证流程**（每只必做）：
1. 打开 HKEX 公司披露页 `https://www.hkexnews.hk/index_c.htm`，搜公司 code → 找最近一份"年度报告"
2. 看年报标题或封面"截至 X 年 X 月 X 日止年度"
3. 若不是 12 月，记录终止月写入表

**单元测试要求**（扩展 `tests/test_hk_selection.py` 或新增 `tests/test_hk_fiscal_year.py`）：
- `test_fiscal_year_lookup_returns_default_dec`：未知 code 返回 12
- `test_fiscal_year_lookup_known_march_year_end`：阿里 09988 返回 3
- `test_fiscal_year_lookup_known_june_year_end`：新鸿基 00016 返回 6（前提是确实被纳入）

---

### 改动 #4 [🟡 中] `plugins/kr/name_resolver.py:96` — 列名 typo

**现状**：
```python
industry = str(row.get("induty_code", "")) or None
```

**问题**：`induty_code` 看起来像 typo（应该是 `industry_code`），但 OpenDartReader 的真实字段名可能两者都不是。

**Codex 必做调试**（**不能直接改名**）：
1. 临时在 `fetch_company` 加 `logger.debug(f"DART company columns: {list(info_df.columns)}")`
2. 用一只韩股（如 005930 三星）实跑一次，从日志看真实列名
3. 根据真实列名改 `row.get("...")`；如果列名就是 `induty_code`（DART 的英文拼写习惯），则**保留原代码并加注释说明"非 typo，DART 字段名"**
4. 调试日志清掉再提交

**单元测试要求**：
- 在 `tests/integration/test_plugins.py` 现有 KR 用例上加断言：`company.industry is not None`（用 monkeypatch 的 fake DataFrame 测试）

---

### 改动 #5 [🟡 中] `plugins/ashare/name_resolver.py:62,132` — `zip(strict=True)`

**现状**：两处 `zip(df["code"], df["name"], strict=False)`。

**问题**：akshare 返回脏数据时（行长度不一致）静默丢数据，没有诊断信号。

**修复要求**：
- 改为 `strict=True`
- 用 `try-except ValueError as exc:` 包裹，catch 后 `logger.error(f"akshare 返回的 code/name 列长度不一致：{exc}")` 并回退到 `strict=False` 路径
- 这样**功能不退化**（极端情况下仍可用），但有了诊断信号

**测试要求**：
- 新增 `tests/test_ashare_name_resolver.py`（如已存在则扩展）
- `test_zip_strict_logs_warning_on_mismatch`：mock 一个 code/name 长度不等的 DataFrame，断言 caplog 中出现 error

---

### 改动 #6 [🟡 中] `app/ui/batch_import_dialog.py` — 失败行反馈

**现状**（`batch_import_dialog.py:50-60`）：
```python
def parse_ticker_codes(text: str, exchange: Exchange) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for token in _candidate_tokens(text, exchange):
        code = _normalize_token(token, exchange)
        if not code or code in seen:
            continue
        seen.add(code)
        out.append(code)
    return out
```

**问题**：失败 token 静默跳过，用户不知道哪一行被忽略。

**修复策略**：

**步骤 1**：重构 `parse_ticker_codes` 签名为：
```python
def parse_ticker_codes(text: str, exchange: Exchange) -> tuple[list[str], list[str]]:
    """Return (valid_codes, rejected_raw_tokens)."""
```

**步骤 2**：`BatchImportDialog.codes()` 返回 `(codes, rejected)` 元组（或新增 `rejected()` 方法）

**步骤 3**：UI 改动 — 在批量导入对话框 OK 后：
- 若有 rejected，用 `QMessageBox.information` 显示"以下 N 行未识别（可能格式错误或与所选市场不匹配）"+ 列出 rejected raw tokens（max 20 行，超过省略）
- 若 rejected 为空，行为不变

**步骤 4**：去重也要反馈 — 重复的 token 不算 reject（不显示），但可以在 status bar 提示"去重了 N 行"。

**测试要求**（扩展 `tests/test_batch_import.py`）：
- `test_parse_returns_rejected_tokens`：输入 `"600519\nINVALID\n000001"` + A_SHARE，断言 `rejected == ["INVALID"]`
- `test_parse_us_stopwords_not_in_rejected`：输入 `"AAPL INC"` + US，断言 INC 不在 rejected（属于已知 stopword 过滤）

---

### 改动 #7 [🟢 低] `app/core/settings.py:117` — 返回值意图修正

**现状**：
```python
if not cfg_file.exists():
    s = Settings()
    save_settings(s)
    return Settings.model_validate({})   # 应该返回 s
```

**问题**：当前等价但语义不一致；如未来 `Settings` 加入无默认值字段，会变成真 bug。

**修复**：
```python
if not cfg_file.exists():
    s = Settings()
    save_settings(s)
    return s
```

**测试**：现有 `tests/test_settings.py` 应已覆盖；如未覆盖，加 `test_load_settings_creates_default_on_missing`。

---

## 实施顺序（Codex 分批 commit）

| 批次 | 内容 | 验证命令 |
|---|---|---|
| **1** | 改动 #1 + #7（core 层小修） | `pytest tests/test_ratelimit.py tests/test_settings.py -v` |
| **2** | 改动 #4 + #5（KR/A 股 plugin 调试与小改） | `pytest tests/integration/test_plugins.py tests/test_ashare_name_resolver.py -v` |
| **3** | 改动 #3（HK 财年表扩充，需实地查证） | `pytest tests/test_hk_selection.py -v`（含新增 fiscal_year 测试） |
| **4** | 改动 #2（HK PDF 验证 drop 逻辑，核心改动） | `pytest tests/test_hk_selection.py -v` + 实跑 09988 阿里 ANNUAL 2024 验证不再选错 |
| **5** | 改动 #6（UI batch import 反馈） | `pytest tests/test_batch_import.py -v` + 手动 EXE 启动跑一遍 |
| **6** | 全量回归 + 打包 | `pytest -m "not e2e" -v` + `python -m build` 重新生成 EXE，跑一次 A/HK/US smoke |

每批结束后 Codex 提 commit，commit message 格式：`v0.6.1: <一句话变更>（改动 #N）`。

---

## 测试矩阵

### 单元测试（必绿）

```bash
cd development
pytest -m "not e2e" -v
```

预期：v0.6 已有 30 passed + 本期新增/修改 ~10 个用例 → 全绿。

### 回归 smoke（实跑验证）

每批末尾 Codex 自检；批次 6 必跑：

| 市场 | 票 | 期间 | 验证点 |
|---|---|---|---|
| A | 600519 贵州茅台 | ANNUAL 2024 | PDF 落地，文件名扁平契约 |
| HK | 09988 阿里巴巴 | ANNUAL FY2024（3 月财年） | **关键**：PDF 验证 drop 逻辑生效 + 财年表 3 月正确取窗口 |
| HK | 00700 腾讯 | ANNUAL 2024 | 默认 12 月财年仍可用 |
| US | AAPL | ANNUAL FY2024 | 10-K 正确选中 |
| TW | 2330 台积电 | ANNUAL 2024 | 现有 v0.6 路径无回归 |

KR 不在 v0.6.1 范围（要等 v0.7 公网爬虫完工才能无 Key 实跑）。

### 限流热更新手动验证（改动 #1 收尾确认）

1. 启动 EXE，开始一个抓取 job
2. 打开 Settings → 把 cninfo 限流从 5 改成 1
3. 保存后**不重启**，再跑一个 A 股 job
4. 观察日志中请求间隔是否变成 ~1 秒/请求（而非 0.2 秒/请求）

---

## Reviewer Checklist（Claude Code 验收用）

### 改动正确性
- [ ] #1 `RateLimiterRegistry.get` 已存在 bucket 时能更新 `rate` 和 `capacity`
- [ ] #2 HK PDF 验证有候选通过时**仅在 verified 中选**；全失败时降级 + warning
- [ ] #3 `NON_DEC_FISCAL_YEAR` 至少新增 5 个**实地查证后确认非 12 月**的港股；commit message 或 PR 描述附查证截图/链接
- [ ] #4 KR `induty_code` 字段名问题已通过日志调试确认真实列名；commit 中清掉了调试日志
- [ ] #5 `ashare/name_resolver.py:62,132` 用 `strict=True` + try/except 回退 + error log
- [ ] #6 `parse_ticker_codes` 签名变为 `tuple[list[str], list[str]]`；批量导入对话框新增 rejected 提示
- [ ] #7 `settings.py:117` `return s` 已修正

### 质量门禁
- [ ] `pytest -m "not e2e" -v` 全绿
- [ ] `pytest -m "not e2e" --co -q` 显示新增用例 ≥ 6 个
- [ ] 手动跑过 5 只 smoke 票，PDF 全部落地且文件名符合扁平契约 `{exchange}_{code}_{name}_{year}_{kind_zh}.pdf`
- [ ] 限流热更新手动验证通过

### 不应变动
- [ ] `app/core/orchestrator.py`、`app/core/http.py`、`app/core/pdf_renderer.py`、`app/core/cache.py` **零改动**（这些留给 v0.7/v0.8）
- [ ] 任何 plugin 的 PDF 下载链路逻辑 **零改动**
- [ ] `pyproject.toml`、`requirements.txt` **零新增依赖**
- [ ] 输出文件命名 schema 未变

### 文档
- [ ] `PROJECT_RETROSPECTIVE.md` 末尾新增 `## §10 v0.6.1 postscript` 章节，记录本次修了哪些 bug（每条 1-2 行）
- [ ] `README.md` 顶部版本号 `v0.6` → `v0.6.1`，"What's new" 段补一句

---

## 风险

- **#2 PDF 验证 drop 后选片更严格**：可能让某些边缘案例（标题正确但 PDF 内容识别不到 year 关键字）选不到结果。缓解：保留降级路径（全失败时回到旧逻辑 + warning），用户可在 UI 看到 warning 决定是否手动核实
- **#3 财年表查证耗时**：Codex 需要实地验证 5-10 家公司的财年，单家约 1-2 分钟，预算半小时
- **#6 UI 反馈对话框测试需 GUI 环境**：Codex 在 headless CI 下跑不到 UI 测试；可用 `pytest-qt` 或仅单测 `parse_ticker_codes` 部分

---

## 后续衔接

v0.6.1 验收通过后：
- 用户决定是否发布 patch tag（建议发，时间允许的话）
- Planner 起草 `SPRINT_v0.7_kr_no_key_and_testing.md`，主体来自 `docs/plans/2026-05-23-kr-no-api-key.md`，叠加测试补强 + 架构清债（httpx verify、plugin 重试统一、UI 字符串集中）

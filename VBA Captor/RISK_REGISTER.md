# 风险登记表 — 上市公司财务数据查询 v1.0

> **作者**: Eric Zhang(214978902@qq.com)
> **更新日期**: 2026-05-05
> **使用说明**: 部署 / 维护 / 用户反馈触发的风险持续登记。**等级**: 高(影响数据准确性 / 用户无法用)/ 中(影响体验 / 偶发失败)/ 低(corner case / 长期债)

## 数据源 / 抓数风险

| 风险 | 等级 | 影响 | 当前控制 | 后续动作 |
|---|---|---|---|---|
| 雪球 token 失效 | 高 | A 股(暂未用)/ 港股 / 美股中概 fallback 抓取失败 | 诊断行 `cookie_expired` 提示 + B5/E5 cell 提示 + Phase 4l retry/backoff 重试 | 改为外置凭证(env var / .secrets.json),Phase 4o+ 候选 |
| 数据源页面 / API 结构变化 | 高 | 字段抓取错误,跨市场指标表显示空 / #N/A | Phase 4d/4h 各市场字段映射 + 诊断 sheet 命中字段记录 + Phase 4m 8 张离线 fixture 测试覆盖解析 | 增加 fixture 真实样本(每季度刷新一次)+ 数据源监控脚本(每周跑一次抓样对比) |
| 汇率缺失(用户没跑过 EnsureFxRateCached)| 高 | RMB 报表错误(韩股可能差 ~200x) | Phase 4k 修:`GetFxRateStatus` 返回 `FX_MISSING` → 写空值 + 诊断行 `FX_MISSING` | **禁止 fallback=1**(已实施);考虑加 一键全抓 前 pre-flight 检查所有需要的汇率是否在 sheet |
| stockanalysis 中概美股 fallback 站点反爬 | 中 | BABA/JD/PDD 在雪球 cookie 失效 + 站点变化时全部失败 | Phase 4h Phase 4l retry/backoff;BABA/JD/PDD 白名单卡死,其他 ticker 不尝试 | Phase 4o+:扩展白名单到 BIDU/TCOM/NIO,基于实际使用统计 |
| SEC EDGAR Fair Access 触发(>10 req/sec)| 中 | EDGAR 抓数 429 / IP ban | Phase 4l SEC 独立限流 ≥110ms 间隔(实测 469ms);retry/backoff 接 429/5xx | 监控诊断 sheet 的 RetryCount 列,如果稳定 > 0 说明 SEC 端紧 |
| 港股 / 美股报告币种识别错误(`HKD` vs `CNY` 大陆港股)| 中 | RMB 换算用错币种 | Phase 4f 报告币种 dict 写入诊断;雪球 finance API 返回的 `currency` 字段为准 | Phase 4o+:加更多公司 sample 验证币种识别 |
| 韩股财报 stockanalysis HTML 表格结构变化 | 中 | 韩股字段抓取失败 | Phase 4d HTML 表格 parse + 诊断 sheet `taxonomy` 列;Phase 4m fixture 测试覆盖 | 季度刷新 fixture |
| 单位识别错误(USD millions / KRW billions / HKD millions)| 中 | 数值差 1000 倍 | Phase 4d/4h `Unit` 列写诊断;按市场 hardcode 单位规则 | 加 unit smoke test 到 fixture |

## 用户环境 / 部署风险

| 风险 | 等级 | 影响 | 当前控制 | 后续动作 |
|---|---|---|---|---|
| 宏安全限制(Excel 默认 block 互联网下载的宏)| 中 | 用户无法运行 | 发布说明:第一次打开点"启用宏";README 加 trust 设置说明 | 长期:数字签名(需要 code signing 证书)/ 受信任位置 |
| 用户 Excel 版本 < 2016 | 中 | UDF / Shape API / outline grouping 可能失效 | 不主动支持,文档声明最低 Excel 2016 | 加版本检测 startup hook,< 2016 弹错误 |
| Windows PowerShell ExecutionPolicy 严格 | 中 | 汇率抓取 PowerShell shell-out 失败 | Phase 4f shell-out 用 `-ExecutionPolicy Bypass` | 企业部署改 WinHttp / XMLHTTP 直接抓数(P0-04 defer)|
| 个人 cookie 通过 xlsm 分享泄露 | 高 | 雪球账号被滥用 | Phase 4l `CleanReleaseWorkbook` 宏(用户分享前主动调)+ release 版默认空 cookie | 长期:外置 secrets 管理(`.secrets.json` / env var)|
| 用户误操作清空样本池数据 | 低 | 自己输入的公司列表丢失 | 数据区 R14+ 不被任何脚本主动清空(Phase 4j 严禁项)| 长期:加 backup 宏(每次 一键全抓 前自动备份样本池到 archive)|

## 工具内部 / 代码质量风险

| 风险 | 等级 | 影响 | 当前控制 | 后续动作 |
|---|---|---|---|---|
| 缓存过期误用(用户拿到几小时前的旧数据)| 中 | 数据不是最新 | Phase 4l `CacheStatus` 列写诊断 sheet;Phase 4m source-aware TTL(SEC 168h / 雪球 12h / EDGAR 24h)| 加显眼警告:跨市场指标表 A1 注释加"最新抓数:YYYY-MM-DD HH:MM" |
| Excel COM 跨进程 stale state | 中 | inspect / smoke test 偶发 0x800ac472 | 文档化"先 kill stale Excel + Python 再跑回归" | release 包含 helper script: `tools/kill_stale_com.py` |
| 合并单元格触发 VBA 模态错误 | 中 | 清理 / 测试路径偶发 hang | Phase 4j.1 抑制 BuildCrossMarketIndicatorSheet 的合并弹窗;Phase 4n release_v1_clean 用 UnMerge + 写值 + Re-Merge | 长期审计所有 Range().Value 调用,识别合并区域 |
| 模块_工具函数.bas 3000+ 行 god module | 低 | 维护性差 | ARCHITECTURE.md 标记拆分计划(P2-01) | Phase 4o+:按 Phase 4n plan 模板拆 9 个 sub-module |
| OERN 84 处分布全工程 | 低 | 关键错误可能静默消失 | Phase 4l 关键路径已用 ErrorStage/ErrorText 写诊断 | 顺手做(不专 sprint)|
| AppStateGuard 仅覆盖 5 入口 | 低 | toggle tabs / CleanReleaseWorkbook 等 sub 出错可能脏 Excel 状态 | 这些 sub 逻辑轻量,不需要 AppStateGuard | 用户反馈触发再扩 |

## 文档 / 维护风险

| 风险 | 等级 | 影响 | 当前控制 | 后续动作 |
|---|---|---|---|---|
| 用户不读使用说明 sheet 直接乱用 | 低 | 误操作 / 报错时找不到诊断 | 样本池 O1:Q5 内联使用提示 + 错误时 StatusBar 写明 | 加新手引导(第一次打开弹欢迎 dialog)|
| 字段映射 doc 跟代码不同步 | 低 | 维护时改了代码忘改 README | ARCHITECTURE.md 数据流 section + 各市场 fetch 模块顶部注释 | code review checklist 加"doc 同步" |
| STATUS.md 越来越长不可读 | 低 | 找历史决策费劲 | 100K+ 行已经接近 GitHub 显示极限 | Phase 4o+:STATUS.md 拆成 STATUS_PHASE_4*.md 按版本归档 |

---

## 风险等级总览

- **高风险 4 项**:雪球 token 失效 / 数据源结构变化 / 汇率缺失 / cookie 泄露 — 都已有控制 + 后续动作
- **中风险 9 项**:多数有缓解,部分待 Phase 4o+ 执行
- **低风险 7 项**:技术债 / corner case,不影响核心功能

## 监控建议

每周跑一次:
```bash
py tools/test_fx_live.py --skip-install     # 验证 FX 抓数链路
py tools/run_offline_tests.py                # 验证字段映射稳定
```

每月跑一次:
- 真实抓数 4 公司样本(300866 / AAPL / 00700 / 005930),目视检查跨市场指标表数值合理
- 检查诊断 sheet `RetryCount` / `FX_MISSING` / `CacheStatus` 列异常分布

每季度:
- 刷新 fixture(`tests/fixtures/sec_aapl_companyfacts.json` 等)
- 验证字段映射依然命中

— Eric Zhang(214978902@qq.com),2026-05-05

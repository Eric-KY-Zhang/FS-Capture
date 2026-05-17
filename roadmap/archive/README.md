# roadmap/archive/ — 历史规划文档

本目录保存 v0.1 → v0.5 的 Sprint 规划与协作记录，归档于 v0.6 发布后（2026-05-17）。

## 文件清单

| 文件 | 内容 | 日期 |
|---|---|---|
| `ROADMAP.md` | v0.2 重定位时的 5-Sprint 总规划（v0.1 → v0.5） | 2026-05-09 |
| `SPRINT_01_drop_financial_extraction.md` | Sprint 01 详细方案：删除抓数据 / Excel 模板代码 | 2026-05-09 |
| `CODEX_BUNDLE_PROMPT.md` | 给 Codex 端到端跑 Sprint 02-05 的 887 行 bundle prompt | 2026-05-09 |
| `current_v0.2_planner_log.md` | Planner（Claude）↔ Generator（Codex）↔ Reviewer（Claude）轮次日志 | 2026-05-09 → 2026-05-10 |

## 用途

- **回溯设计决策**：理解为什么 v0.1 的财务数据抓取代码被砍掉，为什么选 plugin-per-exchange 架构等
- **复用规划模板**：未来再做大版本重构时可参考三人协作工作流
- **审计完整性**：每个 Sprint 都有具体验收标准、commit 哈希、Reviewer 签字

## 不再适用

- v0.6 起新增市场（如台股）不再走 Sprint 流程，直接 commit
- `current.md` 已不再维护；项目状态以 `README.md` + `ARCHITECTURE.md` + `PROJECT_RETROSPECTIVE.md` 为准

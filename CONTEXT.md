# Filings Atlas / 全球披露图谱 — Glossary

> 产品语言的 source of truth。**不**写实现细节、不写架构决策、不写 sprint 计划。
> 实现 / 架构 → `ARCHITECTURE.md`；sprint → `roadmap/`；架构决策记录 → `docs/adr/`。

---

## Filings Atlas / 全球披露图谱

英文 + 中文双语产品名，**不可拆分**为单独使用。两个名字都是正式名称：
- 英文场景：`Filings Atlas`
- 中文场景：`全球披露图谱`
- 仓库目录仍是 `FS Capture/`（v0.9 重命名 sprint 决定不动目录名）

**避免**写成：`FS Capture`（v0.8 之前的旧名）/ `Atlas` 单用 / `图谱` 单用。

---

## 披露 PDF（Disclosure PDF）

工具的**唯一**抓取对象。指上市公司在官方监管渠道（A 股 cninfo / 港股 HKEXnews / 美股 SEC / 韩股 DART / 台股 MOPS / 日股 EDINET / 英股 NSM / 新加坡 SGXNet）提交的年报、季报、半年报、IPO 招股说明书的 PDF 文件。

**不是**：财务数据 / 报表数字 / Excel 底稿 / 财务指标 / 估值。这些在 v0.2 重定位时被砍掉，不能"顺手做"。

---

## 云端 Atlas（v1.0 视觉语言）

UI refresh 的设计哲学，由 grill 会话（2026-05-24）确定为**混合风**：

- **主体仍是工作台**：信息密度 + 操作效率优先，"专业感"是硬约束，不让步给"地图感"
- **Atlas 识别作为点缀**，集中在 3 处：
  1. 应用 logo：云端 + 图钉 + 等高线轮廓的单一图形
  2. TitleBar 品牌区：低存在感的等高线背景纹饰（淡到几乎不打扰操作）
  3. ExchangeChip：用 route dot 元素暗示市场位置（**不画**真实地图）

**明确不做**：
- 全屏地图等高线作为背景
- 航线连接市场之间的视觉
- 任何让 UI 看上去像旅游 / 航空 / 物流 app 的元素
- 抛弃 v0.9 已经验证可用的"工作台密度"

依据：v0.9 polish 已经验证"工作台"产品定位可用（用户的不满是"没有标志性"，**不是**"太工具感"）。地图感作为点缀风险最低。

---

## 八个市场（Eight Markets）

v1.0 起 Filings Atlas 覆盖 8 个上市公司披露源（按 plugin 实施顺序）：

| Market | Code | 数据源 | Period |
|---|---|---|---|
| A 股 | `A_SHARE` | cninfo + akshare | ANNUAL / Q1 / Q2 / Q3 |
| 港股 | `HK` | HKEXnews | ANNUAL |
| 美股 | `US` | SEC EDGAR | ANNUAL / Q1 / Q2 / Q3 / IPO |
| 韩股 | `KR` | DART API + 公网 fallback | ANNUAL / Q1 / Q2 / Q3 |
| 台股 | `TW` | TWSE + MOPS | ANNUAL / Q1 / Q2 / Q3 / IPO |
| 日股 | `JP` | EDINET API + 公网 fallback | ANNUAL / Q1 / Q2 / Q3 |
| 英股 | `UK` | FCA NSM | ANNUAL / interim |
| 新加坡 | `SG` | SGXNet | ANNUAL / interim (H1) / IPO |

**红线**：plugin ABC 三方法签名固定（`resolve_name` / `fetch_company` / `download_reports`），新增市场不改契约。

---

## MarketPin（v1.0 替换 ExchangeChip）

主界面顶部 8 个并排的市场选择按钮，每个对应一个 Market。v1.0 UI refresh 由 grill（2026-05-24）把 v0.9 的 `ExchangeChip` 形态完全替换为 `MarketPin`。

**形态**：map pin 风格（呼应"云端 Atlas"产品语言），不再是矩形 chip。
**布局**：横向 8 pin，地理从西到东 + 文化中心调整（**不**做真二维地图）
**排序**（西 → 东）：

```
[ UK ][ US ] ‖ [ A 股 ][ HK ][ TW ][ KR ][ JP ][ SG ]
└─ 欧美冷色块 ─┘  └────── 亚洲暖色块 ──────────────────┘
```

**色系约束**（A3 sprint 决议）：
- **欧美 2 个走冷蓝系**：UK / US 彼此可区分（如 navy 深蓝 vs cyan 青蓝）
- **亚洲 6 个走暖色系**：必须彼此色相差足够（不能 4 个全橙）
- **A 股保留 rose-600 `#E11D48`** —— 中国股市"涨红"文化语义不可改
- 其余 5 暖色具体 hex 在 Figma 设计阶段由 Codex 提，Reviewer 审

**业务行为不变**（v0.9 → v1.0）：点击切换可见性、触发 ExchangePanel 渲染、active state 视觉强化。

**`palette.py::_EXCHANGE_ACCENT`** 仍是公共 API，但内部 8 个 hex 全部重新设计（与 v0.9 不兼容）。`ExchangePanel` / `ProgressDock` 等用 accent color 的下游同步更新。

⚠️ `ExchangeChip` 这个**词**在 v1.0 后不再使用。代码里如有 `ExchangeChip` 命名应统一改 `MarketPin`。

---

## TickerRow

ExchangePanel 内单行 ticker 输入组件。一个 ExchangePanel 可含多个 TickerRow。

每行含：ticker code 输入框 + 名称解析按钮 + 解析状态徽标（OK / 失败 / Pending）+ 删除按钮。

**红线**：用户输入的 ticker code 不能在语言切换 / 重渲染时丢失（v0.9 i18n sprint 因这个红线选了 in-place re-translate 模式，不重建 widget）。

---

## 设计 token（5 个 objectName）

v0.9 末 UI polish (commit `4832479`) 建立的设计基础设施。**v1.0 UI refresh 可以扩展但不能删**：

| Token | 用途 |
|---|---|
| `#HeroTitle` | 主标题 34px / 800（"全球披露图谱"）|
| `#HeroSubtitle` | 副标题 14px / muted（slogan）|
| `#SectionLabel` | section 上方小标签 11px / uppercase / primary 色 |
| `#TitleBarLogo` | 标题栏品牌名 |
| `#LangSegment` | 中/EN 切换段控制器（带 active / position 属性）|

---

**最后更新**：2026-05-24（grill 会话 + v1.0 UI refresh 启动）

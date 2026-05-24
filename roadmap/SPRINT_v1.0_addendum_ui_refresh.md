# SPRINT v1.0 ADDENDUM 2 — UI Refresh + Icon Redesign

**日期**：2026-05-24
**Planner**：Claude Code (Opus 4.7)
**Worker**：Codex
**Reviewer**：Claude Code
**状态**：待 Codex 启动
**预计工作量**：4-5 天（4 批次 + 1 Reviewer Checkpoint）
**发布策略**：v1.0 首发 GitHub release 推迟到 UI refresh 完成后；v1.0.0 tag 完工后重打到新 HEAD

---

## Context

v1.0 release readiness 状态：
- 主 sprint（SG + 性能 + 体积）✅
- Addendum 1（JP 公网真双模式）✅
- 8 市场 e2e sweep（81 MB PDF）✅
- 打包修复（pandas.plotting 锁定）✅
- 当前 EXE：`Filings Atlas.exe` 2026/5/24 16:46，含全部修复

**用户决策**（2026-05-24）：v1.0 首发 release 视觉必须完整，UI refresh 算 v1.0 增量（不走 v1.1）。release 推迟 3-5 天换"开箱视觉到位"。

### 用户核心诉求

- 当前 UI 虽然 v0.9 已做过一轮 polish（HeroTitle + 主界面语言切换 + SectionLabel 设计 token）但仍偏"工具感"，缺少"Atlas/全球披露图谱"产品识别度
- 希望"云端 Atlas"风格：清爽浅色、低噪声、地图等高线/航线感、云端归档感
- 明确**避免**：旧 FS 标识残留（v0.8 之前的视觉）/ 金融终端黑蓝风 / 股票图表风
- 图标也要重做（不再用 v0.9 时做的简单文字 logo）

### v0.9 时已经做过的 UI polish（**Codex 必须先读懂当前现状**）

新一轮 refresh 不能"重新发明轮子"，先认清现状：

| Commit | 已做 |
|---|---|
| `a801a13` v0.9 批次 1 | 品牌重命名 `FS Capture → Filings Atlas`、`filings_atlas.ico` + `filings_atlas_logo.png` 资源就位 |
| `f7bccbd` v0.9 批次 4 | 双语 strings + i18n 脚手架 + 11 widget `_retranslate` 接入 |
| `4832479` v0.9 末 UI polish | HeroTitle 34px/800、HeroSubtitle、SectionLabel、ExchangeChip 14px border-radius、TitleBar LogoIcon + Divider、主界面 LangSegment 中/EN 切换、palette 锁主色 `#6366F1` indigo-500 |

**v0.9 polish 的 5 个 objectName 设计 token 是基础设施，可以保留也可以替换，但不能漏掉**：
`#HeroTitle` / `#HeroSubtitle` / `#SectionLabel` / `#TitleBarLogo` / `#LangSegment`

### release 与 tag 状态

- `v1.0.0` 本地 tag 在 `6fb8c3d`（addendum 1 完工 commit）
- 打包修复 `0725701` 已 commit 但 tag 未移
- UI refresh 完成后由 Reviewer 评审 → 用户授权 → tag 重打到 UI refresh 完工 commit → push origin

---

## 已锁定决策

| 决策项 | 选定方案 |
|---|---|
| Sprint 归属 | **v1.0 增量**（release 推迟到 UI refresh 完成）|
| 设计前置 | **Figma 设计稿前置**（4 页 audit/system/workbench/icon），Codex 不能跳过 |
| Figma 连接器 | **执行前必须验证 `create_new_file` / `use_figma` 可用**；若不可用立刻 STOP 回 Planner，**不允许 Canva / 本地 PNG 冒充 Figma 源稿** |
| 图标 | **重做**（GPT Image 2 生成 3 个云端地图针概念 → 本地重绘 → 多尺寸 ico）|
| 资源文件命名 | **保持** `filings_atlas.ico` / `filings_atlas_logo.png`（不动 spec extra_datas 拷贝路径）|
| 设计风格 | 云端 Atlas：浅色清爽 + 地图等高线/航线感 + 低噪声密度 |
| 颜色基调 | 保留 `#6366F1` indigo-500 主色或微调；新增云灰/浅蓝中性色构建 Atlas 识别 |
| 设计基准 | 当前 1120x806 + 最小 960x680 双基准 |
| Reviewer Checkpoint | 1 个（批次 3 后，UI 落地完成）|

---

## 批次概览

| 批次 | 名称 | 工作量 | 输出 |
|---|---|---|---|
| 1 | Figma 设计稿 + Current Audit | 1-1.5 天 | Figma 文件 4 页（含当前 UI 审计截图 + design system token + workbench + icon system）|
| 2 | 图标重做 | 0.5-1 天 | GPT Image 2 概念 + 本地重绘 + `filings_atlas.ico` (16/32/48/64/128/256) + `filings_atlas_logo.png` |
| 3 | PySide UI 重构 + QSS/palette | 1.5-2 天 | `main_view.py` / `main_window.py` 等结构调整 + `app.qss` 更新 + `palette.py` 扩展 |
| **🔴** | **Reviewer Checkpoint** | 0.5 天 | Claude 验收 |
| 4 | 截图更新 + 资产测试 + 文档 | 0.5 天 | `docs/screenshots/*.png` 重拍 + `test_assets.py` + CHANGELOG/README 微调 |

总计 **4-5 天**。

---

## 批次 1 — Figma 设计稿 + Current Audit

### 1.1 前置门控（必须）

**第一步**：在 Codex 启动批次 1 前，验证 Figma 连接器可用：
```
是否能调用 mcp__figma__create_new_file / use_figma 工具？
```
- 可用 → 继续
- 不可用 → **立刻 STOP，回 Planner 协调用户重新连接 Figma**
- **不允许 fallback 到 Canva、Mermaid、Excalidraw 或本地 PNG 拼接**冒充 Figma 源稿

### 1.2 Figma 文件结构

新建 Figma 文件 `Filings Atlas UI Refresh`，4 页：

**Page 1 — Current Audit**
- 截图当前 main_view（1120x806）+ main_window（含 titlebar）的中英文版本
- 标注当前已实现的 objectName 设计 token 位置（`#HeroTitle` / `#SectionLabel` / `#LangSegment` 等）
- 标注当前 8 个 ExchangeChip 的 accent 颜色映射（v0.9 palette）
- 标注"工具感重、Atlas 识别弱"的具体痛点
- 这一页是 baseline，不做设计，只做审计

**Page 2 — Design System**
- 颜色 token：保留主色 `#6366F1` 或微调；新增云灰 / 浅蓝中性色 token；深色主题对应色
- Typography：HeroTitle / SectionLabel / SubtitleMuted / ChipName / ChipMeta 等层级
- 形状 token：圆角（chip 14px / button 10px / card 12px）/ 阴影
- 地图视觉元素 token：等高线纹理 / route dot / market tile 视觉
- 状态徽标：StatusPill (ok/error/pending) 三态

**Page 3 — Main Workbench**
- 桌面主稿按 1120x806（默认）+ 960x680（最小）双基准设计
- 顶部 Atlas 品牌区（titlebar + brand icon + 中/EN toggle）
- 市场图谱选择区（重新设计 ExchangeChip，体现 Atlas/地图感）
- 当前市场录入区（TickerRow + 状态徽标）
- 报告/输出配置区（PeriodSelector + OutputCard）
- 固定底部操作栏（设置 + 增量更新 + 抓报告）
- 不能丢失现有任何业务入口

**Page 4 — Icon System**
- 3 个云端地图针概念草稿（GPT Image 2 输出 + Figma 内嵌）
- 多尺寸预览：16/32/48/64/128/256 + 桌面 desktop preview + taskbar preview
- 最终选定方案的 vector 重绘（Figma frame + SVG export）

### 1.3 输出

- Figma 文件 link
- 4 页都完整（不能只做 workbench 跳过 audit）
- 把 Figma link + 4 页 screenshot 写到 `docs/plans/2026-05-25-ui-refresh-figma.md`

### 1.4 Commit

```
v1.0: UI Refresh Figma 设计稿 + Current Audit（addendum 2 批次 1）

- Figma 文件 4 页：audit / design system / main workbench / icon system
- 双基准设计（1120x806 + 960x680）
- 标注 v0.9 已有 5 个 objectName 设计 token 现状
- docs/plans/2026-05-25-ui-refresh-figma.md 含 Figma link + 关键截图
```

### Codex 自检 checklist 批次 1

- [ ] Figma 连接器在 SPRINT 启动**前**已验证可用（如不可用应该 STOP，不在此 checklist 出现）
- [ ] Figma 4 页齐全（不允许跳过 audit 或 icon system）
- [ ] Current Audit 含中英文双语截图（v0.9 现状）
- [ ] Design System 含完整 token 表（颜色 / 字体 / 形状 / 状态）
- [ ] Main Workbench 双基准设计（1120x806 + 960x680）
- [ ] Icon System 含 3 个概念 + 最终方向
- [ ] 没有用 Canva / Mermaid / 本地 PNG 拼接冒充

---

## 批次 2 — 图标重做

### 2.1 GPT Image 2 生成概念

提示词参考：
- "Cloud atlas application icon, soft pastel teal-indigo gradient, abstract map pin + contour lines, minimal flat design, no text, 256x256 transparent background"
- 输出 3 个变体

### 2.2 本地重绘

GPT Image 2 输出**不直接作为最终 ico**，本地用 vector 工具（Figma export / Affinity / Inkscape）重绘：
- 干净、可缩放、无文字
- 16px 尺寸下仍可识别（关键：图标不能在 taskbar 缩小到 16px 后糊成一团）
- 圆角底板（Windows 标准应用图标）或透明外轮廓（视小尺寸可识别度而定，按 Codex 判断）
- alpha 通道干净（无白色锯齿）

### 2.3 多尺寸输出

**`filings_atlas.ico`**：必须含 16/32/48/64/128/256 六个尺寸
**`filings_atlas_logo.png`**：256x256（标题栏 28x28 显示用），可选 512x512 高清版

### 2.4 资产测试

新增 `development/tests/test_assets.py`：

```python
from pathlib import Path
from PIL import Image

_ASSETS = Path(__file__).parents[1] / "app" / "assets"


def test_logo_png_exists():
    logo = _ASSETS / "filings_atlas_logo.png"
    assert logo.exists()
    with Image.open(logo) as img:
        assert img.size[0] >= 128


def test_ico_has_required_sizes():
    ico = _ASSETS / "filings_atlas.ico"
    assert ico.exists()
    required = {16, 32, 48, 64, 128, 256}
    with Image.open(ico) as img:
        sizes = set()
        # ICO container exposes via `n_frames` + `seek`
        for i in range(img.n_frames):
            img.seek(i)
            sizes.add(img.size[0])
    missing = required - sizes
    assert not missing, f"ICO missing sizes: {missing}"
```

### 2.5 验证

```bash
cd development
python -m pytest tests/test_assets.py -v
```

### 2.6 Commit

```
v1.0: UI Refresh 图标重做（addendum 2 批次 2）

- GPT Image 2 生成 3 个云端地图针概念
- 本地重绘最终图标，alpha 干净
- filings_atlas.ico 含 16/32/48/64/128/256 多尺寸
- filings_atlas_logo.png 256x256
- 新增 tests/test_assets.py 锁定资产存在性 + ico 尺寸
```

### Codex 自检 checklist 批次 2

- [ ] GPT Image 2 概念图保留在 `docs/plans/ui_refresh_icon_concepts/` 作为参考
- [ ] 最终图标本地重绘（不直接用 AI 生成图作为 production icon）
- [ ] 16px 缩略图可识别（用 PIL 缩到 16px 后人眼能认出）
- [ ] ico 含 6 个标准尺寸（test_assets 通过）
- [ ] PySide 应用启动后 EXE 图标 + 标题栏 + taskbar 一致
- [ ] 资源文件名保持 `filings_atlas.ico` / `filings_atlas_logo.png`（不动 spec extra_datas）

---

## 批次 3 — PySide UI 重构 + QSS/palette

### 3.1 改动范围

**主要改 UI 层文件**：
- `development/app/ui/main_view.py`（主工作台结构）
- `development/app/ui/main_window.py`（titlebar 调整 + brand icon 渲染）
- `development/app/ui/exchange_selector.py`（ExchangeChip 视觉重构）
- `development/app/ui/exchange_panel.py`（panel 标题 + subtitle 优化）
- `development/app/ui/ticker_row.py`（如有状态徽标需求）
- `development/app/ui/period_selector.py`（如有调整）
- `development/app/ui/output_card.py`（如有调整）
- `development/app/ui/progress_dock.py`（如有调整）
- `development/app/ui/styles/app.qss`（QSS 视觉系统更新）
- `development/app/ui/styles/palette.py`（palette token 扩展，含深色主题）
- `development/app/ui/styles/qss_loader.py`（如需新 token）

**不动**（红线）：
- `development/app/core/*`（orchestrator / http / sidecar / settings / 任何业务逻辑）
- 7 个其他市场 plugin 业务逻辑
- `development/app/main.py` stdio 守护
- `development/plugins/jp/edinet_web.py` 等公网爬虫
- `development/filings_atlas.spec` 打包配置（资源文件名保持，hidden imports 不动）

### 3.2 结构要求（从 Figma Main Workbench 落地）

按用户诉求"顶部 Atlas 品牌区、市场图谱选择区、当前市场录入区、报告/输出配置区、固定底部操作栏"5 块结构：

1. **顶部 Atlas 品牌区**：保留现有 `_TitleBar`，但 brand icon 用新 logo；保留中/EN 切换 `#LangSegment`
2. **市场图谱选择区**：ExchangeChip 重新设计（保留 8 chip + accent 颜色映射逻辑），但视觉走 Atlas/地图风
3. **当前市场录入区**：TickerRow 视觉调整，加状态徽标（OK/失败/Pending 三态）
4. **报告/输出配置区**：PeriodSelector + OutputCard 视觉对齐
5. **固定底部操作栏**：保留"设置 / 增量更新 / 抓报告"3 按钮 + 现有 cursor + variant 属性

### 3.3 i18n + 现有测试不能破

**硬性约束**：
- 任何新增 QLabel 必须接入 `LanguageManager.language_changed` 信号 + `_retranslate()` 方法
- 任何新增 UI 字符串必须加到 `strings.py` 的 `STRINGS["zh"]` + `STRINGS["en"]` 双语字典（不允许 inline literal）
- 现有测试必须不破：
  - `tests/test_language_switch.py`（遍历 QLabel 断言 en 模式无 CJK）
  - `tests/test_ui_strings.py`（zh/en key 对齐 + en 无 CJK）
  - `tests/test_packaging_spec.py`（pandas.plotting 锁定）
  - 全量 `tests/` 181 passed 不退化

### 3.4 palette 扩展

`palette.py` 新增字段（建议，不强制）：
- `map_line`: 等高线 / 地图骨架色（云灰）
- `route_accent`: 航线 / 连线 hover 色
- `tile_bg` / `tile_border`: market tile 背景边框
- `pill_ok` / `pill_error` / `pill_pending`: 状态徽标三态

保留 `light_palette` / `dark_palette` 双套（v0.9 已有，扩展时同步）。

### 3.5 验证

```bash
cd development
python -m pytest -m "not e2e" -v
ruff check .
```

预期 181+（如新增 test_assets 后约 183 passed）+ ruff 干净。

### 3.6 Commit

```
v1.0: UI Refresh PySide 落地 + QSS（addendum 2 批次 3）

- main_view.py 重构为 5 块工作台结构（品牌/市场图谱/录入/输出/操作栏）
- ExchangeChip 视觉重做（Atlas 地图风）
- app.qss + palette.py Atlas 浅色系统 + 深色主题对应
- 新增 status pill 三态视觉
- 任何新 UI 字符串走 strings.py 双语 dict + _retranslate
- 全量 pytest 不退化，ruff 干净
```

### Codex 自检 checklist 批次 3

- [ ] 全量 pytest 不退化（含 test_language_switch / test_ui_strings / test_packaging_spec）
- [ ] 8 ExchangeChip 全部可点击 + accent 颜色保留
- [ ] 中/EN 切换无遗漏 widget（test_language_switch 遍历 QLabel 通过）
- [ ] 任何新 QSS 选择器在 zh 和 en 模式下都正确（视觉对称）
- [ ] 960x680 最小窗口尺寸下无文字重叠 / 按钮溢出
- [ ] 设置 / 增量更新 / 抓报告 3 按钮全部可用（行为不变）
- [ ] orchestrator / http / 插件 / sidecar 0 改动（grep `git diff --stat` 验证）
- [ ] ruff 干净

---

## 🔴 Reviewer Checkpoint（批次 3 后）

Reviewer Claude Code 启动新会话，执行：

```bash
cd "E:\Claude+CODEX Project\FS Capture/development"
python -m pytest -m "not e2e" -v
ruff check .
```

**逐项审**：
1. ☐ 全量 ~183 passed + ruff 干净
2. ☐ Figma 4 页齐全（含 audit 不只 workbench）
3. ☐ 图标 `filings_atlas.ico` 6 个尺寸（test_assets 通过）
4. ☐ logo + ico 视觉一致（同一 brand）
5. ☐ main_view 5 块结构正确（品牌/市场图谱/录入/输出/操作栏）
6. ☐ 8 ExchangeChip 业务行为不变（点击选中、accent 色保留）
7. ☐ 中/EN 切换无 CJK 残留（test_language_switch 通过）
8. ☐ 960x680 无重叠（手动验证截图）
9. ☐ 设置 / 增量更新 / 抓报告 3 按钮可用
10. ☐ `git diff --stat 6fb8c3d..HEAD development/app/core/ development/plugins/` 0 行（核心 0 改动）
11. ☐ palette 扩展不破现有 token（HeroTitle / SectionLabel / LangSegment 选择器仍有效）

**通过条件**：11 项全过。

---

## 批次 4 — 截图更新 + 资产测试 + 文档同步

### 4.1 截图更新

手动启动 EXE（用最新代码源码或新打包），分别截图：
- `docs/screenshots/main_window_zh.png` — 中文模式 1120x806
- `docs/screenshots/main_window_en.png` — 英文模式 1120x806
- 可选：`docs/screenshots/main_window_zh_compact.png` — 960x680 最小窗口（双语任选）

### 4.2 文档同步

**CHANGELOG.md v1.0.0 段**新增：

```markdown
### UI

- Workbench refresh with Atlas map identity (cloud-light palette, route accents, market tiles).
- New application icon redesigned with cloud-pin concept; ico now ships 16/32/48/64/128/256 sizes.
- Multilingual screenshots refreshed for both Chinese and English UI.
```

**README.md**：
- 顶部截图替换（如 README 有）
- 中英文段如有"UI 描述"句段同步

**PROJECT_RETROSPECTIVE.md** § 14.6 新增（继 § 14.5 之后）：

```markdown
### 14.6 UI Refresh + Icon Redesign（v1.0 addendum 2）

v0.9 已有的 UI polish 在 v1.0 release 前再做一轮系统重构：
- Figma 4 页设计稿前置（audit / design system / workbench / icon）
- 图标重做（GPT Image 2 概念 + 本地重绘，多尺寸 ico）
- PySide 主界面重构为 5 块结构（品牌 / 市场图谱 / 录入 / 输出 / 操作栏）
- palette / QSS 升级为云端 Atlas 浅色系统 + 深色主题对应
- 现有测试不退化（test_language_switch / test_ui_strings / test_packaging_spec 全过）
- 业务层 0 改动（orchestrator / http / 插件 / sidecar）
- 工作量：4-5 天
```

### 4.3 v1.0.0 tag 移动（**Codex 不做，由 Reviewer 通过后用户授权再做**）

Codex 批次 4 完成 commit 后**不要**移动 tag。等：
1. Reviewer Checkpoint 通过
2. 用户审最终 UI + 决定 push

然后由 Planner / 用户做：
```bash
git tag -d v1.0.0
git tag v1.0.0 HEAD
git push origin main
git push origin v1.0.0
```

### 4.4 Commit

```
v1.0: UI Refresh 截图更新 + 文档同步（addendum 2 批次 4）

- docs/screenshots/main_window_{zh,en}.png 重拍
- CHANGELOG v1.0 段加 UI section
- PROJECT_RETROSPECTIVE § 14.6 v1.0 UI refresh postscript
- 不移动 v1.0.0 tag，等用户审核
```

### Codex 自检 checklist 批次 4

- [ ] 中英文截图各 1 张（1120x806）
- [ ] 截图清晰可辨（无截图压缩损失）
- [ ] CHANGELOG / RETROSPECTIVE 准确反映本 sprint 工作
- [ ] **未移动 v1.0.0 tag**（等用户授权）
- [ ] 工作树干净（最终 commit 后 `git status` 应空）

---

## 关键文件路径速查

### 新建
- `docs/plans/2026-05-25-ui-refresh-figma.md`（Figma link + 4 页截图）
- `docs/plans/ui_refresh_icon_concepts/`（GPT Image 2 概念图保留）
- `development/tests/test_assets.py`（ico/logo 资产锁定）
- `docs/screenshots/main_window_{zh,en}.png`（重拍后覆盖）

### 修改
- `development/app/ui/main_view.py`、`main_window.py`、`exchange_selector.py`、`exchange_panel.py`、`ticker_row.py`、`period_selector.py`、`output_card.py`、`progress_dock.py`（按需）
- `development/app/ui/styles/app.qss`、`palette.py`、可能 `qss_loader.py`
- `development/app/ui/strings.py`（如新增 UI 字符串）
- `development/app/assets/filings_atlas.ico` + `filings_atlas_logo.png`（**保持文件名**）
- `CHANGELOG.md` / `PROJECT_RETROSPECTIVE.md` / `README.md`

### 不动（红线）
- `development/app/core/*` 全部（orchestrator / http / sidecar / settings / output_paths / 等）
- `development/plugins/*` 全部业务逻辑
- `development/app/main.py` stdio 守护
- `development/filings_atlas.spec`（资源文件名保持 → spec 不需改）
- v0.9 + v1.0 已落地的所有功能性测试

---

## 复用现有基础设施

| 设施 | 文件 | 用途 |
|---|---|---|
| i18n 框架 | `app/ui/i18n.py::LanguageManager` | 任何新 widget 接入 |
| 双语字典 | `app/ui/strings.py::STRINGS` | 任何新字符串走 zh/en 双份 |
| palette token | `app/ui/styles/palette.py::light_palette` | 扩展现有 token 而非替换 |
| QSS token 替换 | `app/ui/styles/qss_loader.py` | $token 替换机制 |
| Frameless window | `app/ui/main_window.py::_TitleBar` | titlebar 重构基础 |

---

## 风险与缓解（Top 5）

1. **Figma 连接器不可用 / 工具暴露不全**
   - **缓解**：批次 1 启动前**强制验证**，不可用立刻 STOP；不允许 fallback 到其他工具

2. **图标 16px 小尺寸糊成团**
   - **缓解**：批次 2 自检 checklist 明确要求 16px 可识别；本地重绘必须人眼验证

3. **UI 重构破坏现有 i18n 测试**
   - **缓解**：硬性约束 — 任何新 QLabel 接 `_retranslate`，任何字符串走 STRINGS dict；批次 3 自检跑 test_language_switch

4. **palette 扩展破坏现有 5 个 objectName 设计 token**
   - **缓解**：palette 用**扩展**模式（加新 token），不删旧 token；Reviewer Checkpoint 单独验

5. **意外动到业务层（core / 插件）**
   - **缓解**：批次 3 自检 `git diff --stat 6fb8c3d..HEAD app/core/ plugins/` 必须 0 行；Reviewer 再 verify

---

## 不在本 sprint 范围

- ❌ 新增市场（v1.1+）
- ❌ 抓取性能优化
- ❌ release 策略调整（v1.0 还是首发 release）
- ❌ 业务逻辑改动（任何 plugin / orchestrator / http）
- ❌ 测试 marker / pytest 配置调整
- ❌ 切换 PySide6 主版本
- ❌ 第三语言 UI
- ❌ 跨平台打包（仍 Windows-only）

---

## 验证

### 全量回归（每批必绿）
```bash
cd development
python -m pytest -m "not e2e" -v
ruff check .
```
**v1.0 当前基线**：181 passed → **预期 ~183**（+test_assets 2-3 个）

### 资产测试
```bash
cd development
python -m pytest tests/test_assets.py -v
```

### 手动验证
- 启动源码 `python -m app.main`：3 秒内主窗口可见
- 中/EN 切换：所有可见文本无 CJK 残留
- 960x680 最小窗口：无文字重叠 / 按钮溢出
- 8 ExchangeChip 全部可点击
- 设置 / 增量更新 / 抓报告 3 按钮行为不变
- 旧 FS 视觉残留扫描（特别是 ico/png + titlebar + readme）

### EXE 打包验证（批次 3 后）
```powershell
cd development
.\build.bat
# 启动根目录 EXE，截图 zh + en
```

---

## 最终 Reviewer Checklist

- [ ] Figma 4 页齐全
- [ ] 图标 6 尺寸（test_assets）
- [ ] PySide 落地 + ~183 passed + ruff 干净
- [ ] 8 ExchangeChip 业务不变
- [ ] 双语切换无 CJK 残留
- [ ] 960x680 无重叠
- [ ] core / 插件 0 改动
- [ ] CHANGELOG / RETROSPECTIVE / README 同步
- [ ] 截图重拍（zh + en）
- [ ] v1.0.0 tag **未**移动（等用户授权）

---

## 下一步（用户操作）

1. **审批本 SPRINT 文档**
2. **验证 Figma 连接器可用**（在 Codex 启动批次 1 前必做）
3. **把本文档喂给 Codex 启动批次 1**
4. 批次 3 后 Reviewer Checkpoint
5. Reviewer 通过 + 用户审最终 UI
6. 移动 v1.0.0 tag + push origin + 发 GitHub release

---

**Planner 签名**：Claude Code (Opus 4.7)
**日期**：2026-05-24
**关联**：
- `roadmap/SPRINT_v1.0_sg_and_perf.md`（主 sprint）
- `roadmap/SPRINT_v1.0_addendum_jp.md`（addendum 1）
- 本文（addendum 2）

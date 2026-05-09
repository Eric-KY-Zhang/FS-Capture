# FS Capture

> 一键批量下载 A 股 / 港股 / 美股 / 韩股上市公司的官方披露文件（年报 / 审计报告 / 季报 / 半年报 / IPO 招股书）。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![PySide6](https://img.shields.io/badge/GUI-PySide6-green.svg)](https://doc.qt.io/qtforpython-6/)

FS Capture 是一个 Windows 桌面 EXE 工具，专注解决「批量拿到原始 PDF」这一件事——**不抓三大报表数字、不算财务指标、不生成 Excel 底稿**。需要财务数据 / Excel 装填的场景请使用相关 VBA 工具。

## 功能特点

- **4 个市场覆盖**：A 股、港股、美股、韩股，每个交易所有独立的插件实现
- **多种报告类型**：年报、独立审计报告、一季报、半年报、三季报、IPO 招股书
- **批量并发下载**：QThreadPool 多 worker，per-source TokenBucket 限速避免 IP 封禁
- **现代 GUI**：PySide6 无边框圆角窗口 + 扁平化设计 + 行内确认 + 进度面板
- **PDF 元数据 sidecar**：每个 PDF 旁写 `.meta.json`（含 sha256、source URL、下载时间）
- **本地优先**：HTTP / 限速 / 缓存 / 配置全在本地，无服务端依赖
- **PyInstaller 打包**：单 EXE 双击运行，~340 MB 自包含

## 数据源

| 市场 | 公司名称 | 报告下载 | 鉴权 |
|---|---|---|---|
| A 股 | akshare + cninfo orgId | cninfo `hisAnnouncement` | 不需要 |
| 港股 | 东方财富 `stock/get` | HKEXnews `titlesearch.xhtml` | 不需要 |
| 美股 | SEC `company_tickers.json` | SEC `submissions API` + 分页 fallback | 不需要 |
| 韩股 | OpenDartReader `corp_codes` | DART `list` + `document` | DART API Key（[免费注册](https://opendart.fss.or.kr/)） |

## 日常使用（已构建好的 EXE）

1. 双击根目录的 `FS Capture.exe`
2. 勾选交易所，添加股票代码，点击「确认」识别公司名称
3. 选择年份区间和报告类型（年报 / 季报 / 半年报 / IPO 招股书）
4. 点击「开始抓取」
5. 结果保存在 `output/` 文件夹，PDF 平铺，文件名含市场 / 代码 / 年份元信息

韩股需要 DART API Key——[免费注册](https://opendart.fss.or.kr/)后在程序内「设置」中粘贴，或直接编辑 `config.toml`。若暂时无 Key，可不勾选韩股，A/HK/US 三市场不受影响。

## 从源码运行

```bash
git clone https://github.com/<your-username>/fs-capture.git
cd fs-capture/development
python -m pip install -r requirements.txt
python -m playwright install chromium    # PDF 兜底渲染（KR/US 部分场景）
python -m app.main
```

要求：Windows 10/11 + Python 3.11~3.14。

## 构建 EXE

```bat
cd development
build.bat
```

PyInstaller one-folder 模式，产物在 `development/dist/FS Capture/`，约 340 MB。

## 配置

`config.toml` 由首次启动 onboarding 引导生成，关键字段：

```toml
[concurrency]
max_workers = 4               # 并发任务数

[rate_limits]
cninfo = 5                    # 各数据源独立限速 (rps)
hkexnews = 3
sec = 8
dart = 5
akshare = 4

[sec]
user_agent = "FS Capture (your-email@example.com)"   # SEC 政策强制带邮箱

[dart]
api_key = ""                  # 在此粘贴 DART API key

[ui]
theme = "light"               # light | dark
```

完整模板见 [`config.example.toml`](config.example.toml)。

## 测试

```bash
cd development
pytest -m "not e2e"           # 单元 + 集成测试，~30 个，不联网
FS_CAPTURE_RUN_E2E=1 pytest -m e2e   # 联网真打 4 市场冒烟，~25 秒
ruff check . && ruff format --check .
```

GitHub Actions CI 在 `.github/workflows/ci.yml` 跑 lint + 非 e2e 测试。

## 文档

- [`ARCHITECTURE.md`](ARCHITECTURE.md) — 项目架构 + 模块依赖图 + invariants
- [`PROJECT_RETROSPECTIVE.md`](PROJECT_RETROSPECTIVE.md) — 开发复盘 + 教训
- [`development/DEVELOPMENT_BRIEF.md`](development/DEVELOPMENT_BRIEF.md) — 给代码审核者的开发委托书
- [`roadmap/`](roadmap/) — Sprint 计划归档（v0.1 → v0.5）

## 项目结构

```text
fs-capture/
├── README.md                 # 本文件
├── LICENSE                   # MIT
├── ARCHITECTURE.md
├── PROJECT_RETROSPECTIVE.md
├── config.example.toml
├── roadmap/                  # Sprint plans
└── development/              # 源码
    ├── pyproject.toml
    ├── requirements.txt
    ├── fs_capture.spec       # PyInstaller spec (one-folder)
    ├── run.bat / build.bat
    ├── app/                  # GUI (PySide6) + core (httpx, ratelimit, cache, orchestrator)
    ├── plugins/              # 4 市场 plugin: ashare / hk / us / kr
    └── tests/                # pytest unit + integration + e2e
```

## 隐私与合规

- HTTP 请求只走公开 GET / POST，无登录绕过、无私人数据抓取
- SEC 请求自动注入用户邮箱 User-Agent（SEC fair-use 政策强制）
- DART API Key 与下载内容均落在本地 `config.toml` / `output/`，无外部上传
- 不做交易、不抓实时行情

## 许可

[MIT License](LICENSE) © 2026 Eric Zhang

## 致谢

- [akshare](https://github.com/akfamily/akshare) — 中港美股数据源聚合
- [OpenDartReader](https://github.com/FinanceData/OpenDartReader) — DART OpenAPI 包装器
- [pypdf](https://github.com/py-pdf/pypdf) — HK 选片 PDF 文本验证
- [PySide6](https://wiki.qt.io/Qt_for_Python) — Qt6 Python 绑定
- 工具的 v0.1 → v0.5 5 个 sprint 在 Claude Code（Planner + Reviewer）↔ Codex（Generator）协作模式下完成。详见 [`roadmap/`](roadmap/) 与 [`PROJECT_RETROSPECTIVE.md`](PROJECT_RETROSPECTIVE.md)。

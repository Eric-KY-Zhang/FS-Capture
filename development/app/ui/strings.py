"""UI strings — bilingual zh/en, keyed by widget."""

from __future__ import annotations

from app.ui.i18n import LanguageManager

STRINGS: dict[str, dict[str, str]] = {
    "zh": {
        # Common
        "COMMON_CANCEL": "取消",
        # batch_import_dialog
        "BID_WINDOW_TITLE_FORMAT": "批量添加{exchange_name}股票",
        "BID_TITLE": "粘贴股票代码",
        "BID_HINT": "支持从 Excel、网页或文本复制，多行、逗号和制表符会自动拆分。",
        "BID_AUTO_CONFIRM": "添加后自动确认公司名称",
        "BID_ADD": "添加",
        # exchange_panel
        "EP_BATCH_ADD": "＋ 批量添加",
        "EP_SINGLE_ADD": "＋ 单只添加",
        "EP_EMPTY": "暂无股票，点击右上角添加或批量导入",
        "EP_REJECTED_TITLE": "存在未识别代码",
        "EP_REJECTED_SUFFIX_FORMAT": "\n... 另有 {count} 行未显示",
        "EP_REJECTED_BODY_FORMAT": (
            "以下 {count} 项未识别，可能格式错误或与所选市场不匹配：\n{preview}{suffix}"
        ),
        "EP_NO_CODES_TITLE": "没有可添加的代码",
        "EP_NO_CODES_BODY": "未识别到有效股票代码",
        "EP_BATCH_DONE_TITLE": "批量添加完成",
        "EP_BATCH_DONE_BODY_FORMAT": "已添加 {added} 只股票，跳过 {skipped} 个重复代码。",
        "EP_TITLE_A_SHARE": "A股 · A-Share",
        "EP_TITLE_HK": "港股 · Hong Kong",
        "EP_TITLE_US": "美股 · United States",
        "EP_TITLE_KR": "韩股 · Korea",
        "EP_TITLE_TW": "台股 · Taiwan",
        "EP_TITLE_JP": "日股 · Japan",
        "EP_TITLE_UK": "英股 · United Kingdom",
        "EP_TITLE_SG": "新股 · Singapore",
        "EP_SUBTITLE_A_SHARE": "巨潮资讯网 · 东方财富",
        "EP_SUBTITLE_HK": "披露易 · 东方财富",
        "EP_SUBTITLE_US": "SEC EDGAR 披露系统",
        "EP_SUBTITLE_KR": "DART 电子公示",
        "EP_SUBTITLE_TW": "公開資訊觀測站 MOPS",
        "EP_SUBTITLE_JP": "EDINET · 东京证券交易所",
        "EP_SUBTITLE_UK": "FCA 国家存储机制 NSM",
        "EP_SUBTITLE_SG": "SGXNet 公开披露",
        # exchange_selector
        "ES_NAME_A_SHARE": "A股",
        "ES_NAME_HK": "港股",
        "ES_NAME_US": "美股",
        "ES_NAME_KR": "韩股",
        "ES_NAME_TW": "台股",
        "ES_NAME_JP": "日股",
        "ES_NAME_UK": "英股",
        "ES_NAME_SG": "新股",
        "ES_META_A_SHARE": "上交所 · 深交所 · 北交所",
        "ES_META_HK": "香港交易所",
        "ES_META_US": "纽交所 · 纳斯达克",
        "ES_META_KR": "韩国交易所 KOSPI · KOSDAQ",
        "ES_META_TW": "台交所 · 櫃買中心",
        "ES_META_JP": "EDINET · 东京证券交易所",
        "ES_META_UK": "伦敦证券交易所 · NSM",
        "ES_META_SG": "新加坡交易所 SGX",
        # main_view
        "MV_SECTION": "FILINGS ATLAS · 全球披露图谱",
        "MV_TITLE": "全球披露图谱",
        "MV_SUBTITLE": "跨 8 大市场 · 上市公司官方披露 · 一键 PDF 归档",
        "MV_SETTINGS_BUTTON": "⚙  设置",
        "MV_RUN_BUTTON": "▶  抓报告",
        "MV_INCREMENTAL_BUTTON": "增量更新",
        "MV_CANNOT_START_TITLE": "无法开始",
        "MV_NO_TICKERS_BODY": "请先添加并确认至少一只股票",
        "MV_UNCONFIRMED_SUFFIX": " 中存在尚未确认的股票代码",
        "MV_CONTINUE_TITLE": "继续？",
        "MV_ONLY_CONFIRMED_BODY": "\n\n仅继续抓取已确认的股票？",
        "MV_NO_PERIOD_BODY": "请至少选择一种报告类型",
        "MV_KR_NO_KEY_LOG": "未配置 DART OpenAPI Key，韩股将使用 DART 公网披露页，速度可能较慢。",
        "MV_JP_NO_KEY_LOG": "未配置 EDINET API Key，日股将使用 EDINET 公网搜索页；配置 Key 可加速。",
        "MV_INVALID_PATH_TITLE": "无效路径",
        "MV_INVALID_PATH_BODY": "请选择有效的输出文件夹",
        "MV_JOB_SUBMITTED_FORMAT": (
            "提交报告下载任务：{tickers} 只股票 × {periods} 个期间 = {tasks} 个 task"
        ),
        "MV_JOB_STARTED_FORMAT": "Job 开始，共 {total} 个 task",
        "MV_INCREMENTAL_SKIPPED_FORMAT": "增量模式：跳过 {count} 个已下载任务",
        "MV_INCREMENTAL_NONE_TITLE": "无新报告",
        "MV_INCREMENTAL_NONE_BODY": "所选股票/期间均已下载过。",
        "MV_DONE_TITLE": "下载完成",
        "MV_DONE_BODY_FORMAT": (
            "总计 {total} 个 task：成功 {ok}，失败 {fail}\n\n"
            "已下载 {reports} 个 PDF\n"
            "输出位置：{output_dir}"
        ),
        # main_window
        "MW_WINDOW_TITLE": "Filings Atlas / 全球披露图谱",
        "MW_SUBTITLE": "全球披露图谱 · 跨市场官方披露归档",
        # onboarding_dialog
        "OB_WINDOW_TITLE": "欢迎使用 Filings Atlas / 全球披露图谱",
        "OB_TITLE": "Filings Atlas / 全球披露图谱帮你一键下载 8 市场上市公司官方披露 PDF",
        "OB_HINT": "输入第一个股票代码试试。确认公司名称后，选择年份和报告类型即可开始下载。",
        "OB_DART_BODY": (
            "韩股默认通过 DART 公网披露页抓取，无需配置。如需更快更稳的体验，"
            '可在 <a href="https://opendart.fss.or.kr/">opendart.fss.or.kr</a> '
            "免费申请 API Key 后填入设置。"
        ),
        "OB_LATER": "稍后再说",
        "OB_SETTINGS": "先填 Key（可选）",
        # output_card
        "OC_TITLE": "输出位置",
        "OC_BROWSE": "浏览…",
        "OC_OPEN_DIR": "打开文件夹",
        "OC_SELECT_DIR_TITLE": "选择输出文件夹",
        # period_selector
        "PS_TITLE": "报告期间",
        "PS_SUBTITLE": "选择年份区间和期间报告",
        "PS_FROM_YEAR": "起始年份",
        "PS_TO_YEAR": "终止年份",
        "PS_TYPE_LABEL": "期间类型",
        "PS_ANNUAL": "年报",
        "PS_Q3": "三季报",
        "PS_Q2": "半年报",
        "PS_Q1": "一季报",
        "PS_IPO": "IPO 招股书",
        # progress_dock
        "PD_WAITING": "等待中",
        "PD_TITLE": "抓取进度",
        "PD_STARTED": "已开始",
        "PD_RUNNING": "进行中",
        "PD_FILE_COUNT_FORMAT": "文件 {count}",
        "PD_NO_FILE": "无文件",
        "PD_DONE": "✓ 完成",
        "PD_FAILED": "× 失败",
        "PD_CANCELLED": "已取消",
        "PD_PERIOD_LABEL_FORMAT": "{year}年度 {period_type}",
        # settings_dialog
        "SD_TITLE": "设置",
        "SD_DART_PLACEHOLDER": "可选——留空走 DART 公网爬虫",
        "SD_DART_LABEL": "DART OpenAPI Key（可选，韩股加速）",
        "SD_EDINET_PLACEHOLDER": "可选——留空走 EDINET 公网搜索",
        "SD_EDINET_LABEL": "EDINET API Key（可选，日股加速）",
        "SD_WORKERS_LABEL": "并发数",
        "SD_THEME_LABEL": "主题（重启生效）",
        "SD_THEME_LIGHT": "浅色",
        "SD_THEME_DARK": "深色",
        "SD_LANGUAGE_LABEL": "界面语言",
        "SD_LANGUAGE_ZH": "中文",
        "SD_LANGUAGE_EN": "英文",
        "SD_SEC_UA_LABEL": "SEC User-Agent（美国市场）",
        "SD_SAVE": "保存",
        # ticker_row
        "TR_CONFIRM": "确认",
        "TR_PENDING": "待确认",
        "TR_DELETE_TOOLTIP": "移除此行",
        "TR_ENTER_CODE": "请输入代码",
        "TR_RESOLVING": "解析中…",
        "TR_RESOLVED": "✓ 已确认",
        "TR_NOT_FOUND": "未找到",
        "TR_UNRESOLVABLE": "无法解析此代码",
        "TR_PLACEHOLDER_A_SHARE": "如 600519 / 000001",
        "TR_PLACEHOLDER_HK": "如 00700 / 9988",
        "TR_PLACEHOLDER_US": "如 AAPL / TSLA",
        "TR_PLACEHOLDER_KR": "如 005930 / 000660",
        "TR_PLACEHOLDER_TW": "如 2330 / 2317",
        "TR_PLACEHOLDER_JP": "如 7203 / 6758",
        "TR_PLACEHOLDER_UK": "如 ULVR / HSBA",
        "TR_PLACEHOLDER_SG": "如 D05 / U11 / Z74",
    },
    "en": {
        # Common
        "COMMON_CANCEL": "Cancel",
        # batch_import_dialog
        "BID_WINDOW_TITLE_FORMAT": "Batch Add {exchange_name} Tickers",
        "BID_TITLE": "Paste Ticker Codes",
        "BID_HINT": "Paste from Excel, web pages, or text. Lines, commas, and tabs are split automatically.",
        "BID_AUTO_CONFIRM": "Confirm company names after adding",
        "BID_ADD": "Add",
        # exchange_panel
        "EP_BATCH_ADD": "+ Batch Add",
        "EP_SINGLE_ADD": "+ Add One",
        "EP_EMPTY": "No tickers yet. Add one or import a batch from the top right.",
        "EP_REJECTED_TITLE": "Unrecognized codes",
        "EP_REJECTED_SUFFIX_FORMAT": "\n... {count} more rows hidden",
        "EP_REJECTED_BODY_FORMAT": (
            "{count} entries were not recognized. They may be malformed or not valid for this market:\n{preview}{suffix}"
        ),
        "EP_NO_CODES_TITLE": "No codes to add",
        "EP_NO_CODES_BODY": "No valid ticker codes were recognized",
        "EP_BATCH_DONE_TITLE": "Batch add complete",
        "EP_BATCH_DONE_BODY_FORMAT": "Added {added} tickers and skipped {skipped} duplicates.",
        "EP_TITLE_A_SHARE": "A-Share",
        "EP_TITLE_HK": "Hong Kong",
        "EP_TITLE_US": "United States",
        "EP_TITLE_KR": "Korea",
        "EP_TITLE_TW": "Taiwan",
        "EP_TITLE_JP": "Japan",
        "EP_TITLE_UK": "United Kingdom",
        "EP_TITLE_SG": "Singapore",
        "EP_SUBTITLE_A_SHARE": "CNINFO · Eastmoney",
        "EP_SUBTITLE_HK": "HKEXnews · Eastmoney",
        "EP_SUBTITLE_US": "SEC EDGAR",
        "EP_SUBTITLE_KR": "DART electronic disclosure",
        "EP_SUBTITLE_TW": "MOPS disclosure portal",
        "EP_SUBTITLE_JP": "EDINET · TSE",
        "EP_SUBTITLE_UK": "FCA National Storage Mechanism",
        "EP_SUBTITLE_SG": "SGXNet public disclosures",
        # exchange_selector
        "ES_NAME_A_SHARE": "A-Share",
        "ES_NAME_HK": "Hong Kong",
        "ES_NAME_US": "United States",
        "ES_NAME_KR": "Korea",
        "ES_NAME_TW": "Taiwan",
        "ES_NAME_JP": "Japan",
        "ES_NAME_UK": "United Kingdom",
        "ES_NAME_SG": "Singapore",
        "ES_META_A_SHARE": "SSE · SZSE · BSE",
        "ES_META_HK": "Hong Kong Exchange",
        "ES_META_US": "NYSE · NASDAQ",
        "ES_META_KR": "KOSPI · KOSDAQ",
        "ES_META_TW": "TWSE · TPEx",
        "ES_META_JP": "EDINET · TSE",
        "ES_META_UK": "LSE · NSM",
        "ES_META_SG": "Singapore Exchange",
        # main_view
        "MV_SECTION": "FILINGS ATLAS · GLOBAL DISCLOSURE",
        "MV_TITLE": "Filings Atlas",
        "MV_SUBTITLE": "Eight markets · Official disclosures · One-click PDF archive",
        "MV_SETTINGS_BUTTON": "⚙  Settings",
        "MV_RUN_BUTTON": "▶  Download Reports",
        "MV_INCREMENTAL_BUTTON": "Incremental Update",
        "MV_CANNOT_START_TITLE": "Cannot start",
        "MV_NO_TICKERS_BODY": "Add and confirm at least one ticker first",
        "MV_UNCONFIRMED_SUFFIX": " has unconfirmed ticker codes",
        "MV_CONTINUE_TITLE": "Continue?",
        "MV_ONLY_CONFIRMED_BODY": "\n\nContinue with confirmed tickers only?",
        "MV_NO_PERIOD_BODY": "Select at least one report type",
        "MV_KR_NO_KEY_LOG": "No DART OpenAPI key configured. Korea will use the public DART disclosure pages and may be slower.",
        "MV_JP_NO_KEY_LOG": "No EDINET API key configured. Japan will use the public EDINET search; adding a key can speed it up.",
        "MV_INVALID_PATH_TITLE": "Invalid path",
        "MV_INVALID_PATH_BODY": "Choose a valid output folder",
        "MV_JOB_SUBMITTED_FORMAT": (
            "Submitted report download job: {tickers} tickers x {periods} periods = {tasks} tasks"
        ),
        "MV_JOB_STARTED_FORMAT": "Job started with {total} tasks",
        "MV_INCREMENTAL_SKIPPED_FORMAT": "Incremental: skipped {count} already-downloaded tasks",
        "MV_INCREMENTAL_NONE_TITLE": "Nothing New",
        "MV_INCREMENTAL_NONE_BODY": "All selected ticker/period pairs are already downloaded.",
        "MV_DONE_TITLE": "Download complete",
        "MV_DONE_BODY_FORMAT": (
            "{total} tasks total: {ok} succeeded, {fail} failed\n\n"
            "Downloaded {reports} PDFs\n"
            "Output folder: {output_dir}"
        ),
        # main_window
        "MW_WINDOW_TITLE": "Filings Atlas",
        "MW_SUBTITLE": "Global Disclosure Atlas · Cross-market archive",
        # onboarding_dialog
        "OB_WINDOW_TITLE": "Welcome to Filings Atlas",
        "OB_TITLE": "Filings Atlas downloads official disclosure PDFs across 8 markets",
        "OB_HINT": "Try your first ticker code. Confirm the company name, choose years and report types, then start downloading.",
        "OB_DART_BODY": (
            "Korea uses the public DART disclosure pages by default and does not require setup. "
            'For faster and more stable downloads, apply for a free API key at <a href="https://opendart.fss.or.kr/">opendart.fss.or.kr</a> '
            "and add it in Settings."
        ),
        "OB_LATER": "Later",
        "OB_SETTINGS": "Add key (Optional)",
        # output_card
        "OC_TITLE": "Output",
        "OC_BROWSE": "Browse...",
        "OC_OPEN_DIR": "Open folder",
        "OC_SELECT_DIR_TITLE": "Choose output folder",
        # period_selector
        "PS_TITLE": "Report period",
        "PS_SUBTITLE": "Choose years and report types",
        "PS_FROM_YEAR": "From year",
        "PS_TO_YEAR": "To year",
        "PS_TYPE_LABEL": "Period type",
        "PS_ANNUAL": "Annual report",
        "PS_Q3": "Q3",
        "PS_Q2": "Interim",
        "PS_Q1": "Q1",
        "PS_IPO": "IPO prospectus",
        # progress_dock
        "PD_WAITING": "Waiting",
        "PD_TITLE": "Download progress",
        "PD_STARTED": "Started",
        "PD_RUNNING": "Running",
        "PD_FILE_COUNT_FORMAT": "Files {count}",
        "PD_NO_FILE": "No files",
        "PD_DONE": "Done",
        "PD_FAILED": "Failed",
        "PD_CANCELLED": "Cancelled",
        "PD_PERIOD_LABEL_FORMAT": "{period_type} {year}",
        # settings_dialog
        "SD_TITLE": "Settings",
        "SD_DART_PLACEHOLDER": "Optional - leave blank to use the public DART crawler",
        "SD_DART_LABEL": "DART OpenAPI key (Optional, speeds up Korea)",
        "SD_EDINET_PLACEHOLDER": "Optional - leave blank to use public EDINET search",
        "SD_EDINET_LABEL": "EDINET API key (Optional, speeds up Japan)",
        "SD_WORKERS_LABEL": "Concurrency",
        "SD_THEME_LABEL": "Theme (restart required)",
        "SD_THEME_LIGHT": "Light",
        "SD_THEME_DARK": "Dark",
        "SD_LANGUAGE_LABEL": "UI language",
        "SD_LANGUAGE_ZH": "Chinese",
        "SD_LANGUAGE_EN": "English",
        "SD_SEC_UA_LABEL": "SEC User-Agent",
        "SD_SAVE": "Save",
        # ticker_row
        "TR_CONFIRM": "Confirm",
        "TR_PENDING": "Pending",
        "TR_DELETE_TOOLTIP": "Remove this row",
        "TR_ENTER_CODE": "Enter a ticker code",
        "TR_RESOLVING": "Resolving...",
        "TR_RESOLVED": "Confirmed",
        "TR_NOT_FOUND": "Not found",
        "TR_UNRESOLVABLE": "Cannot resolve this code",
        "TR_PLACEHOLDER_A_SHARE": "e.g. 600519 / 000001",
        "TR_PLACEHOLDER_HK": "e.g. 00700 / 9988",
        "TR_PLACEHOLDER_US": "e.g. AAPL / TSLA",
        "TR_PLACEHOLDER_KR": "e.g. 005930 / 000660",
        "TR_PLACEHOLDER_TW": "e.g. 2330 / 2317",
        "TR_PLACEHOLDER_JP": "e.g. 7203 / 6758",
        "TR_PLACEHOLDER_UK": "e.g. ULVR / HSBA",
        "TR_PLACEHOLDER_SG": "e.g. D05 / U11 / Z74",
    },
}


def tr(key: str) -> str:
    return STRINGS[LanguageManager.instance().current_language][key]


# Language-independent constants — shown identically regardless of UI language.
# Used by the title-bar 中/EN segmented toggle so an English speaker can find
# the language switch without first reading the current UI language.
LANG_INDEPENDENT: dict[str, str] = {
    "MW_LANG_TOGGLE_ZH": "中",
    "MW_LANG_TOGGLE_EN": "EN",
    "MW_LANG_TOGGLE_TOOLTIP": "中文 / English",
}


def __getattr__(name: str) -> str:
    if name in LANG_INDEPENDENT:
        return LANG_INDEPENDENT[name]
    if name in STRINGS["zh"]:
        return tr(name)
    raise AttributeError(f"module 'strings' has no attribute {name!r}")

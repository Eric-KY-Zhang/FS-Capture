"""UI string constants centralized for future i18n."""

from __future__ import annotations

# Common
COMMON_CANCEL = "取消"

# batch_import_dialog
BID_WINDOW_TITLE_FORMAT = "批量添加{exchange_name}股票"
BID_TITLE = "粘贴股票代码"
BID_HINT = "支持从 Excel、网页或文本复制，多行、逗号和制表符会自动拆分。"
BID_AUTO_CONFIRM = "添加后自动确认公司名称"
BID_ADD = "添加"

# exchange_panel
EP_BATCH_ADD = "＋ 批量添加"
EP_SINGLE_ADD = "＋ 单只添加"
EP_EMPTY = "暂无股票，点击右上角添加或批量导入"
EP_REJECTED_TITLE = "存在未识别代码"
EP_REJECTED_SUFFIX_FORMAT = "\n... 另有 {count} 行未显示"
EP_REJECTED_BODY_FORMAT = (
    "以下 {count} 项未识别，可能格式错误或与所选市场不匹配：\n{preview}{suffix}"
)
EP_NO_CODES_TITLE = "没有可添加的代码"
EP_NO_CODES_BODY = "未识别到有效股票代码"
EP_BATCH_DONE_TITLE = "批量添加完成"
EP_BATCH_DONE_BODY_FORMAT = "已添加 {added} 只股票，跳过 {skipped} 个重复代码。"
EP_TITLE_A_SHARE = "A股 · A-Share"
EP_TITLE_HK = "港股 · Hong Kong"
EP_TITLE_US = "美股 · United States"
EP_TITLE_KR = "韩股 · Korea"
EP_TITLE_TW = "台股 · Taiwan"
EP_SUBTITLE_A_SHARE = "巨潮资讯网 · 东方财富"
EP_SUBTITLE_HK = "披露易 · 东方财富"
EP_SUBTITLE_KR = "DART 电子公示"
EP_SUBTITLE_TW = "公開資訊觀測站 MOPS"

# exchange_selector
ES_NAME_A_SHARE = "A股"
ES_NAME_HK = "港股"
ES_NAME_US = "美股"
ES_NAME_KR = "韩股"
ES_NAME_TW = "台股"
ES_META_A_SHARE = "上交所 · 深交所 · 北交所"
ES_META_HK = "香港交易所"
ES_META_TW = "台交所 · 櫃買中心"

# main_view
MV_SECTION = "FILINGS ATLAS · 一键抓取"
MV_TITLE = "上市公司官方披露 PDF 下载"
MV_SUBTITLE = "勾选交易所 → 录入股票代码 → 选择期间 → 下载披露文件 PDF"
MV_SETTINGS_BUTTON = "⚙  设置"
MV_RUN_BUTTON = "▶  抓报告"
MV_CANNOT_START_TITLE = "无法开始"
MV_NO_TICKERS_BODY = "请先添加并确认至少一只股票"
MV_UNCONFIRMED_SUFFIX = " 中存在尚未确认的股票代码"
MV_CONTINUE_TITLE = "继续？"
MV_ONLY_CONFIRMED_BODY = "\n\n仅继续抓取已确认的股票？"
MV_NO_PERIOD_BODY = "请至少选择一种报告类型"
MV_KR_NO_KEY_LOG = "未配置 DART OpenAPI Key，韩股将使用 DART 公网披露页，速度可能较慢。"
MV_INVALID_PATH_TITLE = "无效路径"
MV_INVALID_PATH_BODY = "请选择有效的输出文件夹"
MV_JOB_SUBMITTED_FORMAT = (
    "提交报告下载任务：{tickers} 只股票 × {periods} 个期间 = {tasks} 个 task"
)
MV_JOB_STARTED_FORMAT = "Job 开始，共 {total} 个 task"
MV_DONE_TITLE = "下载完成"
MV_DONE_BODY_FORMAT = (
    "总计 {total} 个 task：成功 {ok}，失败 {fail}\n\n"
    "已下载 {reports} 个 PDF\n"
    "输出位置：{output_dir}"
)

# main_window
MW_WINDOW_TITLE = "Filings Atlas / 全球披露图谱"
MW_SUBTITLE = "· 上市公司披露报告一键下载"

# onboarding_dialog
OB_WINDOW_TITLE = "欢迎使用 Filings Atlas / 全球披露图谱"
OB_TITLE = "Filings Atlas / 全球披露图谱帮你一键下载 5 市场上市公司官方披露 PDF"
OB_HINT = "输入第一个股票代码试试。确认公司名称后，选择年份和报告类型即可开始下载。"
OB_DART_BODY = (
    "韩股默认通过 DART 公网披露页抓取，无需配置。如需更快更稳的体验，"
    '可在 <a href="https://opendart.fss.or.kr/">opendart.fss.or.kr</a> '
    "免费申请 API Key 后填入设置。"
)
OB_LATER = "稍后再说"
OB_SETTINGS = "先填 Key（可选）"

# output_card
OC_TITLE = "输出位置"
OC_BROWSE = "浏览…"
OC_OPEN_DIR = "打开文件夹"
OC_SELECT_DIR_TITLE = "选择输出文件夹"

# period_selector
PS_TITLE = "报告期间"
PS_SUBTITLE = "选择年份区间和期间报告"
PS_FROM_YEAR = "起始年份"
PS_TO_YEAR = "终止年份"
PS_TYPE_LABEL = "期间类型"
PS_ANNUAL = "年报"
PS_Q3 = "三季报"
PS_Q2 = "半年报"
PS_Q1 = "一季报"
PS_IPO = "IPO 招股书"

# progress_dock
PD_WAITING = "等待中"
PD_TITLE = "抓取进度"
PD_STARTED = "已开始"
PD_RUNNING = "进行中"
PD_FILE_COUNT_FORMAT = "文件 {count}"
PD_NO_FILE = "无文件"
PD_DONE = "✓ 完成"
PD_FAILED = "× 失败"
PD_CANCELLED = "已取消"

# settings_dialog
SD_TITLE = "设置"
SD_DART_PLACEHOLDER = "可选——留空走 DART 公网爬虫"
SD_DART_LABEL = "DART OpenAPI Key（可选，韩股加速）"
SD_WORKERS_LABEL = "并发数"
SD_THEME_LABEL = "主题（重启生效）"
SD_SAVE = "保存"

# ticker_row
TR_CONFIRM = "确认"
TR_PENDING = "待确认"
TR_DELETE_TOOLTIP = "移除此行"
TR_ENTER_CODE = "请输入代码"
TR_RESOLVING = "解析中…"
TR_RESOLVED = "✓ 已确认"
TR_NOT_FOUND = "未找到"
TR_UNRESOLVABLE = "无法解析此代码"
TR_PLACEHOLDER_A_SHARE = "如 600519 / 000001"
TR_PLACEHOLDER_HK = "如 00700 / 9988"
TR_PLACEHOLDER_US = "如 AAPL / TSLA"
TR_PLACEHOLDER_KR = "如 005930 / 000660"
TR_PLACEHOLDER_TW = "如 2330 / 2317"

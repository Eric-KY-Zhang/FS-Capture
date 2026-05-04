"""
把 modules/*.bas 通过 Excel COM 自动注入到 上市公司财务数据查询.xlsm

为什么走 COM 不让用户手动 Import?
  - VBE 的『Import File』在中文 Windows 上对 UTF-8 .bas 偶尔会乱码
  - COM .CodeModule.AddFromString 走 Unicode, 永远不会编码出错
  - 而且 build_template.py 输出的是 .xlsx (openpyxl 不能直接生成有效 .xlsm),
    本脚本顺便用 Excel COM 把 .xlsx 转成 .xlsm

前置:
  - Excel 已安装 (本机 Office 任何版本)
  - 信任访问 VBA 项目对象模型: Excel 选项 → 信任中心 → 宏设置 → 勾选『信任对 VBA 工程对象模型的访问』
  - pywin32 (一般 Office 预装, 否则 `py -m pip install pywin32`)

用法:
    cd "E:\\Claude+CODEX Project\\FS Capture\\VBA Captor"
    py tools/install_modules.py

效果:
    1. 找到 上市公司财务数据查询.xlsx (build_template.py 的输出)
    2. 用 Excel COM 打开, 注入 .bas 模块, 另存为 上市公司财务数据查询.xlsm
    3. 删除中转 .xlsx (避免歧义)
"""

import sys
from pathlib import Path

import win32com.client as win32

VBE_CT_STDMODULE = 1   # vbext_ct_StdModule
XL_FILEFORMAT_XLSM = 52  # xlOpenXMLWorkbookMacroEnabled
MSO_SHAPE_ROUNDED_RECT = 5
MSO_ANCHOR_CENTER = 2
MSO_ANCHOR_MIDDLE = 3
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1


def rgb_long(hex_str: str) -> int:
    """Convert HTML hex color (e.g. 'FF4472C4' or '4472C4') to Excel BGR Long."""
    h = hex_str.lstrip("#").lstrip("FF") if len(hex_str.lstrip("#")) == 8 else hex_str.lstrip("#")
    if len(h) != 6:
        h = h[-6:]
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r


def set_cell_comment(cell, text: str):
    """Replace an Excel note/comment with explicit text; works around AddComment text quirks."""
    try:
        if cell.Comment is not None:
            cell.Comment.Delete()
    except Exception:
        pass
    try:
        cell.AddComment()
        cell.Comment.Text(text)
    except Exception:
        try:
            cell.AddComment(text)
        except Exception:
            pass


# Phase 4e 按钮规格: (name, caption, macro, target_range, fill_hex, font_color_hex, font_size, primary?)
PRIMARY_FILL = "4472C4"   # 深蓝
PRIMARY_FG = "FFFFFF"     # 白
SECONDARY_FILL = "D9E1F2" # 浅蓝
SECONDARY_FG = "1F4E79"   # 深蓝字
CROSS_FILL = "ED7D31"     # 橙 — 跨市场对比

US_FILL = "C00000"        # 深红 — 美股按钮区分
US_FG = "FFFFFF"
HK_FILL = "548235"        # 深绿 — 港股按钮区分
HK_FG = "FFFFFF"
KR_FILL = "7030A0"        # 深紫 — 韩股按钮区分
KR_FG = "FFFFFF"

BUTTONS = [
    ("BtnRunAll",       "一键全抓 4 市场",     "模块_总入口.一键全抓",            "Q1:Q3", SECONDARY_FILL, SECONDARY_FG, 13, True),
    ("BtnHideAll",      "显示/隐藏 所有市场数据", "模块_总入口.切换所有分市场tabs", "Q5:Q7", SECONDARY_FILL, SECONDARY_FG, 11, True),
    ("BtnClearCache",   "清空 HTTP 缓存",       "模块_工具函数.ClearLocalCache",   "Q9:Q11", SECONDARY_FILL, SECONDARY_FG, 10, True),
    ("BtnRunA",         "一键 A 股",           "模块_总入口.一键A股",             "A8:B8", PRIMARY_FILL,   PRIMARY_FG,   10, True),
    ("BtnRunUS",        "一键 美股",           "模块_总入口.一键美股",            "E8:F8", US_FILL,        US_FG,        12, True),
    ("BtnRunHK",        "一键 港股",           "模块_总入口.一键港股",            "I8:J8", HK_FILL,        HK_FG,        12, True),
    ("BtnRunKR",        "一键 韩股",           "模块_总入口.一键韩股",            "M8:N8", KR_FILL,        KR_FG,        12, True),
    ("BtnHideA",        "显示/隐藏 A 股数据",    "模块_总入口.切换A股tabs",          "A9:B9", SECONDARY_FILL, SECONDARY_FG, 9, False),
    ("BtnHideUS",       "显示/隐藏 美股数据",    "模块_总入口.切换美股tabs",         "E9:F9", SECONDARY_FILL, SECONDARY_FG, 9, False),
    ("BtnHideHK",       "显示/隐藏 港股数据",    "模块_总入口.切换港股tabs",         "I9:J9", SECONDARY_FILL, SECONDARY_FG, 9, False),
    ("BtnHideKR",       "显示/隐藏 韩股数据",    "模块_总入口.切换韩股tabs",         "M9:N9", SECONDARY_FILL, SECONDARY_FG, 9, False),
]

# 已废弃: install 时从当前 xlsm 主动移除 (即使 modules/ 下仍有遗留也清掉)
DECOMMISSIONED_MODULES = ["模块_抓基本资料"]
DECOMMISSIONED_BUTTONS = [
    "BtnRunInfo",
    "BtnBuildCrossAll",
    "BtnHideCrossMarket",
    "BtnBuildCrossInd",
    "BtnBuildCrossBS",
    "BtnBuildCrossIS",
    "BtnBuildCrossCF",
    "BtnRunBalance",
    "BtnRunProfit",
    "BtnRunCash",
    "BtnRunInd",
    "BtnRunUSBalance",
    "BtnRunUSProfit",
    "BtnRunUSCash",
    "BtnRunUSInd",
    "BtnRunHKBalance",
    "BtnRunHKProfit",
    "BtnRunHKCash",
    "BtnRunHKInd",
    "BtnRunKRBalance",
    "BtnRunKRProfit",
    "BtnRunKRCash",
    "BtnRunKRInd",
]

SHEET_RENAMES = {
    # Phase 2 → Phase 3: 去 _宽表 后缀
    "资产负债表_宽表": "A股_资产负债表",
    "利润表_宽表": "A股_利润表",
    "现金流量表_宽表": "A股_现金流量表",
    "指标表_宽表": "A股_指标表",
    # Phase 4b-3 → 4b-4: A 股 sheet 加 A股_ 前缀
    "资产负债表": "A股_资产负债表",
    "利润表": "A股_利润表",
    "现金流量表": "A股_现金流量表",
    "指标表": "A股_指标表",
}

# 已废弃 sheet (install 时主动删除)
#   - 上市公司基本资料: Phase 4b-3 不再使用
#   - 资产负债表/利润表/现金流量表/指标表 (无前缀): Phase 4b-4 改用 A股_ 前缀;
#       理论上 SHEET_RENAMES 处理大部分情况, 但若新旧 sheet 都存在 (rename skip 后)
#       这里兜底删除老的 (用户数据已迁到 A股_ 前缀的新 sheet 里)
DECOMMISSIONED_SHEETS = [
    "上市公司基本资料",
    "跨市场_资产负债表",
    "跨市场_利润表",
    "跨市场_现金流量表",
    "字段映射",
    "资产负债表",
    "利润表",
    "现金流量表",
    "指标表",
]

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "上市公司财务数据查询.xlsx"
XLSM = ROOT / "上市公司财务数据查询.xlsm"
LEGACY_XLSX = ROOT / "新浪财经行业数据查询V3.xlsx"
LEGACY_XLSM = ROOT / "新浪财经行业数据查询V3.xlsm"
MODULES_DIR = ROOT / "modules"


def install_quarter_cell(ws_pool):
    """
    Phase 3: 在样本池 A3 / A4 装季度选择器 (如果不存在)
      A3 = "季度 (Q1/Q2/Q3/Q4 或 全部)"  (label)
      A4 = 当前值, 默认 "全部"; data validation = list of 全部/Q1/Q2/Q3/Q4
    """
    cell_label = ws_pool.Range("A3")
    cell_value = ws_pool.Range("A4")

    # 仅在 A3 还空 (用户没自己填过) 时才覆盖, 避免破坏用户后续手工编辑
    if not cell_label.Value:
        cell_label.Value = "季度 (Q1/Q2/Q3/Q4 或 全部)"
        cell_label.Font.Name = "微软雅黑"
        cell_label.Font.Size = 10
        cell_label.Font.Bold = True
        cell_label.Interior.Color = rgb_long("B4C7E7")    # 浅蓝
        cell_label.HorizontalAlignment = -4108            # xlCenter
        cell_label.VerticalAlignment = -4108
        print("  + 季度标签 A3 已写入")
    else:
        print(f"  ~ A3 已有内容, 保留: {cell_label.Value!r}")

    if not cell_value.Value:
        cell_value.Value = "全部"
        cell_value.Font.Name = "微软雅黑"
        cell_value.Font.Size = 11
        cell_value.Font.Bold = True
        cell_value.Interior.Color = rgb_long("FFE699")    # 浅黄突出
        cell_value.HorizontalAlignment = -4108
        cell_value.VerticalAlignment = -4108
        print("  + 季度默认值 A4=全部 已写入")
    else:
        print(f"  ~ A4 已有内容, 保留: {cell_value.Value!r}")

    # Data validation (无论 A4 原来是什么, 都重置 list 验证以保证下拉箭头可用)
    try:
        cell_value.Validation.Delete()
    except Exception:
        pass
    try:
        cell_value.Validation.Add(
            Type=XL_VALIDATE_LIST,
            AlertStyle=XL_VALID_ALERT_STOP,
            Operator=1,    # xlBetween
            Formula1="全部,Q1,Q2,Q3,Q4",
        )
        cell_value.Validation.IgnoreBlank = False
        cell_value.Validation.InCellDropdown = True
        print("  + A4 数据验证 (下拉) 已加")
    except Exception as e:
        print(f"  ! A4 数据验证添加失败: {e}")


def install_xueqiu_cookie_cell(ws_pool):
    """
    Phase 4b-5: 样本池 row 5 装雪球 cookie 输入位
      A5 = 标签『雪球 Cookie』
      B5:C5 合并 = cookie 值 (用户从 浏览器登录 xueqiu.com 后 F12 → Cookies → 拷 xq_a_token 值)
    """
    label_cell = ws_pool.Range("A5")
    if not label_cell.Value:
        label_cell.Value = "雪球 Cookie"
    label_cell.Font.Name = "微软雅黑"
    label_cell.Font.Size = 10
    label_cell.Font.Bold = True
    label_cell.Interior.Color = rgb_long("B4C7E7")    # 浅蓝
    label_cell.HorizontalAlignment = -4108
    label_cell.VerticalAlignment = -4108

    # B5:C5 合并做长 cookie 值 cell
    val_range = ws_pool.Range("B5:C5")
    try:
        val_range.UnMerge()
    except Exception:
        pass
    val_range.Merge()
    # 不覆盖用户已粘的值
    if not ws_pool.Range("B5").Value:
        ws_pool.Range("B5").Value = ""    # 占位
    val_cell = ws_pool.Range("B5")
    val_cell.Font.Name = "Consolas"
    val_cell.Font.Size = 9
    val_cell.Interior.Color = rgb_long("FFF2CC")    # 浅黄, 提示用户填写
    val_cell.HorizontalAlignment = -4131    # left
    val_cell.VerticalAlignment = -4108
    val_cell.WrapText = True

    # 给 B5 加一个批注/提示 (Excel Comment)
    try:
        if val_cell.Comment is None:
            val_cell.AddComment(
                "雪球 Cookie (美股中概/20-F fallback 与港股抓数使用)\n\n"
                "1. 浏览器打开 https://xueqiu.com (登录 / 不登录都可以)\n"
                "2. F12 → Application → Cookies → xueqiu.com\n"
                "3. 找到 xq_a_token, 拷它的 Value\n"
                "4. 粘到这个单元格\n\n"
                "VBA 会在美股 EDGAR 404 时自动 fallback 到雪球, 港股也会调用雪球.\n"
                "Cookie 有效期约 1 个月, 过期 API 会报 400016, 重新拷一次"
            )
    except Exception:
        pass

    print("  + A5/B5 雪球 Cookie 输入位已配置 (美股 fallback + 港股)")


def _cell_text(cell) -> str:
    """Return Excel cell text without losing leading zeros when possible."""
    try:
        text = str(cell.Text or "").strip()
        if text and set(text) != {"#"}:
            return text
    except Exception:
        pass
    value = cell.Value
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalize_market(value, code: str) -> str:
    raw = str(value or "").strip().upper()
    aliases = {
        "A": "A", "ASHARE": "A", "A股": "A", "沪深": "A", "CN": "A",
        "US": "US", "USA": "US", "美股": "US", "美国": "US",
        "HK": "HK", "H": "HK", "港股": "HK", "香港": "HK",
        "KR": "KR", "KOREA": "KR", "韩股": "KR", "韩国": "KR",
    }
    if raw in aliases:
        return aliases[raw]
    c = str(code or "").strip()
    if c.isalpha():
        return "US"
    if c.isdigit() and len(c) == 5:
        return "HK"
    if c.isdigit() and len(c) == 6:
        return "A"
    return ""


def _normalize_code_for_market(code: str, market: str) -> str:
    c = str(code or "").strip()
    if not c:
        return ""
    if market == "HK" and c.isdigit():
        return c.zfill(5)
    if market == "KR" and c.isdigit():
        return c.zfill(6)
    return c.upper() if market == "US" else c


def migrate_old_sample_pool(ws_pool):
    """
    旧布局 A:C (代码/简称/市场) 或更早 H=市场 布局迁移到 Phase 4e 四市场分栏。
    新布局已存在时保持幂等,不重复迁移。
    """
    a7 = str(ws_pool.Range("A7").Value or "").strip()
    e7 = str(ws_pool.Range("E7").Value or "").strip()
    if "A 股" in a7 and "美股" in e7:
        print("  ~ 样本池 已是 4 市场分栏布局, 跳过迁移")
        return

    c7 = str(ws_pool.Range("C7").Value or "").strip()
    h7 = str(ws_pool.Range("H7").Value or "").strip()
    if c7 == "市场":
        market_col = 3
        print("  ~ 检测到旧样本池 A:C 布局, 开始迁移到 4 市场分栏")
    elif h7 == "市场":
        market_col = 8
        print("  ~ 检测到更早样本池 H=市场 布局, 开始迁移到 4 市场分栏")
    else:
        print(f"  ~ 未检测到旧样本池布局 (A7={a7!r}, C7={c7!r}, H7={h7!r}), 仅刷新新布局")
        return

    last_row = ws_pool.Cells(ws_pool.Rows.Count, 1).End(-4162).Row  # xlUp
    if last_row < 8:
        print("  ~ 旧样本池无公司数据, 跳过迁移")
        return

    by_market = {"A": [], "US": [], "HK": [], "KR": []}
    for row in range(8, last_row + 1):
        code = _cell_text(ws_pool.Cells(row, 1))
        if not code:
            continue
        name = _cell_text(ws_pool.Cells(row, 2))
        market_value = _cell_text(ws_pool.Cells(row, market_col))
        market = _normalize_market(market_value, code)
        if market not in by_market:
            print(f"  ! 跳过无法判断市场的样本: row={row}, code={code!r}, market={market_value!r}")
            continue
        by_market[market].append((_normalize_code_for_market(code, market), name))

    try:
        ws_pool.Range("A7:Q1000").UnMerge()
    except Exception:
        pass
    ws_pool.Range("A7:Q1000").Clear()

    col_map = {"A": (1, 2), "US": (5, 6), "HK": (9, 10), "KR": (13, 14)}
    for market, companies in by_market.items():
        code_col, name_col = col_map[market]
        for idx, (code, name) in enumerate(companies, start=11):
            ws_pool.Cells(idx, code_col).NumberFormat = "@"
            ws_pool.Cells(idx, code_col).Value = code
            ws_pool.Cells(idx, name_col).Value = name

    summary = " / ".join(f"{m}: {len(v)}" for m, v in by_market.items())
    print(f"  + 样本池迁移完成: {summary}")


def migrate_phase4g_sample_rows(ws_pool):
    """
    Phase 4g: 已经是 4 市场分栏的 Phase 4e/4f 工作簿,公司数据原来从 Row 10 开始。
    新布局 Row 10 留给表头,因此仅在旧 Row 10 有样本时把 Row 10+ 下移一行。
    """
    if _cell_text(ws_pool.Range("A10")) == "代码":
        print("  ~ 样本池 Row 10 已是新表头,跳过 Phase 4g 数据行迁移")
        return

    header_addrs = ("A9", "B9", "E9", "F9", "I9", "J9", "M9", "N9")
    old_headers = [_cell_text(ws_pool.Range(addr)) for addr in header_addrs]
    if "代码" not in old_headers or "简称" not in old_headers:
        return

    market_cols = (1, 2, 5, 6, 9, 10, 13, 14)
    row10_has_data = any(_cell_text(ws_pool.Cells(10, col)) for col in market_cols)
    if not row10_has_data:
        print("  ~ 样本池 Row 10 无公司数据,跳过 Phase 4g 数据行迁移")
        return

    last_row = max(ws_pool.Cells(ws_pool.Rows.Count, col).End(-4162).Row for col in market_cols)
    if last_row < 10:
        return

    for row in range(last_row, 9, -1):
        for col in market_cols:
            src = ws_pool.Cells(row, col)
            dst = ws_pool.Cells(row + 1, col)
            dst.NumberFormat = src.NumberFormat
            dst.Value = src.Value

    print("  + Phase 4g 样本池数据行已从 Row 10+ 下移到 Row 11+")


def layout_sample_pool(ws_pool):
    """重画 Phase 4g 样本池配置区、市场分栏和数据区格式,不清除 Row 11+ 公司数据。"""
    year_value = ws_pool.Range("A2").Value or 2025
    quarter_value = ws_pool.Range("A4").Value or "全部"
    cookie_value = ws_pool.Range("B5").Value or ""
    # Phase 4f Step 2: 保留用户已选的 B6 显示币种 (空 → install_currency_toggle_cell 写默认)
    currency_value = ws_pool.Range("B6").Value or ""

    migrate_phase4g_sample_rows(ws_pool)

    try:
        ws_pool.Range("A1:R10").UnMerge()
        ws_pool.Range("S1:U24").UnMerge()
        ws_pool.Range("W1:Y5").UnMerge()
    except Exception:
        pass
    ws_pool.Range("A1:R10").Clear()
    ws_pool.Range("S1:U24").Clear()
    ws_pool.Range("W1:Y5").Clear()

    widths = {
        "A": 11, "B": 16, "C": 2, "D": 2,
        "E": 8, "F": 18, "G": 2, "H": 2,
        "I": 7, "J": 14, "K": 2, "L": 2,
        "M": 8, "N": 16, "O": 2, "P": 2,
        "Q": 13, "R": 16, "S": 2, "T": 13, "U": 18,
        "V": 2, "W": 18, "X": 18, "Y": 24,
    }
    for col, width in widths.items():
        ws_pool.Columns(col).ColumnWidth = width

    label_fill = rgb_long("B4C7E7")
    value_fill = rgb_long("FFE699")
    cookie_fill = rgb_long("FFF2CC")

    for addr, value in (
        ("A1", "年份 (留空=取最新)"),
        ("A3", "季度 (Q1/Q2/Q3/Q4 或 全部)"),
        ("A5", "雪球 Cookie"),
        ("A6", "显示币种"),
    ):
        cell = ws_pool.Range(addr)
        cell.Value = value
        cell.Font.Name = "微软雅黑"
        cell.Font.Size = 10
        cell.Font.Bold = True
        cell.Interior.Color = label_fill
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108

    ws_pool.Range("A2").Value = year_value
    ws_pool.Range("A4").Value = quarter_value
    for addr in ("A2", "A4"):
        cell = ws_pool.Range(addr)
        cell.Font.Name = "微软雅黑"
        cell.Font.Size = 11
        cell.Font.Bold = True
        cell.Interior.Color = value_fill
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108

    try:
        ws_pool.Range("A4").Validation.Delete()
        ws_pool.Range("A4").Validation.Add(
            Type=XL_VALIDATE_LIST,
            AlertStyle=XL_VALID_ALERT_STOP,
            Operator=1,
            Formula1="全部,Q1,Q2,Q3,Q4",
        )
        ws_pool.Range("A4").Validation.IgnoreBlank = False
        ws_pool.Range("A4").Validation.InCellDropdown = True
    except Exception as e:
        print(f"  ! A4 数据验证添加失败: {e}")

    try:
        ws_pool.Range("B5:F5").UnMerge()
    except Exception:
        pass
    ws_pool.Range("B5:F5").Merge()
    val_cell = ws_pool.Range("B5")
    val_cell.Value = cookie_value
    val_cell.Font.Name = "Consolas"
    val_cell.Font.Size = 9
    val_cell.Interior.Color = cookie_fill
    val_cell.HorizontalAlignment = -4131
    val_cell.VerticalAlignment = -4108
    val_cell.WrapText = True
    try:
        if val_cell.Comment is None:
            val_cell.AddComment(
                "雪球登录凭证,用于港股和部分中概美股备用数据源。\n"
                "如果港股或中概美股取数失败,请重新登录雪球后复制新的 xq_a_token。"
            )
    except Exception:
        pass

    # Phase 4f Step 2: 恢复 B6 显示币种 (clear A1:Q10 已清掉, 这里写回; 空 → 后续 install_currency_toggle_cell 写默认)
    if currency_value:
        ws_pool.Range("B6").Value = currency_value

    markets = [
        ("A7:B7", "A 股(新浪)", PRIMARY_FILL),
        ("E7:F7", "美股(EDGAR+雪球)", US_FILL),
        ("I7:J7", "港股(雪球 HK)", HK_FILL),
        ("M7:N7", "韩股(stockanalysis)", KR_FILL),
    ]
    for addr, caption, fill_hex in markets:
        rng = ws_pool.Range(addr)
        if addr.split(":")[0] != addr.split(":")[-1]:
            rng.Merge()
        rng.Value = caption
        rng.Font.Name = "微软雅黑"
        rng.Font.Size = 11
        rng.Font.Bold = True
        rng.Font.Color = rgb_long("FFFFFF")
        rng.Interior.Color = rgb_long(fill_hex)
        rng.HorizontalAlignment = -4108
        rng.VerticalAlignment = -4108

    placeholders = [
        ("A8:B8", "一键 A 股", PRIMARY_FILL, "FFFFFF", 10),
        ("E8:F8", "一键 美股", US_FILL, "FFFFFF", 11),
        ("I8:J8", "一键 港股", HK_FILL, "FFFFFF", 11),
        ("M8:N8", "一键 韩股", KR_FILL, "FFFFFF", 11),
        ("Q1:Q3", "一键全抓 4 市场", SECONDARY_FILL, SECONDARY_FG, 12),
        ("Q5:Q7", "显示/隐藏 所有市场数据", SECONDARY_FILL, SECONDARY_FG, 10),
        ("Q9:Q11", "清空 HTTP 缓存", SECONDARY_FILL, SECONDARY_FG, 10),
        ("A9:B9", "显示/隐藏 A 股数据", SECONDARY_FILL, SECONDARY_FG, 9),
        ("E9:F9", "显示/隐藏 美股数据", SECONDARY_FILL, SECONDARY_FG, 9),
        ("I9:J9", "显示/隐藏 港股数据", SECONDARY_FILL, SECONDARY_FG, 9),
        ("M9:N9", "显示/隐藏 韩股数据", SECONDARY_FILL, SECONDARY_FG, 9),
    ]
    for addr, caption, fill_hex, font_color_hex, font_size in placeholders:
        rng = ws_pool.Range(addr)
        if addr.split(":")[0] != addr.split(":")[-1]:
            rng.Merge()
            rng.Value = caption
        else:
            rng.Value = caption
        rng.Font.Name = "微软雅黑"
        rng.Font.Size = font_size
        rng.Font.Bold = True
        rng.Font.Color = rgb_long(font_color_hex)
        rng.Interior.Color = rgb_long(fill_hex)
        rng.HorizontalAlignment = -4108
        rng.VerticalAlignment = -4108

    section_cells = [
        ("Q4", "全局显示", "F2F2F2"),
        ("Q8", "工具", "F2F2F2"),
    ]
    for addr, caption, fill_hex in section_cells:
        cell = ws_pool.Range(addr)
        cell.Value = caption
        cell.Font.Name = "微软雅黑"
        cell.Font.Size = 9
        cell.Font.Bold = True
        cell.Font.Color = rgb_long("404040")
        cell.Interior.Color = rgb_long(fill_hex)
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108

    try:
        set_cell_comment(
            ws_pool.Range("Q9"),
            "清除本地暂存的抓数结果。\n"
            "下次抓数会重新从公开数据来源取数,适合强制刷新或排查数据陈旧问题。\n"
            "日常使用无需点击。"
        )
    except Exception:
        pass


    hint_rng = ws_pool.Range("S1:U5")
    hint_rng.Merge()
    hint_rng.Value = (
        "【使用提示】\n"
        "1. A2 / A4 选择年度 + 季度\n"
        "2. B5 填雪球 cookie;B6 切原币 / 统一RMB\n"
        "3. 第 11 行起按市场录入公司\n"
        "4. 先点各市场一键按钮抓数\n"
        "5. 一键全抓会自动刷新跨市场指标表"
    )
    hint_rng.Font.Name = "微软雅黑"
    hint_rng.Font.Size = 9
    hint_rng.Font.Color = rgb_long("7F6000")
    hint_rng.Interior.Color = rgb_long("FFF2CC")
    hint_rng.HorizontalAlignment = -4131
    hint_rng.VerticalAlignment = -4160
    hint_rng.WrapText = True

    for code_col, name_col in (("A", "B"), ("E", "F"), ("I", "J"), ("M", "N")):
        for col, caption in ((code_col, "代码"), (name_col, "简称")):
            cell = ws_pool.Range(f"{col}10")
            cell.Value = caption
            cell.Font.Name = "微软雅黑"
            cell.Font.Size = 11
            cell.Font.Bold = True
            cell.Font.Color = rgb_long("FFFFFF")
            cell.Interior.Color = rgb_long(PRIMARY_FILL)
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108

    for col in ("A", "E", "I", "M"):
        ws_pool.Range(f"{col}11:{col}1000").NumberFormat = "@"

    for row in range(1, 7):    # Phase 4f Step 2: 含 row 6 显示币种
        ws_pool.Rows(row).RowHeight = 22
    ws_pool.Rows(7).RowHeight = 24
    ws_pool.Rows(8).RowHeight = 34
    ws_pool.Rows(9).RowHeight = 22
    ws_pool.Rows(10).RowHeight = 22
    for row in range(11, 52):
        ws_pool.Rows(row).RowHeight = 20

    data_range = ws_pool.Range("A10:N51")
    data_range.Font.Name = "微软雅黑"
    data_range.Font.Size = 10
    for border_idx in (7, 8, 9, 10, 11, 12):
        try:
            b = data_range.Borders(border_idx)
            b.LineStyle = 1
            b.Weight = 2
            b.Color = rgb_long("BFBFBF")
        except Exception:
            pass

    try:
        ws_pool.Activate()
        ws_pool.Application.ActiveWindow.SplitColumn = 0
        ws_pool.Application.ActiveWindow.SplitRow = 10
        ws_pool.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    # Phase 4f Step 2: 配置 B6 显示币种 toggle (默认 "原币" + 数据验证下拉)
    install_currency_toggle_cell(ws_pool)
    try:
        set_cell_comment(
            ws_pool.Range("Q9"),
            "清除本地暂存的抓数结果。\n"
            "下次抓数会重新从公开数据来源取数,适合强制刷新或排查数据陈旧问题。\n"
            "日常使用无需点击。"
        )
    except Exception:
        pass

    print("  + 样本池 4 市场分栏布局已刷新")


def style_sample_pool_data_area(ws_pool):
    """
    Phase 4b-4: 样本池数据区美化
      - Row 7 表头: 深蓝白字, 高 22, 加粗
      - Row 1-4 配置区: A1/A3 浅蓝标签, A2/A4 浅黄值
      - A8:C50 数据区: 微软雅黑 11pt + 细灰边框 + 行高 20
      - C 列条件格式: A=浅蓝 / HK=浅黄 / US=浅红
      - D 列 = spacer (宽 3, 无内容)
    """
    # ---- 列宽 ----
    ws_pool.Columns("A").ColumnWidth = 13
    ws_pool.Columns("B").ColumnWidth = 16
    ws_pool.Columns("C").ColumnWidth = 10
    # D / E / F 由 BUTTON_COL_WIDTHS 处理

    # ---- Row 7 表头 ----
    header_row_range = ws_pool.Range("A7:C7")
    header_row_range.Font.Name = "微软雅黑"
    header_row_range.Font.Size = 11
    header_row_range.Font.Bold = True
    header_row_range.Font.Color = rgb_long("FFFFFF")
    header_row_range.Interior.Color = rgb_long("4472C4")
    header_row_range.HorizontalAlignment = -4108
    header_row_range.VerticalAlignment = -4108
    ws_pool.Rows(7).RowHeight = 24

    # ---- A8:C50 数据区 ----
    data_range = ws_pool.Range("A8:C50")
    data_range.Font.Name = "微软雅黑"
    data_range.Font.Size = 11

    # 对齐: A 中心 / B 左 / C 中心
    ws_pool.Range("A8:A50").HorizontalAlignment = -4108
    ws_pool.Range("A8:A50").VerticalAlignment = -4108
    ws_pool.Range("B8:B50").HorizontalAlignment = -4131    # left
    ws_pool.Range("B8:B50").VerticalAlignment = -4108
    ws_pool.Range("C8:C50").HorizontalAlignment = -4108
    ws_pool.Range("C8:C50").VerticalAlignment = -4108

    # 行高
    for r in range(8, 51):
        try:
            ws_pool.Rows(r).RowHeight = 20
        except Exception:
            pass

    # 边框: 细灰 (xlEdgeLeft=7, Top=8, Bottom=9, Right=10, InsideV=11, InsideH=12)
    XL_CONTINUOUS = 1
    XL_THIN = 2
    for border_idx in (7, 8, 9, 10, 11, 12):
        try:
            b = data_range.Borders(border_idx)
            b.LineStyle = XL_CONTINUOUS
            b.Weight = XL_THIN
            b.Color = rgb_long("BFBFBF")
        except Exception:
            pass

    # 也给表头一圈边框
    try:
        for border_idx in (7, 8, 9, 10, 11, 12):
            b = header_row_range.Borders(border_idx)
            b.LineStyle = XL_CONTINUOUS
            b.Weight = XL_THIN
            b.Color = rgb_long("FFFFFF")
    except Exception:
        pass

    # ---- C 列条件格式: A / HK / US / KR 不同底色 ----
    XL_CELL_VALUE = 1
    XL_EQUAL = 3
    cf_range = ws_pool.Range("C8:C50")
    try:
        cf_range.FormatConditions.Delete()
    except Exception:
        pass
    for value, color_hex in [("A", "D9E1F2"), ("HK", "FFF2CC"), ("US", "FCE4D6"), ("KR", "E4DFEC")]:
        try:
            cf = cf_range.FormatConditions.Add(
                Type=XL_CELL_VALUE, Operator=XL_EQUAL,
                Formula1=f'"{value}"',
            )
            cf.Interior.Color = rgb_long(color_hex)
            cf.Font.Bold = True
        except Exception as e:
            print(f"  ! C 列条件格式 ({value}) 添加失败: {e}")
    print("  + 样本池数据区美化完成 (表头 + 边框 + 对齐 + 市场列条件格式)")


def cleanup_legacy_sample_pool(ws_pool):
    """
    Phase 4b-3: 把老布局 (URL 列在 C-G, 市场列在 H) 迁移到新布局 (市场列在 C)。
      - 检测: 如果 H7 = "市场", 说明是老布局, 执行清理
      - 如果 C7 已经是 "市场", 跳过 (新布局, 幂等)
      - 老布局动作:
          1. 清 row 1-6 的 URL 模板 (B1:D6 cell content)
          2. 清 row 7 的 URL 表头 (C7:G7)
          3. 删除 C:G 列 (整列删除, H→C, I→D, J→E)
          4. 删除老 button shapes (位置变了, 接下来 install_buttons 会重建)
    """
    c7 = str(ws_pool.Range("C7").Value or "").strip()
    h7 = str(ws_pool.Range("H7").Value or "").strip()
    if c7 == "市场":
        print("  ~ 样本池 已是新布局 (C 列=市场), 跳过 legacy 迁移")
        return
    if h7 != "市场":
        print(f"  ~ 样本池 既不是老布局也不是新布局 (C7={c7!r}, H7={h7!r}), 跳过迁移")
        return

    print("  ~ 检测到老布局 (URL 列 + H=市场), 开始迁移到新布局 (市场列移到 C)")

    # 1. 删除老按钮 shapes (位置在 I 列, 列删除会破坏 anchor)
    legacy_btn_names = {
        "BtnRunAll", "BtnRunBalance", "BtnRunProfit", "BtnRunCash", "BtnRunInd",
        "BtnRunInfo",
        "BtnRunUSBalance", "BtnRunUSProfit", "BtnRunUSCash", "BtnRunUSInd",
        "BtnRunHKBalance", "BtnRunHKProfit", "BtnRunHKCash", "BtnRunHKInd",
        "BtnRunKRBalance", "BtnRunKRProfit", "BtnRunKRCash", "BtnRunKRInd",
    }
    shape_names_snapshot = [sh.Name for sh in ws_pool.Shapes]
    for n in shape_names_snapshot:
        if n in legacy_btn_names:
            ws_pool.Shapes(n).Delete()
    print(f"  - 删除 legacy buttons before column delete")

    # 2. 清 row 1-6 的 URL 模板
    ws_pool.Range("B1:D6").ClearContents()

    # 3. 删除整列 C:G (5 列), H 自动滑到 C
    ws_pool.Columns("C:G").Delete()
    print("  - 删除整列 C:G (URL 列 + URL 模板), 市场 H→C, 按钮列 I→D")

    # 4. C7 应该已经自动滑成 "市场" (从 H7 滑过来), 但保险起见重设
    ws_pool.Range("C7").Value = "市场"


def ensure_market_column(ws_pool):
    """
    Phase 4: 在样本池 C 列装『市场』
      - C7 表头 + 蓝底白字
      - C8:C1000 自动检测公式: 由 A 列代码推断 A股 / HK / US (用户可手填覆盖)
      - C8:C1000 数据验证下拉 (A/HK/US/KR)
    """
    header_cell = ws_pool.Range("C7")
    if not header_cell.Value:
        header_cell.Value = "市场"
    header_cell.Font.Name = "微软雅黑"
    header_cell.Font.Size = 11
    header_cell.Font.Bold = True
    header_cell.Font.Color = rgb_long("FFFFFF")
    header_cell.Interior.Color = rgb_long("4472C4")
    header_cell.HorizontalAlignment = -4108     # xlCenter
    header_cell.VerticalAlignment = -4108
    print("  + C7『市场』表头已配置")

    ws_pool.Columns("C").ColumnWidth = 10

    auto_formula = '=IF(A{r}="","",IF(ISNUMBER(--A{r}),IF(LEN(A{r})=5,"HK","A"),"US"))'
    written = 0
    for r in range(8, 1001):
        cell = ws_pool.Range(f"C{r}")
        try:
            if cell.HasFormula:
                cell.Formula = auto_formula.format(r=r)
                written += 1
            elif cell.Value in (None, ""):
                cell.Formula = auto_formula.format(r=r)
                written += 1
            # else: 用户手填的 A/HK/US/KR 硬值, 保留
        except Exception:
            pass
    print(f"  + C8:C1000 写入市场自动推断公式 ({written} 行, 含未来空白行)")

    try:
        rng = ws_pool.Range("C8:C1000")
        rng.Validation.Delete()
        rng.Validation.Add(
            Type=XL_VALIDATE_LIST,
            AlertStyle=XL_VALID_ALERT_STOP,
            Operator=1,
            Formula1="A,HK,US,KR",
        )
        rng.Validation.IgnoreBlank = True
        rng.Validation.InCellDropdown = True
        print("  + C8:C1000 市场下拉 (A/HK/US/KR) 已加")
    except Exception as e:
        print(f"  ! C 列下拉添加失败: {e}")


def _make_wide_table_sheet(wb, name):
    """创建空宽表结构 sheet。指标表使用 A:C 三列静态描述列。"""
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    is_indicator = name in ("A股_指标表", "美股_指标表", "港股_指标表", "韩股_指标表")
    ws.Range("A1").Value = "指标类型" if is_indicator else "大类"
    ws.Range("B1").Value = "指标名称"
    header_addrs = ["A1", "B1"]
    if is_indicator:
        ws.Range("C1").Value = "英文指标名"
        header_addrs.append("C1")
    for addr in header_addrs:
        c = ws.Range(addr)
        c.Font.Name = "微软雅黑"
        c.Font.Size = 11
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108     # xlCenter
        c.VerticalAlignment = -4108
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 40
    if is_indicator:
        ws.Columns("A").ColumnWidth = 18
        ws.Columns("B").ColumnWidth = 28
        ws.Columns("C").ColumnWidth = 34
    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 20
    # 冻结数据区
    try:
        ws.Activate()
        wb.Application.ActiveWindow.SplitColumn = 3 if is_indicator else 2
        wb.Application.ActiveWindow.SplitRow = 2
        wb.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass
    return ws


def _make_corp_info_sheet(wb, name):
    """创建空基本资料 sheet (平表): A=代码 B=简称 C=上市日期 D=所属行业 E=主营业务"""
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    headers = [("A", "股票代码", 12), ("B", "股票简称", 14),
               ("C", "上市日期", 14), ("D", "所属行业", 24), ("E", "主营业务", 80)]
    for col, txt, width in headers:
        c = ws.Range(f"{col}1")
        c.Value = txt
        c.Font.Name = "微软雅黑"
        c.Font.Size = 11
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        ws.Columns(col).ColumnWidth = width
    ws.Rows(1).RowHeight = 22
    return ws


def _make_diagnostic_sheet(wb, name="美股_抓取诊断"):
    """创建空抓取诊断 sheet。
    Row 1 = 大标题(合并 A1:K1, 深蓝白字), Row 2 = 11 列表头, Row 3+ 由 VBA 写
    冻结 Row 2; 列宽 + 表头颜色与 VBA 端 EnsureDiagnosticSheet() 保持一致 (避免双方互踩)。
    """
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name

    # Row 1: 标题, 合并 A1:K1
    ws.Range("A1").Value = f"{name.replace('_', '')} (每次跑数后自动刷新)"
    ws.Range("A1:K1").Merge()
    title = ws.Range("A1:K1")
    title.Font.Name = "微软雅黑"
    title.Font.Size = 12
    title.Font.Bold = True
    title.Font.Color = rgb_long("FFFFFF")
    title.Interior.Color = rgb_long("4472C4")
    title.HorizontalAlignment = -4108
    title.VerticalAlignment = -4108

    # Row 2: 11 列表头
    headers = ["公司", "报表", "输出指标", "状态", "数据源",
               "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注", "FX_Rate"]
    for j, txt in enumerate(headers, start=1):
        c = ws.Cells(2, j)
        c.Value = txt
        c.Font.Name = "微软雅黑"
        c.Font.Size = 10
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108

    widths = [14, 16, 30, 18, 18, 14, 42, 14, 10, 58, 12]
    for j, w in enumerate(widths, start=1):
        ws.Columns(j).ColumnWidth = w
    ws.Columns("A").NumberFormat = "@"
    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 20

    # 冻结 Row 2 (滚动时表头常驻)
    try:
        ws.Activate()
        wb.Application.ActiveWindow.SplitColumn = 0
        wb.Application.ActiveWindow.SplitRow = 2
        wb.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass
    return ws


def _refresh_diagnostic_headers(ws):
    """Phase 4g Step 1: self-heal diagnostic header to 11 columns without touching row 3+."""
    headers = ["公司", "报表", "输出指标", "状态", "数据源",
               "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注", "FX_Rate"]
    widths = [14, 16, 30, 18, 18, 14, 42, 14, 10, 58, 12]

    try:
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).UnMerge()
    except Exception:
        pass

    ws.Cells(1, 1).Value = f"{ws.Name.replace('_', '')} (每次跑数后自动刷新)"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).Merge()
    title = ws.Range(ws.Cells(1, 1), ws.Cells(1, 11))
    title.Font.Name = "微软雅黑"
    title.Font.Size = 12
    title.Font.Bold = True
    title.Font.Color = rgb_long("FFFFFF")
    title.Interior.Color = rgb_long("4472C4")
    title.HorizontalAlignment = -4108
    title.VerticalAlignment = -4108

    for j, txt in enumerate(headers, start=1):
        c = ws.Cells(2, j)
        c.Value = txt
        c.Font.Name = "微软雅黑"
        c.Font.Size = 10
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
    for j, w in enumerate(widths, start=1):
        ws.Columns(j).ColumnWidth = w
    ws.Columns("A").NumberFormat = "@"
    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 20


def _make_cross_market_indicator_sheet(wb, name="跨市场_指标表"):
    """Phase 4g Step 2: cross-market indicator view sheet."""
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name

    headers = [("A", "指标类型", 18), ("B", "指标名称", 28), ("C", "英文指标名", 34)]
    for col, txt, width in headers:
        c = ws.Range(f"{col}1")
        c.Value = txt
        c.Font.Name = "微软雅黑"
        c.Font.Size = 11
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        ws.Columns(col).ColumnWidth = width

    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 20
    try:
        ws.Activate()
        wb.Application.ActiveWindow.SplitColumn = 3
        wb.Application.ActiveWindow.SplitRow = 2
        wb.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass
    return ws



def _make_fx_sheet(wb, name="汇率"):
    """Phase 4f Step 2: 汇率 sheet (8 列表头, 跨市场共享缓存)
      Row 1: 报告期/USDCNY期末/USDCNY期均/HKDCNY期末/HKDCNY期均/KRWCNY期末/KRWCNY期均/备注
      Row 2+: 由 VBA 模块_抓汇率 自动写; 用户可手填 override
      A 列文本格式 (防 yyyy-mm-dd 数字化), 冻结 A2
    """
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name

    headers = ["报告期", "USDCNY期末", "USDCNY期均",
               "HKDCNY期末", "HKDCNY期均",
               "KRWCNY期末", "KRWCNY期均", "备注/override"]
    widths = [14, 14, 14, 14, 14, 14, 14, 40]
    for j, (txt, w) in enumerate(zip(headers, widths), start=1):
        c = ws.Cells(1, j)
        c.Value = txt
        c.Font.Name = "微软雅黑"
        c.Font.Size = 11
        c.Font.Bold = True
        c.Font.Color = rgb_long("FFFFFF")
        c.Interior.Color = rgb_long("4472C4")
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        ws.Columns(j).ColumnWidth = w

    ws.Columns("A").NumberFormat = "@"
    ws.Rows(1).RowHeight = 22
    _write_fx_legend(ws)

    try:
        ws.Activate()
        wb.Application.ActiveWindow.SplitColumn = 0
        wb.Application.ActiveWindow.SplitRow = 1
        wb.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass
    return ws


def _write_fx_legend(ws_fx):
    """Phase 4i Step 3: 在汇率 sheet J10+ 追加数据源与取数逻辑说明区。"""
    start_row = 10
    try:
        if str(ws_fx.Range("A10").Value or "").strip() == "数据源与取数逻辑":
            ws_fx.Range("A10:H40").UnMerge()
            ws_fx.Range("A10:H40").Clear()
    except Exception:
        pass
    try:
        ws_fx.Range("J10:Q40").UnMerge()
    except Exception:
        pass
    ws_fx.Range("J10:Q40").Clear()

    header = ws_fx.Range("J10:Q10")
    header.Merge()
    header.Value = "汇率数据来源与折算说明"
    header.Font.Name = "微软雅黑"
    header.Font.Size = 14
    header.Font.Bold = True
    header.Font.Color = rgb_long("FFFFFF")
    header.Interior.Color = rgb_long("1F3864")
    header.HorizontalAlignment = -4108
    header.VerticalAlignment = -4108

    rows = [
        ("汇率数据来源", "工具会通过本地脚本从公开渠道获取 USD/HKD/KRW 对人民币的历史汇率;首次取数后会本地暂存 24 小时,降低重复抓数的网络消耗。"),
        ("  USDCNY", "美元兑人民币。"),
        ("  HKDCNY", "港币兑人民币。"),
        ("  KRWCNY", "韩元兑人民币。"),
        ("", ""),
        ("字段定义", ""),
        ("  期末汇率", "用于资产负债表,反映报告期最后一天的折算价值。"),
        ("  期间均值汇率", "用于利润表 / 现金流量表,反映整个报告期内损益和现金流的平均折算。"),
        ("", ""),
        ("应用规则", ""),
        ("  B6 = 原币", "各公司按报告币种显示,不做人民币折算。"),
        ("  B6 = 统一RMB", "非人民币口径会按报告期汇率折算成人民币,便于横向对比。"),
        ("  RMB / CNY", "本身就是人民币口径,无需折算。"),
        ("  手工调整汇率", "如需人工修正,可直接改 B-G 列;已生成的报表需要重新点击对应『一键 X 股』按钮后才会反映。"),
        ("", ""),
        ("本地暂存", ""),
        ("  24 小时复用", "同一报告期、同一币种在 24 小时内重复使用已取到的汇率,让重复跑数更快。"),
        ("  清空 HTTP 缓存", "样本池 Q9:Q11 按钮;清空后下次跑数会重新取汇率和财报数据。"),
        ("", ""),
        ("注意事项", ""),
        ("  报表已生成后", "如需重新换汇请重跑对应『一键 X 股』按钮。"),
        ("  B5 登录凭证", "汇率本身不需要 B5;B5 主要影响港股和部分中概美股备用数据源。"),
    ]

    for idx, (label, desc) in enumerate(rows, start=start_row + 1):
        a = ws_fx.Cells(idx, 10)
        b = ws_fx.Cells(idx, 11)
        a.Value = label
        b.Value = desc
        a.Font.Name = "微软雅黑"
        b.Font.Name = "微软雅黑"
        a.Font.Size = 10
        b.Font.Size = 11
        a.Font.Bold = True if label and not label.startswith("  ") else False
        b.Font.Bold = False
        a.HorizontalAlignment = -4131
        b.HorizontalAlignment = -4131
        a.VerticalAlignment = -4108
        b.VerticalAlignment = -4108
        a.WrapText = True
        b.WrapText = True
        if label and not label.startswith("  "):
            a.Interior.Color = rgb_long("D9E1F2")
            b.Interior.Color = rgb_long("F3F6FB")
        elif label:
            a.Interior.Color = rgb_long("FFFFFF")
            b.Interior.Color = rgb_long("FFFFFF")

    ws_fx.Columns("J").ColumnWidth = 18
    ws_fx.Columns("K").ColumnWidth = 60
    ws_fx.Range("J11:K32").WrapText = True
    for row in range(10, 33):
        ws_fx.Rows(row).RowHeight = 22


def install_currency_toggle_cell(ws_pool):
    """
    Phase 4f Step 2: 样本池 row 6 装『显示币种』toggle
      A6 = 标签『显示币种』(浅蓝)
      B6 = 默认 "原币" (浅黄), 数据验证下拉 "原币,统一RMB"; 不覆盖用户已设的值
    """
    label_cell = ws_pool.Range("A6")
    if not label_cell.Value:
        label_cell.Value = "显示币种"
    label_cell.Font.Name = "微软雅黑"
    label_cell.Font.Size = 10
    label_cell.Font.Bold = True
    label_cell.Interior.Color = rgb_long("B4C7E7")    # 浅蓝
    label_cell.HorizontalAlignment = -4108
    label_cell.VerticalAlignment = -4108

    val_cell = ws_pool.Range("B6")
    # 不覆盖用户已选的值; 仅在空时写默认
    if not val_cell.Value:
        val_cell.Value = "原币"
    val_cell.Font.Name = "微软雅黑"
    val_cell.Font.Size = 11
    val_cell.Font.Bold = True
    val_cell.Interior.Color = rgb_long("FFE699")    # 浅黄
    val_cell.HorizontalAlignment = -4108
    val_cell.VerticalAlignment = -4108

    # 数据验证: 下拉 原币/统一RMB
    try:
        val_cell.Validation.Delete()
    except Exception:
        pass
    try:
        val_cell.Validation.Add(
            Type=XL_VALIDATE_LIST,
            AlertStyle=XL_VALID_ALERT_STOP,
            Operator=1,
            Formula1="原币,统一RMB",
        )
        val_cell.Validation.IgnoreBlank = False
        val_cell.Validation.InCellDropdown = True
    except Exception as e:
        print(f"  ! B6 数据验证添加失败: {e}")

    # 加 comment 解释 toggle 含义
    try:
        if val_cell.Comment is None:
            val_cell.AddComment(
                "原币:按各公司报告币种显示。\n"
                "统一RMB:把非人民币口径折算成人民币,便于跨市场横向比较。\n"
                "资产负债表使用期末汇率;利润表和现金流量表使用期间平均汇率。"
            )
    except Exception:
        pass

    print("  + A6/B6 显示币种 toggle 已配置 (默认 '原币')")


def ensure_market_sheets(wb):
    """确保 A股/美股/港股/韩股报表 sheet 和诊断 sheet 存在。
    """
    wide_targets = [
        "A股_资产负债表", "A股_利润表", "A股_现金流量表", "A股_指标表",
        "美股_资产负债表", "美股_利润表", "美股_现金流量表", "美股_指标表",
        "港股_资产负债表", "港股_利润表", "港股_现金流量表", "港股_指标表",
        "韩股_资产负债表", "韩股_利润表", "韩股_现金流量表", "韩股_指标表",
    ]
    existing = {sh.Name for sh in wb.Sheets}
    for name in wide_targets:
        if name in existing:
            print(f"  ~ sheet 已存在: {name}")
        else:
            _make_wide_table_sheet(wb, name)
            print(f"  + sheet 新建: {name}")

    for diag_name in ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"):
        if diag_name in {sh.Name for sh in wb.Sheets}:
            ws_diag = wb.Sheets(diag_name)
            _refresh_diagnostic_headers(ws_diag)
            try:
                ws_diag.Visible = 0  # xlSheetHidden
            except Exception:
                pass
            print(f"  ~ sheet 已存在 (表头已升级到 11 列): {diag_name}")
        else:
            ws_diag = _make_diagnostic_sheet(wb, diag_name)
            try:
                ws_diag.Visible = 0  # xlSheetHidden, 用户可右键取消隐藏
            except Exception:
                pass
            print(f"  + sheet 新建: {diag_name}")


    # ---- Phase 4g Step 2: 跨市场指标合并视图 ----
    if "跨市场_指标表" in {sh.Name for sh in wb.Sheets}:
        print("  ~ sheet 已存在: 跨市场_指标表")
    else:
        _make_cross_market_indicator_sheet(wb, "跨市场_指标表")
        print("  + sheet 新建: 跨市场_指标表")

    # ---- Phase 4f Step 2: 汇率 sheet (跨市场共享缓存) ----
    if "汇率" in {sh.Name for sh in wb.Sheets}:
        _write_fx_legend(wb.Sheets("汇率"))
        print("  ~ sheet 已存在: 汇率")
    else:
        _make_fx_sheet(wb, "汇率")
        print("  + sheet 新建: 汇率")



def update_intro_sheet(wb):
    """刷新使用说明页: 商务化封面、目录、7 个 section 和表格化说明。"""
    try:
        ws = wb.Sheets("使用说明")
    except Exception:
        ws = wb.Worksheets.Add(Before=wb.Sheets(1))
        ws.Name = "使用说明"

    try:
        ws.Cells.UnMerge()
    except Exception:
        pass
    ws.Cells.Clear()
    ws.Cells.Font.Name = "微软雅黑"
    ws.Cells.Font.Size = 11
    ws.Rows.RowHeight = 18

    widths = {"A": 4, "B": 18, "C": 42, "D": 15, "E": 18, "F": 18, "G": 18, "H": 18}
    for col, width in widths.items():
        ws.Columns(col).ColumnWidth = width

    deep_blue = "1F3864"
    mid_blue = "4472C4"
    light_blue = "D9E1F2"
    white = "FFFFFF"

    def style_range(rng, size=11, bold=False, font_hex="000000", fill_hex=None, align=-4131):
        rng.Font.Name = "微软雅黑"
        rng.Font.Size = size
        rng.Font.Bold = bold
        rng.Font.Color = rgb_long(font_hex)
        if fill_hex:
            rng.Interior.Color = rgb_long(fill_hex)
        rng.HorizontalAlignment = align
        rng.VerticalAlignment = -4108
        rng.WrapText = True

    def merge_write(addr, value, size=11, bold=False, font_hex="000000", fill_hex=None, align=-4131):
        rng = ws.Range(addr)
        if addr.split(":")[0] != addr.split(":")[-1]:
            rng.Merge()
        rng.Value = value
        style_range(rng, size=size, bold=bold, font_hex=font_hex, fill_hex=fill_hex, align=align)
        return rng

    def apply_borders(addr):
        rng = ws.Range(addr)
        for border_idx in (7, 8, 9, 10, 11, 12):
            try:
                b = rng.Borders(border_idx)
                b.LineStyle = 1
                b.Weight = 2
                b.Color = rgb_long("BFBFBF")
            except Exception:
                pass

    def section(row, title):
        merge_write(f"A{row}:H{row}", title, size=14, bold=True, font_hex=white, fill_hex=mid_blue, align=-4108)
        ws.Rows(row).RowHeight = 24
        return row + 2

    def paragraph(row, text, height=34):
        merge_write(f"B{row}:H{row}", text, size=11, fill_hex="FFFFFF")
        ws.Rows(row).RowHeight = height
        return row + 1

    def table(row, headers, data, widths_override=None):
        for offset, header in enumerate(headers):
            cell = ws.Cells(row, 2 + offset)
            cell.Value = header
            style_range(cell, size=10, bold=True, font_hex="1F4E79", fill_hex=light_blue, align=-4108)
        for r_offset, record in enumerate(data, start=1):
            for c_offset, value in enumerate(record):
                cell = ws.Cells(row + r_offset, 2 + c_offset)
                cell.Value = value
                style_range(cell, size=10, fill_hex="FFFFFF", align=-4131)
        end_col = chr(ord("B") + len(headers) - 1)
        apply_borders(f"B{row}:{end_col}{row + len(data)}")
        for rr in range(row, row + len(data) + 1):
            ws.Rows(rr).RowHeight = 20
        if widths_override:
            for col, width in widths_override.items():
                ws.Columns(col).ColumnWidth = width
        return row + len(data) + 2

    # Cover
    merge_write("A1:H1", "上市公司财务数据查询", size=24, bold=True, font_hex=white, fill_hex=deep_blue, align=-4108)
    merge_write("A2:E2", "Multi-Market Financial Data Tool", size=14, bold=False, font_hex=white, fill_hex=deep_blue, align=-4108)
    merge_write("F2:H2", "Phase 4j", size=10, bold=True, font_hex=white, fill_hex=deep_blue, align=-4108)
    merge_write("A3:E3", "面向财务 / 审计专业人士的本地 Excel 同业对标工具", size=10, font_hex=white, fill_hex=deep_blue, align=-4108)
    merge_write("F3:H3", "修订日期: 2026-05-04", size=10, font_hex=white, fill_hex=deep_blue, align=-4108)
    ws.Rows(1).RowHeight = 36
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 18

    # TOC
    merge_write("A5:H5", "目录", size=13, bold=True, font_hex="1F4E79", fill_hex=light_blue, align=-4108)
    toc = [
        ("§ 1", "项目概览", "R14"),
        ("§ 2", "快速开始", "R21"),
        ("§ 3", "输出 sheet 说明", "R32"),
        ("§ 4", "数据源声明", "R43"),
        ("§ 5", "汇率换算说明", "R51"),
        ("§ 6", "常见问题", "R61"),
        ("§ 7", "版本历史", "R75"),
    ]
    for idx, (sec, title, target) in enumerate(toc, start=6):
        ws.Cells(idx, 2).Value = sec
        ws.Cells(idx, 3).Value = title
        ws.Cells(idx, 4).Value = target
        style_range(ws.Range(f"B{idx}:D{idx}"), size=10, fill_hex="FFFFFF")
    apply_borders("B6:D12")

    row = 14
    row = section(row, "§ 1 项目概览")
    row = paragraph(
        row,
        "本工具用于对比 A 股、港股、美股、韩股上市公司的财务表现。用户在样本池录入公司后,工作簿会自动整理资产负债表、利润表、现金流量表和关键指标,并自动生成跨市场指标对比视图。它适合做同业初筛、审计分析底稿准备、跨市场口径差异识别和管理层问询前的数据准备。",
        height=42,
    )
    row = table(
        row,
        ("市场", "状态", "数据源"),
        [
            ("A股", "已支持", "新浪财经"),
            ("美股", "已支持", "美国上市公司披露数据;必要时使用中概股备用数据源"),
            ("港股", "已支持", "雪球港股财务数据"),
            ("韩股", "已支持", "公开韩股财务报表页面"),
        ],
    )

    row = section(row, "§ 2 快速开始")
    quick_steps = [
        ("1", "在『样本池』A2 选择年度,A4 选择季度。年度留空时,工具尽量使用可取得的最新报告期。"),
        ("2", "如需港股或部分中概美股,在 B5 填入雪球登录凭证;B6 选择『原币』或『统一RMB』。"),
        ("3", "第 11 行起按市场录入公司:A:B=A股,E:F=美股,I:J=港股,M:N=韩股。"),
        ("4", "点击对应市场的一键按钮。抓数完成后,对应市场的 4 张正式报表会自动显示。"),
        ("5", "点击『一键全抓 4 市场』后,跨市场_指标表会自动刷新。"),
        ("6", "跑数后『跨市场_指标表』自动展示 18 项标准指标的 4 市场对比。"),
    ]
    row = table(row, ("步骤", "操作"), quick_steps, widths_override={"C": 58})

    row = section(row, "§ 3 输出 sheet 说明")
    output_rows = [
        ("样本池", "录入公司、选择年度和显示币种,并通过按钮完成抓数和跨市场对比。"),
        ("跨市场_指标表", "18 项标准指标的跨市场合并视图,用于快速比较盈利、周转、现金流和增长。"),
        ("A股 / 美股 / 港股 / 韩股 16 张分市场表", "各市场原始写表结果,保留市场披露口径。"),
        ("抓取诊断", "默认隐藏;仅排查数据来源、单位或字段缺口时使用。日常分析无需打开。"),
        ("汇率", "记录外币报表折算成人民币时使用的报告期汇率。"),
    ]
    row = table(row, ("Tab", "用途"), output_rows, widths_override={"B": 30, "C": 64})

    row = section(row, "§ 4 数据源声明")
    source_rows = [
        ("A股", "新浪财经公开财报数据。"),
        ("美股", "美国上市公司披露数据为主;中概股缺口按既有规则使用备用数据源。"),
        ("港股", "雪球港股财务数据;通常需要有效登录凭证。"),
        ("韩股", "公开韩股财务报表页面;不需要雪球登录凭证。"),
        ("汇率", "公开历史汇率数据,本地记录到『汇率』sheet。"),
    ]
    row = table(row, ("对象", "说明"), source_rows, widths_override={"B": 18, "C": 70})
    row = paragraph(row, "登录凭证只保存在本地工作簿中。过期后请重新复制到样本池 B5。", height=24)

    row = section(row, "§ 5 汇率换算说明")
    fx_rows = [
        ("B6 = 原币", "各市场按报告币种或市场默认单位显示,不做 RMB 换算。"),
        ("B6 = 统一RMB", "非人民币报告币种按报告期汇率折算:资产负债表用期末汇率,利润表 / 现金流量表用期间均值。"),
        ("RMB / CNY", "报告币种为人民币时无需折算。"),
        ("手工改汇率", "已写表公式不会自动反向刷新;需要重跑对应写表或合表按钮。"),
    ]
    row = table(row, ("规则", "说明"), fx_rows, widths_override={"B": 20, "C": 68})

    row = section(row, "§ 6 常见问题")
    faq_rows = [
        ("Q1: 港股或中概美股取数失败怎么办?", "先检查 B5 登录凭证是否过期。重新登录后复制新的 xq_a_token,再重跑对应市场。"),
        ("Q2: 切换 B6 后数字变了是否正常?", "正常。『统一RMB』会把外币报告币种折算成人民币,用于跨市场可比。"),
        ("Q3: 跨市场指标表何时刷新?", "点击『一键全抓 4 市场』后自动刷新;也可以单独运行 BuildCrossMarketIndicatorSheet。"),
        ("Q4: 为什么只保留跨市场指标表?", "BS/IS/CF 保留在分市场表中按原披露口径查看;跨市场对比聚焦 18 项标准指标。"),
        ("Q5: 老版样本池升级会丢数据吗?", "不会。第 11 行起的公司代码和简称会保留;安装脚本只重画上方配置区和按钮区。"),
    ]
    row = table(row, ("问题", "处理建议"), faq_rows, widths_override={"B": 30, "C": 64})

    row = section(row, "§ 7 版本历史")
    history_rows = [
        ("Phase 4b", "2026-05", "A股宽表基础模板与 VBA 注入流程。"),
        ("Phase 4c-4d", "2026-05", "扩展美股、港股、韩股数据源。"),
        ("Phase 4f", "2026-05", "RMB toggle 与本地汇率缓存。"),
        ("Phase 4g", "2026-05", "跨市场指标合表与分市场 tab 显隐。"),
        ("Phase 4h", "2026-05", "跨市场 BS/IS/CF 合表、实时币种切换、本地暂存和中概备用数据源。"),
        ("Phase 4i", "2026-05", "样本池布局、使用说明和汇率说明区 UX 抛光。"),
        ("Phase 4j", "2026-05", "审计视角说明、重点 sheet 视觉标准化和样本池布局重整。"),
    ]
    row = table(row, ("版本", "日期", "主线交付"), history_rows, widths_override={"B": 16, "C": 14, "D": 72})

    ws.Range(f"A1:H{row + 2}").WrapText = True
    try:
        ws.Activate()
        wb.Application.ActiveWindow.SplitColumn = 0
        wb.Application.ActiveWindow.SplitRow = 4
        wb.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    print("  + 使用说明 已刷新 (商务化版)")


def reorder_report_sheets(wb):
    """固定工作表 Tab 顺序;诊断 sheet 排序后保持 xlSheetHidden。"""
    desired_order = [
        "使用说明", "样本池", "跨市场_指标表",
        "A股_资产负债表", "A股_利润表", "A股_现金流量表", "A股_指标表",
        "美股_资产负债表", "美股_利润表", "美股_现金流量表", "美股_指标表",
        "美股_抓取诊断",
        "港股_资产负债表", "港股_利润表", "港股_现金流量表", "港股_指标表",
        "港股_抓取诊断",
        "韩股_资产负债表", "韩股_利润表", "韩股_现金流量表", "韩股_指标表",
        "韩股_抓取诊断",
        "汇率",   # ← Phase 4f Step 2 新增 (跨市场共享 FX 缓存)
    ]
    diagnostic_names = {"美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"}
    for name in diagnostic_names:
        try:
            sh = wb.Sheets(name)
            if sh.Visible != -1:
                sh.Visible = -1  # 临时显示,兼容 Excel 对 hidden sheet Move 的限制
        except Exception:
            pass

    pos = 1
    for name in desired_order:
        try:
            sh = wb.Sheets(name)
        except Exception:
            continue
        try:
            sh.Move(Before=wb.Sheets(pos))
            pos += 1
        except Exception as e:
            print(f"  ! sheet 顺序调整失败 {name}: {e}")
    for name in diagnostic_names:
        try:
            wb.Sheets(name).Visible = 0
        except Exception:
            pass
    print("  + sheet Tab 顺序已调整")


def colorize_sheet_tabs(wb):
    """Phase 4i.1: 按市场给 sheet tab 染色,共享 sheet 保持默认无色。"""
    rules = [
        ("A股_", PRIMARY_FILL),
        ("美股_", US_FILL),
        ("港股_", HK_FILL),
        ("韩股_", KR_FILL),
        ("跨市场_", PRIMARY_FILL),
    ]
    for ws in wb.Worksheets:
        matched = False
        for prefix, fill_hex in rules:
            if ws.Name.startswith(prefix):
                ws.Tab.Color = rgb_long(fill_hex)
                matched = True
                break
        if not matched:
            try:
                ws.Tab.ColorIndex = -4142  # xlColorIndexNone
            except Exception:
                pass
    print("  + sheet Tab 颜色已按市场刷新")


def install_buttons(ws_pool):
    """
    Phase 4e: 顶部 4 个市场一键 + Q1 全局一键;16 个单表按钮折叠到 Row 30+。
    """
    # 删旧按钮: 当前 BUTTONS 列表内的 + 已废弃的
    target_names = {b[0] for b in BUTTONS} | set(DECOMMISSIONED_BUTTONS)
    # 必须遍历 Shapes 副本, 否则 Delete 会改变集合
    shape_names = [sh.Name for sh in ws_pool.Shapes]
    for name in shape_names:
        if name in target_names:
            ws_pool.Shapes(name).Delete()
            tag = "decommissioned" if name in DECOMMISSIONED_BUTTONS else "existing"
            print(f"  - removed {tag} button: {name}")

    for r in range(30, 34):
        try:
            ws_pool.Rows(r).RowHeight = 24
        except Exception:
            pass

    for name, caption, macro, addr, fill_hex, font_hex, font_size, is_primary in BUTTONS:
        rng = ws_pool.Range(addr)
        left = rng.Left + 1
        top = rng.Top + 2
        width = max(20, rng.Width - 2)
        height = max(18, rng.Height - 4)
        shape = ws_pool.Shapes.AddShape(
            MSO_SHAPE_ROUNDED_RECT, left, top, width, height
        )
        shape.Name = name
        shape.Fill.Visible = True
        shape.Fill.ForeColor.RGB = rgb_long(fill_hex)
        shape.Line.Visible = False

        tf = shape.TextFrame2
        tf.MarginLeft = 4
        tf.MarginRight = 4
        tf.MarginTop = 2
        tf.MarginBottom = 2
        tf.HorizontalAnchor = MSO_ANCHOR_CENTER
        tf.VerticalAnchor = MSO_ANCHOR_MIDDLE
        tf.WordWrap = -1   # msoTrue
        tf.TextRange.Text = caption
        tf.TextRange.Font.Size = font_size
        tf.TextRange.Font.Bold = -1
        tf.TextRange.Font.Fill.ForeColor.RGB = rgb_long(font_hex)
        # 设字体名 (中文用微软雅黑)
        try:
            tf.TextRange.Font.Name = "微软雅黑"
            tf.TextRange.Font.NameFarEast = "微软雅黑"
        except Exception:
            pass
        # 段落水平居中 (msoAlignCenter = 2)
        try:
            tf.TextRange.ParagraphFormat.Alignment = 2
        except Exception:
            pass

        if name == "BtnClearCache":
            try:
                shape.AlternativeText = (
                    "清除本地暂存的抓数结果。"
                    "下次抓数会重新从公开数据来源取数,适合强制刷新或排查数据陈旧问题。日常无需点击。"
                )
            except Exception:
                pass

        shape.OnAction = macro
        print(f"  + button: {name:15s} [{caption}] @ {addr} → {macro}")

def parse_bas(path: Path) -> tuple[str, str]:
    """Read a .bas file. Return (module_name, body_without_attribute_line)."""
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    name = None
    body_start = 0
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith("Attribute VB_Name"):
            name = s.split("=", 1)[1].strip().strip('"')
            body_start = i + 1
            break
    if name is None:
        name = path.stem
        body_start = 0
    body = "\n".join(lines[body_start:]).lstrip("\n")
    return name, body


def main():
    # 决定打开哪个文件:
    #   - 如果 .xlsx 存在 (build_template.py 刚跑过), 优先用它, 转成 xlsm
    #     (会覆盖任何残留的旧 .xlsm)
    #   - 否则 .xlsm 存在, 直接打开补 VBA
    open_path = None
    save_as_xlsm = False
    legacy_source = False
    if XLSX.exists():
        open_path = XLSX
        save_as_xlsm = True
        if XLSM.exists():
            print(f"Note: 会覆盖旧的 {XLSM.name}")
        print(f"Opening {XLSX.name}, will save as {XLSM.name}")
    elif XLSM.exists():
        open_path = XLSM
        print(f"Opening existing {XLSM.name}")
    elif LEGACY_XLSX.exists():
        open_path = LEGACY_XLSX
        save_as_xlsm = True
        legacy_source = True
        print(f"Opening legacy {LEGACY_XLSX.name}, will save as {XLSM.name}")
    elif LEGACY_XLSM.exists():
        open_path = LEGACY_XLSM
        save_as_xlsm = True
        legacy_source = True
        print(f"Opening legacy {LEGACY_XLSM.name}, will save as {XLSM.name}")
    else:
        print(f"FATAL: 未找到 {XLSX.name} / {XLSM.name} 或旧版 V3 工作簿")
        print(f"先跑 `py tools/build_template.py`")
        sys.exit(1)

    if not MODULES_DIR.exists():
        print(f"FATAL: {MODULES_DIR} not found.")
        sys.exit(1)
    bas_files = sorted(MODULES_DIR.glob("*.bas"))
    if not bas_files:
        print(f"No .bas files in {MODULES_DIR}")
        sys.exit(1)

    print(f"Found {len(bas_files)} .bas files:")
    for p in bas_files:
        print(f"  {p.name}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(open_path))
        try:
            try:
                vbproject = wb.VBProject
            except Exception as e:
                print("\nFATAL: 无法访问 VBProject。请在 Excel 启用:")
                print("  文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置")
                print("  → 勾选『信任对 VBA 工程对象模型的访问』")
                print(f"  ({e})")
                sys.exit(2)

            # ---- Sheet 重命名 (累积所有迁移规则) ----
            existing_sheets = {sh.Name for sh in wb.Sheets}
            for old_name, new_name in SHEET_RENAMES.items():
                if old_name in existing_sheets and new_name not in existing_sheets:
                    wb.Sheets(old_name).Name = new_name
                    print(f"  ~ renamed sheet: {old_name} → {new_name}")
                    existing_sheets.discard(old_name)
                    existing_sheets.add(new_name)
                elif old_name in existing_sheets and new_name in existing_sheets:
                    print(f"  ! both {old_name} and {new_name} exist, skipping rename")

            # ---- 删除已废弃 sheet ----
            for dead in DECOMMISSIONED_SHEETS:
                if dead in existing_sheets:
                    try:
                        wb.Application.DisplayAlerts = False
                        wb.Sheets(dead).Delete()
                        print(f"  ~ sheet 已删除: {dead}")
                    except Exception as e:
                        print(f"  ! 删除 {dead} 失败: {e}")
                    finally:
                        wb.Application.DisplayAlerts = True

            # ---- Phase 4b: 加 Microsoft Scripting Runtime 引用 (JsonConverter 需要) ----
            #   GUID: {420B2830-E718-11CF-893D-00A0C9054228} = Microsoft Scripting Runtime
            try:
                vbproject.References.AddFromGuid(
                    "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0
                )
                print("  + Reference: Microsoft Scripting Runtime")
            except Exception as e:
                msg = str(e).lower()
                if (
                    "already" in msg
                    or "name conflicts" in msg
                    or "名称与已存在" in msg
                    or "冲突" in msg
                    or "0x800a03ec" in msg
                ):
                    print("  ~ Reference 已存在: Microsoft Scripting Runtime")
                else:
                    print(f"  ! 无法添加 Scripting Runtime 引用: {e}")
                    print(f"    JsonConverter 可能编译失败 (美股抓数会报错)")

            # ---- 移除废弃模块 (即使本地 modules/ 已删, 老 xlsm 里 VBComponent 可能还在) ----
            for old_name in DECOMMISSIONED_MODULES:
                try:
                    old_comp = vbproject.VBComponents(old_name)
                    vbproject.VBComponents.Remove(old_comp)
                    print(f"  x decommissioned module: {old_name}")
                except Exception:
                    pass

            for path in bas_files:
                name, body = parse_bas(path)
                try:
                    existing = vbproject.VBComponents(name)
                    vbproject.VBComponents.Remove(existing)
                    print(f"  - removed existing: {name}")
                except Exception:
                    pass
                comp = vbproject.VBComponents.Add(VBE_CT_STDMODULE)
                comp.Name = name
                if body:
                    comp.CodeModule.AddFromString(body)
                print(f"  + installed: {name} ({len(body)} chars)")

            # ---- Phase 3 + 4: 季度选择器 + 市场列 + 市场 sheet + 圆角按钮 ----
            try:
                ws_pool = wb.Sheets("样本池")
                migrate_old_sample_pool(ws_pool)       # 旧 A:C 混合布局 → 4 市场分栏
                layout_sample_pool(ws_pool)            # Phase 4e 样本池布局
                ensure_market_sheets(wb)
                update_intro_sheet(wb)
                reorder_report_sheets(wb)
                colorize_sheet_tabs(wb)
                install_buttons(ws_pool)
            except Exception as e:
                print(f"! Failed to install quarter / market / sheets / buttons: {e}")

            if save_as_xlsm:
                # SaveAs xlsm format, then clean up the xlsx
                wb.SaveAs(str(XLSM), FileFormat=XL_FILEFORMAT_XLSM)
                print(f"\n+ Saved as {XLSM.name}")
            else:
                wb.Save()
                print(f"\n+ Saved {XLSM.name}")
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    if save_as_xlsm and XLSX.exists():
        try:
            XLSX.unlink()
            print(f"+ Removed leftover {XLSX.name}")
        except Exception as e:
            print(f"! Could not remove {XLSX.name}: {e}  (可手动删除)")
    if save_as_xlsm and legacy_source and LEGACY_XLSX.exists():
        try:
            LEGACY_XLSX.unlink()
            print(f"+ Removed leftover legacy {LEGACY_XLSX.name}")
        except Exception as e:
            print(f"! Could not remove {LEGACY_XLSX.name}: {e}  (可手动删除)")


if __name__ == "__main__":
    main()

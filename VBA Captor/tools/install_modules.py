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
COVER_NAVY = "1F3864"     # Phase 4j.2: 样本池封面深蓝
A_LIGHT_FILL = "D9E1F2"
US_LIGHT_FILL = "FCE4E4"
HK_LIGHT_FILL = "EAF4E3"
KR_LIGHT_FILL = "EEE8F7"
SECTION_LABEL_FILL = "F2F2F2"

US_FILL = "C00000"        # 深红 — 美股按钮区分
US_FG = "FFFFFF"
HK_FILL = "548235"        # 深绿 — 港股按钮区分
HK_FG = "FFFFFF"
KR_FILL = "7030A0"        # 深紫 — 韩股按钮区分
KR_FG = "FFFFFF"

BUTTONS = [
    ("BtnRunAll",       "一键全抓 4 市场",      "模块_总入口.一键全抓",            "N2:Q3",  COVER_NAVY,     PRIMARY_FG,   13, True),
    ("BtnBuildCrossInd", "一键抓取跨市场指标表", "模块_总入口.一键跨市场指标表",    "N4:Q5",  A_LIGHT_FILL,   SECONDARY_FG, 10, True),
    ("BtnHideAll",      "显示/隐藏 所有市场数据", "模块_总入口.切换所有分市场tabs", "N6:Q7",  A_LIGHT_FILL,   SECONDARY_FG, 11, True),
    ("BtnClearCache",   "清空 HTTP 缓存",       "模块_工具函数.ClearLocalCache",   "N12:Q13", A_LIGHT_FILL,   SECONDARY_FG, 11, True),
    ("BtnRunA",         "一键 A 股",           "模块_总入口.一键A股",             "A9:C10", COVER_NAVY,     PRIMARY_FG,   14, True),
    ("BtnRunUS",        "一键 美股",           "模块_总入口.一键美股",            "D9:F10", US_FILL,        US_FG,        14, True),
    ("BtnRunHK",        "一键 港股",           "模块_总入口.一键港股",            "G9:I10", HK_FILL,        HK_FG,        14, True),
    ("BtnRunKR",        "一键 韩股",           "模块_总入口.一键韩股",            "J9:M10", KR_FILL,        KR_FG,        14, True),
    ("BtnHideA",        "显示/隐藏 A股数据",     "模块_总入口.切换A股tabs",          "A11:C12", A_LIGHT_FILL,   SECONDARY_FG, 11, False),
    ("BtnHideUS",       "显示/隐藏 美股数据",    "模块_总入口.切换美股tabs",         "D11:F12", US_LIGHT_FILL,  "9C0006",    11, False),
    ("BtnHideHK",       "显示/隐藏 港股数据",    "模块_总入口.切换港股tabs",         "G11:I12", HK_LIGHT_FILL,  "375623",    11, False),
    ("BtnHideKR",       "显示/隐藏 韩股数据",    "模块_总入口.切换韩股tabs",         "J11:M12", KR_LIGHT_FILL,  "5B2B82",    11, False),
]

# 已废弃: install 时从当前 xlsm 主动移除 (即使 modules/ 下仍有遗留也清掉)
DECOMMISSIONED_MODULES = ["模块_抓基本资料"]
DECOMMISSIONED_BUTTONS = [
    "BtnRunInfo",
    "BtnBuildCrossAll",
    "BtnHideCrossMarket",
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


def capture_sample_pool_companies(ws_pool):
    """Collect existing four-market sample rows before repainting the UI."""
    source_cols = {
        "A": [(1, 2)],
        "US": [(4, 5), (5, 6)],
        "HK": [(7, 8), (9, 10)],
        "KR": [(10, 11), (13, 14)],
    }
    by_market = {market: [] for market in source_cols}
    seen = {market: set() for market in source_cols}
    skip_tokens = {"代码", "简称", "一键 A 股", "一键 美股", "一键 港股", "一键 韩股"}

    for start_row in (14, 13, 11, 10):
        for row in range(start_row, 1001):
            for market, pairs in source_cols.items():
                for code_col, name_col in pairs:
                    code = _cell_text(ws_pool.Cells(row, code_col))
                    if not code or code in skip_tokens or "显示/隐藏" in code:
                        continue
                    key = code.upper()
                    if key in seen[market]:
                        break
                    if market in {"A", "HK", "KR"} and not code.isdigit():
                        continue
                    if market == "US" and code != code.upper():
                        continue
                    name = _cell_text(ws_pool.Cells(row, name_col))
                    by_market[market].append((code, name))
                    seen[market].add(key)
                    break

    return by_market


def restore_sample_pool_companies(ws_pool, by_market, start_row=14):
    col_map = {"A": (1, 2, 3), "US": (4, 5, 6), "HK": (7, 8, 9), "KR": (10, 11, 13)}
    for market, (code_col, name_col, end_col) in col_map.items():
        ws_pool.Range(ws_pool.Cells(start_row, code_col), ws_pool.Cells(1000, end_col)).ClearContents()
        for offset, (code, name) in enumerate(by_market.get(market, [])):
            row = start_row + offset
            ws_pool.Cells(row, code_col).NumberFormat = "@"
            ws_pool.Cells(row, code_col).Value = code
            ws_pool.Cells(row, name_col).Value = name


def _style_merged_range(rng, value, fill_hex, font_hex="000000", font_size=11, bold=False, align=-4108):
    rng.Merge()
    rng.Value = value
    rng.Font.Name = "微软雅黑"
    rng.Font.Size = font_size
    rng.Font.Bold = bool(bold)
    rng.Font.Color = rgb_long(font_hex)
    rng.Interior.Color = rgb_long(fill_hex)
    rng.HorizontalAlignment = align
    rng.VerticalAlignment = -4108
    rng.WrapText = True


def _apply_range_border(rng, color_hex="D9D9D9", weight=2, inside=False):
    border_ids = (7, 8, 9, 10, 11, 12) if inside else (7, 8, 9, 10)
    for border_idx in border_ids:
        try:
            border = rng.Borders(border_idx)
            border.LineStyle = 1
            border.Weight = weight
            border.Color = rgb_long(color_hex)
        except Exception:
            pass


def _apply_card_brand(ws_pool, left_col, right_col, title, brand_hex, light_hex, code_col, name_start_col, name_end_col):
    _style_merged_range(ws_pool.Range(ws_pool.Cells(7, left_col), ws_pool.Cells(7, right_col)), title, brand_hex, "FFFFFF", 12, True)
    band = ws_pool.Range(ws_pool.Cells(8, left_col), ws_pool.Cells(8, right_col))
    band.Merge()
    band.Interior.Color = rgb_long("FFFFFF")  # Phase 4j.3: 白色窄带,作为 header 与按钮之间的呼吸间距(去掉 Phase 4j.2 的 brand 色重复带)

    _style_merged_range(ws_pool.Range(ws_pool.Cells(9, left_col), ws_pool.Cells(10, right_col)), "", brand_hex, "FFFFFF", 14, True)
    _style_merged_range(ws_pool.Range(ws_pool.Cells(11, left_col), ws_pool.Cells(12, right_col)), "", light_hex, SECONDARY_FG, 11, True)

    code_header = ws_pool.Cells(13, code_col)
    code_header.Value = "代码"
    code_header.Font.Name = "微软雅黑"
    code_header.Font.Size = 10
    code_header.Font.Bold = True
    code_header.Font.Color = rgb_long("FFFFFF")
    code_header.Interior.Color = rgb_long(brand_hex)
    code_header.HorizontalAlignment = -4108
    code_header.VerticalAlignment = -4108
    name_header = ws_pool.Range(ws_pool.Cells(13, name_start_col), ws_pool.Cells(13, name_end_col))
    _style_merged_range(name_header, "简称", brand_hex, "FFFFFF", 10, True)

    data_area = ws_pool.Range(ws_pool.Cells(14, left_col), ws_pool.Cells(50, right_col))
    data_area.Interior.Color = rgb_long("FFFFFF")
    data_area.Font.Name = "微软雅黑"
    data_area.Font.Size = 10
    data_area.HorizontalAlignment = -4108
    data_area.VerticalAlignment = -4108
    _apply_range_border(ws_pool.Range(ws_pool.Cells(7, left_col), ws_pool.Cells(50, right_col)), "D9D9D9", 2, True)


def layout_sample_pool(ws_pool):
    """重画样本池配置区、市场卡片和数据区格式,保留已有公司数据。"""
    year_value = ws_pool.Range("E3").Value or ws_pool.Range("A2").Value or 2025
    quarter_value = ws_pool.Range("E4").Value or ws_pool.Range("A4").Value or "全部"
    cookie_value = ws_pool.Range("E5").Value or ws_pool.Range("B5").Value or ""
    currency_value = ws_pool.Range("E6").Value or ws_pool.Range("B6").Value or "原币"

    migrate_phase4g_sample_rows(ws_pool)
    companies = capture_sample_pool_companies(ws_pool)

    try:
        ws_pool.Range("A1:V50").UnMerge()
        ws_pool.Range("W1:Y8").UnMerge()
    except Exception:
        pass
    ws_pool.Range("A1:V50").Clear()
    ws_pool.Range("W1:Y8").Clear()

    widths = {
        "A": 11, "B": 16, "C": 1, "D": 11,
        "E": 16, "F": 1, "G": 11, "H": 16,
        "I": 1, "J": 11, "K": 16, "L": 1,
        "M": 1, "N": 8, "O": 8, "P": 8,
        "Q": 8, "R": 4, "S": 14, "T": 14,
        "U": 14, "V": 14,
    }
    for col, width in widths.items():
        ws_pool.Columns(col).ColumnWidth = width

    # Phase 4j.4: 彻底 hide 5 个 padding 列, 让 4 张卡片视觉上紧贴
    # (Phase 4j.3 缩到 width=1 仍可见空列, 改 Hidden=True 才消失)
    for hide_col in ("C", "F", "I", "L", "M"):
        try:
            ws_pool.Columns(hide_col).Hidden = True
        except Exception as e:
            print(f"  ! 隐藏列 {hide_col} 失败: {e}")

    label_fill = rgb_long("F7FBFF")
    value_fill = rgb_long("FFF8DF")
    border_blue = "7EA6D9"

    _style_merged_range(ws_pool.Range("A2:M2"), "参数设置", COVER_NAVY, "FFFFFF", 12, True, -4131)
    try:
        ws_pool.Range("A2:M2").IndentLevel = 1
    except Exception:
        pass

    config_rows = [
        (3, "年份（留空=取最新）", year_value),
        (4, "Q1/Q2/Q3/Q4 或 全部", quarter_value),
        (5, "雪球 Cookie", cookie_value),
        (6, "显示币种", currency_value),
    ]
    for row, label, value in config_rows:
        label_rng = ws_pool.Range(f"A{row}:D{row}")
        value_rng = ws_pool.Range(f"E{row}:M{row}")
        label_rng.Merge()
        value_rng.Merge()
        label_rng.Value = label
        value_rng.Value = value
        label_rng.Font.Name = "微软雅黑"
        label_rng.Font.Size = 11
        label_rng.Font.Bold = True
        label_rng.Font.Color = rgb_long(COVER_NAVY)
        label_rng.Interior.Color = label_fill
        label_rng.HorizontalAlignment = -4131
        label_rng.VerticalAlignment = -4108
        value_rng.Font.Name = "微软雅黑"
        value_rng.Font.Size = 11
        value_rng.Interior.Color = value_fill
        value_rng.HorizontalAlignment = -4131
        value_rng.VerticalAlignment = -4108
        value_rng.WrapText = True

    try:
        ws_pool.Range("E4").Validation.Delete()
        ws_pool.Range("E4").Validation.Add(
            Type=XL_VALIDATE_LIST,
            AlertStyle=XL_VALID_ALERT_STOP,
            Operator=1,
            Formula1="全部,Q1,Q2,Q3,Q4",
        )
        ws_pool.Range("E4").Validation.IgnoreBlank = False
        ws_pool.Range("E4").Validation.InCellDropdown = True
    except Exception as e:
        print(f"  ! E4 数据验证添加失败: {e}")
    try:
        ws_pool.Range("E6").Validation.Delete()
        ws_pool.Range("E6").Validation.Add(
            Type=XL_VALIDATE_LIST,
            AlertStyle=XL_VALID_ALERT_STOP,
            Operator=1,
            Formula1="原币,统一RMB",
        )
        ws_pool.Range("E6").Validation.IgnoreBlank = False
        ws_pool.Range("E6").Validation.InCellDropdown = True
    except Exception as e:
        print(f"  ! E6 数据验证添加失败: {e}")

    for addr in ("E5", "E6"):
        try:
            if ws_pool.Range(addr).Comment is not None:
                ws_pool.Range(addr).Comment.Delete()
        except Exception:
            pass
    try:
        ws_pool.Range("E5").AddComment(
            "雪球登录凭证,用于港股和部分中概美股备用数据源。\n"
            "如果港股或中概美股取数失败,请重新登录雪球后复制新的 xq_a_token。"
        )
    except Exception:
        pass

    _apply_range_border(ws_pool.Range("A2:M6"), border_blue, 2, True)

    _apply_card_brand(ws_pool, 1, 3, "A股(新浪)", COVER_NAVY, A_LIGHT_FILL, 1, 2, 3)
    _apply_card_brand(ws_pool, 4, 6, "美股(EDGAR+雪球)", US_FILL, US_LIGHT_FILL, 4, 5, 6)
    _apply_card_brand(ws_pool, 7, 9, "港股(雪球 HK)", HK_FILL, HK_LIGHT_FILL, 7, 8, 9)
    _apply_card_brand(ws_pool, 10, 13, "韩股(stockanalysis)", KR_FILL, KR_LIGHT_FILL, 10, 11, 13)

    placeholders = [
        ("A9:C10", "一键 A 股", COVER_NAVY, "FFFFFF", 14),
        ("D9:F10", "一键 美股", US_FILL, "FFFFFF", 14),
        ("G9:I10", "一键 港股", HK_FILL, "FFFFFF", 14),
        ("J9:M10", "一键 韩股", KR_FILL, "FFFFFF", 14),
        ("A11:C12", "显示/隐藏 A股数据", A_LIGHT_FILL, SECONDARY_FG, 11),
        ("D11:F12", "显示/隐藏 美股数据", US_LIGHT_FILL, "9C0006", 11),
        ("G11:I12", "显示/隐藏 港股数据", HK_LIGHT_FILL, "375623", 11),
        ("J11:M12", "显示/隐藏 韩股数据", KR_LIGHT_FILL, "5B2B82", 11),
        ("N2:Q3", "一键全抓 4 市场", COVER_NAVY, "FFFFFF", 13),
        ("N4:Q5", "一键抓取跨市场指标表", A_LIGHT_FILL, SECONDARY_FG, 10),
        ("N6:Q7", "显示/隐藏 所有市场数据", A_LIGHT_FILL, SECONDARY_FG, 11),
        ("N12:Q13", "清空 HTTP 缓存", A_LIGHT_FILL, SECONDARY_FG, 11),
    ]
    for addr, caption, fill_hex, font_color_hex, font_size in placeholders:
        rng = ws_pool.Range(addr)
        rng.Merge()
        rng.Value = caption
        rng.Font.Name = "微软雅黑"
        rng.Font.Size = font_size
        rng.Font.Bold = True
        rng.Font.Color = rgb_long(font_color_hex)
        rng.Interior.Color = rgb_long(fill_hex) if isinstance(fill_hex, str) else fill_hex
        rng.HorizontalAlignment = -4108
        rng.VerticalAlignment = -4108
        rng.WrapText = True

    for addr, caption in (("N11:Q11", "工具"),):
        rng = ws_pool.Range(addr)
        rng.Merge()
        rng.Value = caption
        rng.Font.Name = "微软雅黑"
        rng.Font.Size = 11
        rng.Font.Bold = True
        rng.Font.Color = rgb_long(COVER_NAVY)
        rng.Interior.Color = rgb_long(SECTION_LABEL_FILL)
        rng.HorizontalAlignment = -4108
        rng.VerticalAlignment = -4108
    _apply_range_border(ws_pool.Range("N2:Q13"), "B7C9E2", 2, False)

    try:
        set_cell_comment(
            ws_pool.Range("N12"),
            "删除 .cache/ 目录中的 24h 本地 HTTP 响应缓存。\n"
            "下次抓数会重新发起 HTTP 请求,用于强制刷新或排查数据陈旧问题。\n"
            "日常无需点击。"
        )
    except Exception:
        pass

    hint_rng = ws_pool.Range("S2:V8")
    hint_rng.Merge()
    hint_rng.Value = (
        "💡 使用提示\n\n"
        "1. A2 / A4 选择年度 + 季度\n"
        "2. B5 填雪球 cookie;B6 切原币 / 统一RMB\n"
        "3. 第 14 行起按市场录入公司\n"
        "4. 先点击右侧一键按钮抓取\n"
        "5. 一键全抓会自动刷新所有市场指标表"
    )
    hint_rng.Font.Name = "微软雅黑"
    hint_rng.Font.Size = 10
    hint_rng.Font.Color = rgb_long("000000")
    hint_rng.Interior.Color = rgb_long("FFF8E6")
    hint_rng.HorizontalAlignment = -4131
    hint_rng.VerticalAlignment = -4160
    hint_rng.WrapText = True
    _apply_range_border(hint_rng, "F1C232", 2, False)

    restore_sample_pool_companies(ws_pool, companies, start_row=14)

    for col in ("A", "D", "G", "J"):
        ws_pool.Range(f"{col}14:{col}1000").NumberFormat = "@"

    row_heights = {1: 8, 2: 30, 7: 26, 8: 10, 9: 24, 10: 24, 11: 22, 12: 22, 13: 22}
    for row in range(3, 7):
        row_heights[row] = 22
    for row in range(14, 51):
        row_heights[row] = 22
    for row, height in row_heights.items():
        ws_pool.Rows(row).RowHeight = height

    for addr in ("A14:C50", "D14:F50", "G14:I50", "J14:M50"):
        rng = ws_pool.Range(addr)
        rng.Font.Name = "微软雅黑"
        rng.Font.Size = 10
        _apply_range_border(rng, "D9D9D9", 2, True)

    try:
        ws_pool.Activate()
        ws_pool.Application.ActiveWindow.SplitColumn = 0
        ws_pool.Application.ActiveWindow.SplitRow = 13
        ws_pool.Application.ActiveWindow.FreezePanes = True
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
        ("  E6 = 原币", "各公司按报告币种显示,不做人民币折算。"),
        ("  E6 = 统一RMB", "非人民币口径会按报告期汇率折算成人民币,便于横向对比。"),
        ("  RMB / CNY", "本身就是人民币口径,无需折算。"),
        ("  手工调整汇率", "如需人工修正,可直接改 B-G 列;已生成的报表需要重新点击对应『一键 X 股』按钮后才会反映。"),
        ("", ""),
        ("本地暂存", ""),
        ("  24 小时复用", "同一报告期、同一币种在 24 小时内重复使用已取到的汇率,让重复跑数更快。"),
        ("  清空 HTTP 缓存", "样本池 O11:P12 按钮;清空后下次跑数会重新取汇率和财报数据。"),
        ("", ""),
        ("注意事项", ""),
        ("  报表已生成后", "如需重新换汇请重跑对应『一键 X 股』按钮。"),
        ("  E5 登录凭证", "汇率本身不需要 E5;E5 主要影响港股和部分中概美股备用数据源。"),
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
      E6:N6 = 默认 "原币" (浅黄), 数据验证下拉 "原币,统一RMB"; 不覆盖用户已设的值
    """
    val_cell = ws_pool.Range("E6")
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
        print(f"  ! E6 数据验证添加失败: {e}")

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

    print("  + E6 显示币种 toggle 已配置 (默认 '原币')")


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
        ("1", "在『样本池』E3 选择年度,E4 选择季度。年度留空时,工具尽量使用可取得的最新报告期。"),
        ("2", "如需港股或部分中概美股,在 E5 填入雪球登录凭证;E6 选择『原币』或『统一RMB』。"),
        ("3", "第 14 行起按市场录入公司:A:B=A股,D:E=美股,G:H=港股,J:K=韩股。"),
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
    row = paragraph(row, "登录凭证只保存在本地工作簿中。过期后请重新复制到样本池 E5。", height=24)

    row = section(row, "§ 5 汇率换算说明")
    fx_rows = [
        ("E6 = 原币", "各市场按报告币种或市场默认单位显示,不做 RMB 换算。"),
        ("E6 = 统一RMB", "非人民币报告币种按报告期汇率折算:资产负债表用期末汇率,利润表 / 现金流量表用期间均值。"),
        ("RMB / CNY", "报告币种为人民币时无需折算。"),
        ("手工改汇率", "已写表公式不会自动反向刷新;需要重跑对应写表或合表按钮。"),
    ]
    row = table(row, ("规则", "说明"), fx_rows, widths_override={"B": 20, "C": 68})

    row = section(row, "§ 6 常见问题")
    faq_rows = [
        ("Q1: 港股或中概美股取数失败怎么办?", "先检查 E5 登录凭证是否过期。重新登录后复制新的 xq_a_token,再重跑对应市场。"),
        ("Q2: 切换 E6 后数字变了是否正常?", "正常。『统一RMB』会把外币报告币种折算成人民币,用于跨市场可比。"),
        ("Q3: 跨市场指标表何时刷新?", "点击『一键全抓 4 市场』后自动刷新;也可以单独运行 BuildCrossMarketIndicatorSheet。"),
        ("Q4: 为什么只保留跨市场指标表?", "BS/IS/CF 保留在分市场表中按原披露口径查看;跨市场对比聚焦 18 项标准指标。"),
        ("Q5: 老版样本池升级会丢数据吗?", "不会。已有公司代码和简称会迁移到第 14 行起;安装脚本只重画上方配置区和按钮区。"),
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

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


# Phase 3 按钮规格: (name, caption, macro, height_pt, fill_hex, font_color_hex, font_size, primary?)
PRIMARY_FILL = "4472C4"   # 深蓝
PRIMARY_FG = "FFFFFF"     # 白
SECONDARY_FILL = "D9E1F2" # 浅蓝
SECONDARY_FG = "1F4E79"   # 深蓝字

US_FILL = "C00000"        # 深红 — 美股按钮区分
US_FG = "FFFFFF"
HK_FILL = "548235"        # 深绿 — 港股按钮区分
HK_FG = "FFFFFF"
KR_FILL = "7030A0"        # 深紫 — 韩股按钮区分
KR_FG = "FFFFFF"

BUTTONS = [
    ("BtnRunAll",       "一键全抓",            "模块_总入口.一键全抓",            42, PRIMARY_FILL,   PRIMARY_FG,   13, True),
    ("BtnRunBalance",   "更新A股资产负债表",   "模块_抓资产负债表.Main",         30, SECONDARY_FILL, SECONDARY_FG, 11, False),
    ("BtnRunProfit",    "更新A股利润表",       "模块_抓利润表.Main",             30, SECONDARY_FILL, SECONDARY_FG, 11, False),
    ("BtnRunCash",      "更新A股现金流量表",   "模块_抓现金流量表.Main",         30, SECONDARY_FILL, SECONDARY_FG, 11, False),
    ("BtnRunInd",       "更新A股指标表",       "模块_抓指标表.Main",             30, SECONDARY_FILL, SECONDARY_FG, 11, False),
    # Phase 4b: 美股 (4 张表)
    ("BtnRunUSBalance", "更新美股资产负债表",  "模块_抓美股资产负债表.Main",     30, US_FILL,        US_FG,        11, False),
    ("BtnRunUSProfit",  "更新美股利润表",      "模块_抓美股利润表.Main",         30, US_FILL,        US_FG,        11, False),
    ("BtnRunUSCash",    "更新美股现金流量表",  "模块_抓美股现金流量表.Main",     30, US_FILL,        US_FG,        11, False),
    ("BtnRunUSInd",     "更新美股指标表",      "模块_抓美股指标表.Main",         30, US_FILL,        US_FG,        11, False),
    # Phase 4c: 港股 (4 张表)
    ("BtnRunHKBalance", "更新港股资产负债表",  "模块_抓港股资产负债表.Main",     30, HK_FILL,        HK_FG,        11, False),
    ("BtnRunHKProfit",  "更新港股利润表",      "模块_抓港股利润表.Main",         30, HK_FILL,        HK_FG,        11, False),
    ("BtnRunHKCash",    "更新港股现金流量表",  "模块_抓港股现金流量表.Main",     30, HK_FILL,        HK_FG,        11, False),
    ("BtnRunHKInd",     "更新港股指标表",      "模块_抓港股指标表.Main",         30, HK_FILL,        HK_FG,        11, False),
    # Phase 4d: 韩股 (4 张表)
    ("BtnRunKRBalance", "更新韩股资产负债表",  "模块_抓韩股资产负债表.Main",     30, KR_FILL,        KR_FG,        11, False),
    ("BtnRunKRProfit",  "更新韩股利润表",      "模块_抓韩股利润表.Main",         30, KR_FILL,        KR_FG,        11, False),
    ("BtnRunKRCash",    "更新韩股现金流量表",  "模块_抓韩股现金流量表.Main",     30, KR_FILL,        KR_FG,        11, False),
    ("BtnRunKRInd",     "更新韩股指标表",      "模块_抓韩股指标表.Main",         30, KR_FILL,        KR_FG,        11, False),
]

# 已废弃: install 时从当前 xlsm 主动移除 (即使 modules/ 下仍有遗留也清掉)
DECOMMISSIONED_MODULES = ["模块_抓基本资料"]
DECOMMISSIONED_BUTTONS = ["BtnRunInfo"]
BUTTON_WIDTH_PT = 150    # 容得下 9 个汉字 ("更新A股现金流量表")
BUTTON_GAP_PT = 8
BUTTON_TOP_MARGIN_PT = 5
BUTTON_ANCHOR_COL = "E"   # 按钮锚定在 E 列 (D 是市场<>按钮的间隔列)
BUTTON_COL_WIDTHS = {"D": 3, "E": 24, "F": 4}    # D 间隔, E 按钮主, F 右边距

# Sheet 名迁移 (累积所有迁移规则, install 一次性应用)
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
    Row 1 = 大标题(合并 A1:J1, 深蓝白字), Row 2 = 10 列表头, Row 3+ 由 VBA 写
    冻结 Row 2; 列宽 + 表头颜色与 VBA 端 EnsureDiagnosticSheet() 保持一致 (避免双方互踩)。
    """
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name

    # Row 1: 标题, 合并 A1:J1
    ws.Range("A1").Value = f"{name.replace('_', '')} (每次跑数后自动刷新)"
    ws.Range("A1:J1").Merge()
    title = ws.Range("A1:J1")
    title.Font.Name = "微软雅黑"
    title.Font.Size = 12
    title.Font.Bold = True
    title.Font.Color = rgb_long("FFFFFF")
    title.Interior.Color = rgb_long("4472C4")
    title.HorizontalAlignment = -4108
    title.VerticalAlignment = -4108

    # Row 2: 10 列表头
    headers = ["公司", "报表", "输出指标", "状态", "数据源",
               "Taxonomy", "命中字段", "Unit", "Score", "匹配方式+备注"]
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

    widths = [14, 16, 30, 18, 18, 14, 42, 14, 10, 58]
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
            print(f"  ~ sheet 已存在: {diag_name}")
        else:
            _make_diagnostic_sheet(wb, diag_name)
            print(f"  + sheet 新建: {diag_name}")


def update_intro_sheet(wb):
    """刷新使用说明页, 避免旧模板说明和当前 V3 功能不一致。"""
    try:
        ws = wb.Sheets("使用说明")
    except Exception:
        ws = wb.Worksheets.Add(Before=wb.Sheets(1))
        ws.Name = "使用说明"

    ws.Cells.Clear()
    ws.Columns("A").ColumnWidth = 110
    ws.Columns("B").ColumnWidth = 2
    ws.Rows.RowHeight = 18

    ws.Range("A1").Value = "上市公司财务数据查询"
    ws.Range("A1").Font.Name = "微软雅黑"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Interior.Color = rgb_long("D9E1F2")

    lines = [
        "",
        "用途: 把上市公司财务数据抓成同业对标宽表, 方便横向比较。",
        "当前支持: A股、美股、港股、韩股。后续规划: 更多市场。",
        "作者: Eric Zhang",
        "联系邮箱: 214978902@qq.com",
        "",
        "使用步骤",
        "1. 在『样本池』第 8 行起录入公司: A列=股票代码, B列=股票简称, C列=市场。",
        "2. C列市场可填 A / US / HK / KR; 留空时会自动判断 A/US/HK, KR 请手填。",
        "3. A2 填年份, 例如 2025; A2 留空表示取最新可用期间。",
        "4. A4 选择季度: 全部 / Q1 / Q2 / Q3 / Q4。",
        "5. B5 可填写雪球 xq_a_token cookie; POM、HTT 等 EDGAR 不完整的中概/20-F 公司会自动走雪球 fallback, 港股也使用该 cookie。",
        "6. 点『一键全抓』会顺序更新 A股、美股、港股、韩股 16 张表; 也可以点单张表按钮单独更新。",
        "",
        "输出表",
        "A股: 资产负债表 → 利润表 → 现金流量表 → 指标表。",
        "美股: 资产负债表 → 利润表 → 现金流量表 → 指标表。",
        "港股: 资产负债表 → 利润表 → 现金流量表 → 指标表。",
        "韩股: 资产负债表 → 利润表 → 现金流量表 → 指标表。",
        "指标表统一只保留 18 个标准指标, A股、美股、港股和韩股口径一致。",
        "",
        "宽表结构",
        "第 1 行: 公司名(代码), 横向合并该公司的报告期列。",
        "第 2 行: 报告期, 降序排列。",
        "A/B列: 大类或指标类型、指标名称; 指标表额外有 C列英文指标名。",
        "数值列: 单位按表口径输出; 美股三张财报单位为百万美元; 港股为百万(各家公司报告币种, 见 港股_抓取诊断 Unit 列); 韩股为十亿韩元。",
        "",
        "期间对齐规则",
        "A股宽表使用报告期并集对齐, 便于同行横向比较。",
        "美股宽表按每家公司自身可用期间展开, 不再为其他公司的期间保留空列。",
        "港股宽表同样按每家公司自身可用期间展开, 保留不同公司财年末差异。",
        "韩股宽表同样按每家公司自身可用期间展开; Q1/Q2/Q3 来自季度页,Q4 来自年度页。",
        "",
        "数据源与限制",
        "A股财报来自新浪财经。",
        "美股优先使用 SEC EDGAR companyfacts; EDGAR 缺失或字段不匹配时, 支持的中概股 fallback 到雪球。",
        "港股来自雪球 HK API; 不做汇率换算, 币种以诊断表 Unit 列为准。",
        "韩股来自 stockanalysis.com KRX 财报 HTML 表格; 不需要雪球 cookie。",
        "雪球 cookie 过期时, 请重新复制 xq_a_token 到样本池 B5。",
        "诊断 sheet 中同一 (公司, 指标) 先出现 MISSING_NON_USD、随后出现 OK_XUEQIU 属预期行为:表示 ifrs-full 有字段但单位不是 USD,系统改走雪球兜底。",
        "韩股需要在样本池 C 列手填 KR; 6 位纯数字无法自动区分 A 股与韩股。",
    ]

    for idx, text in enumerate(lines, start=2):
        cell = ws.Cells(idx, 1)
        cell.Value = text
        cell.Font.Name = "微软雅黑"
        cell.Font.Size = 11
        cell.WrapText = True
        cell.VerticalAlignment = -4108
        if text in ("使用步骤", "输出表", "宽表结构", "期间对齐规则", "数据源与限制"):
            cell.Font.Bold = True
            cell.Font.Color = rgb_long("1F4E79")
            cell.Interior.Color = rgb_long("EAF2F8")

    print("  + 使用说明 已刷新")


def reorder_report_sheets(wb):
    """固定工作表 Tab 顺序: 使用说明/样本池, A股四表, 美股四表, 美股诊断, 港股四表, 港股诊断, 韩股四表, 韩股诊断。"""
    desired_order = [
        "使用说明", "样本池",
        "A股_资产负债表", "A股_利润表", "A股_现金流量表", "A股_指标表",
        "美股_资产负债表", "美股_利润表", "美股_现金流量表", "美股_指标表",
        "美股_抓取诊断",
        "港股_资产负债表", "港股_利润表", "港股_现金流量表", "港股_指标表",
        "港股_抓取诊断",
        "韩股_资产负债表", "韩股_利润表", "韩股_现金流量表", "韩股_指标表",
        "韩股_抓取诊断",
    ]
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
    print("  + sheet Tab 顺序已调整")


def install_buttons(ws_pool):
    """
    Phase 3: 用 Shape (msoShapeRoundedRectangle) 替换老的 Forms 按钮
      - 主按钮『一键全抓』: 深蓝 + 白字 + 大字
      - A股按钮: 浅蓝 + 深蓝字; 美股按钮: 深红 + 白字; 港股按钮: 深绿 + 白字; 韩股按钮: 深紫 + 白字
      - 都装在样本池 I 列, 绝对垂直坐标, 无可见边框
      - 装前先 Delete 同名旧按钮 (兼容旧 Forms 按钮 / Phase 2 Shape 按钮)
    """
    # 调宽 I 列以容下按钮
    for col_letter, width in BUTTON_COL_WIDTHS.items():
        ws_pool.Columns(col_letter).ColumnWidth = width

    # 删旧按钮: 当前 BUTTONS 列表内的 + 已废弃的
    target_names = {b[0] for b in BUTTONS} | set(DECOMMISSIONED_BUTTONS)
    # 必须遍历 Shapes 副本, 否则 Delete 会改变集合
    shape_names = [sh.Name for sh in ws_pool.Shapes]
    for name in shape_names:
        if name in target_names:
            ws_pool.Shapes(name).Delete()
            tag = "decommissioned" if name in DECOMMISSIONED_BUTTONS else "existing"
            print(f"  - removed {tag} button: {name}")

    # 计算左边距 = I1 单元格的左 (in points)
    base_left = ws_pool.Range(f"{BUTTON_ANCHOR_COL}1").Left
    base_top = ws_pool.Range(f"{BUTTON_ANCHOR_COL}1").Top

    cur_top = base_top + BUTTON_TOP_MARGIN_PT
    for name, caption, macro, height, fill_hex, font_hex, font_size, is_primary in BUTTONS:
        shape = ws_pool.Shapes.AddShape(
            MSO_SHAPE_ROUNDED_RECT, base_left, cur_top, BUTTON_WIDTH_PT, height
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

        shape.OnAction = macro
        print(f"  + button: {name:15s} [{caption}] @ top={cur_top:.0f}pt h={height}pt → {macro}")

        cur_top += height + BUTTON_GAP_PT


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
                        print(f"  x decommissioned sheet: {dead}")
                    except Exception as e:
                        print(f"  ! 删除 {dead} 失败: {e}")
                    finally:
                        wb.Application.DisplayAlerts = False

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
                install_quarter_cell(ws_pool)
                cleanup_legacy_sample_pool(ws_pool)    # 老布局 → 新布局 迁移
                ensure_market_column(ws_pool)          # 市场列在 C
                install_xueqiu_cookie_cell(ws_pool)    # Phase 4b-5: 雪球 cookie 输入位
                style_sample_pool_data_area(ws_pool)   # Phase 4b-4: 美化
                ensure_market_sheets(wb)
                update_intro_sheet(wb)
                reorder_report_sheets(wb)
                install_buttons(ws_pool)               # 按钮锚 BUTTON_ANCHOR_COL = "E"
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

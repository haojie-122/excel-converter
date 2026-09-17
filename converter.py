# -*- coding: utf-8 -*-
"""
Excel 明细表（数据表） -> 汇总表 转换逻辑

核心规则：
- 数据表第1行空，第2行标题 -> 读取用 header=1
- 每个「序号」恰好一行（51 序号 = 51 行），无需分组聚合
- 「单位种类」= 报关后单位（件/个/套/台/吨）-> 输出「单位种类」
- 「单位」= 法定第一单位 -> 输出「单位」
- 单价（美金）= 总价（美金） / 总数量  （P = Q / M，公式）
- 退税金额 = 内贸金额 / 1.13 * 退税率  （Y = W / 1.13 * X，公式）
- 保费 = 总价（美金） * 1.1 * 0.000145  （T = Q * 1.1 * 0.000145，公式）
- 运费 = MAX(ROUND(SUM(J),2), SUM(K)/1000)*15  （S，J=体积, K=总毛重，公式）
- 换汇成本 = (内贸金额 - 退税金额 + 运费 + 保费) / 总价（美金）  （值）

27 列布局（A~AA）：
 A提单号 B船次 C序号 D外贸合同号 E内贸合同号 F品名 G英文品名
 H打包后件数 I单位种类 J体积 K总毛重 L总净重 M总数量 N单位 O法定单位
 P单价（美金） Q总 价（美金） R换汇成本 S运费 T保费
 U报关单号 V出口/海关编码 W内贸金额 X退税率 Y退税金额 Z申报要素 AA供应商

合并规则：仅「打包后件数/单位种类/体积/总毛重」4 列做空格向上合并；船次不合并。
对齐：所有单元格 水平居中 + 垂直居中 + 自动换行。
输出：第1行空，第2行标题，第3行起数据，按序号升序。

用法：
    result = convert_table(df)            # df 已用 header=1 读取
    to_excel_with_layout(result, out, template_path=源文件)   # 推荐：直接复刻源表格式
"""

import re
from collections import OrderedDict
from copy import copy

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import MergedCell
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter


# =========================================================
# 输出列定义（27列，严格按布局）
# =========================================================
OUTPUT_COLUMNS = [
    "提单号",          # A  1
    "船次",            # B  2
    "序号",            # C  3
    "外贸合同号",       # D  4
    "内贸合同号",       # E  5
    "品名",            # F  6
    "英文品名",         # G  7
    "打包后件数",       # H  8  合并
    "单位种类",         # I  9  合并
    "体积",            # J 10  合并（运费公式引用）
    "总毛重",          # K 11  合并（运费公式引用）
    "总净重",          # L 12
    "总数量",          # M 13  （单价公式引用）
    "单位",            # N 14
    "法定单位",         # O 15
    "单价（美金）",      # P 16  公式 =Q/M
    "总 价（美金）",     # Q 17  （保费/单价公式引用）
    "换汇成本",         # R 18
    "运费",            # S 19  公式
    "保费",            # T 20  公式 =Q*1.1*0.000145
    "报关单号",         # U 21
    "出口/海关编码",     # V 22
    "内贸金额",         # W 23  （退税金额公式引用）
    "退税率",          # X 24  （退税金额公式引用）
    "退税金额",         # Y 25  公式 =W/1.13*X
    "申报要素",         # Z 26
    "供应商",          # AA 27
]

# 需要做「空格向上合并」的列（仅这 4 列，船次不合并）
MERGE_COLUMNS = ["打包后件数", "单位种类", "体积", "总毛重"]

# 输出列 -> 数据表真实列的模糊候选（按优先级匹配）
COLUMN_MAP = {
    "提单号":       ["提单号"],
    "船次":         [],                        # 固定值 "989船"
    "序号":         ["序号"],
    "外贸合同号":    ["外贸合同号"],
    "内贸合同号":    ["内贸合同号"],
    "品名":         ["品名"],
    "英文品名":      ["英文品名"],
    "打包后件数":    ["打包后件数", "打包后 件数", "打包后\n件数"],
    "单位种类":      ["单位种类", "单位种类.1", "单位 种类", "单位 种类.1"],
    "体积":         ["体积"],
    "总毛重":       ["总毛重"],
    "总净重":       ["总净重"],
    "总数量":       ["总数量"],
    "单位":         ["单位", "报关单位"],
    "法定单位":      ["法定单位"],
    "单价（美金）":   ["单价（美金）", "单价\n（美金）"],
    "总 价（美金）":  ["总 价（美金）", "总价（美金）", "总价", "总 价 （美金）",
                     "总 价      \n（美金）"],
    "换汇成本":      ["换汇成本"],
    "运费":         ["运费"],
    "保费":         ["保费"],
    "报关单号":      ["报关单号"],
    "出口/海关编码":  ["出口/海关编码", "海关编码"],
    "内贸金额":      ["内贸金额"],
    "退税率":       ["退税率"],
    "退税金额":      ["退税金额"],
    "申报要素":      ["申报要素"],
    "供应商":       ["供应商"],
}


# =========================================================
# 工具函数
# =========================================================
def clean_col_name(c):
    """列名归一化：去掉换行/空格/全角空格，转成可匹配形式。"""
    if c is None or (isinstance(c, float) and pd.isna(c)):
        return ""
    return re.sub(r"[\s\u3000]+", "", str(c)).strip()


def safe_float(x, default=None):
    if x is None:
        return default
    if isinstance(x, float) and pd.isna(x):
        return default
    if isinstance(x, str):
        s = x.replace(",", "").replace("，", "").strip()
        if s in ("", "-", "—", "–", "null", "None", "NA", "N/A", "无", "空"):
            return default
        try:
            return float(s)
        except ValueError:
            return default
    try:
        return float(x)
    except (ValueError, TypeError):
        return default


def safe_int(x, default=None):
    v = safe_float(x, default)
    if v is None:
        return default
    if float(v) == int(v):
        return int(v)
    return v


def find_real_column(columns_cleaned_map, candidates):
    """在「清洗后列名 -> 原始列名」映射里按候选列表顺序匹配。"""
    if not candidates:
        return None
    # 1) 精确匹配（清洗后）
    for cand in candidates:
        ncand = clean_col_name(cand)
        if ncand in columns_cleaned_map:
            return columns_cleaned_map[ncand]
    # 2) 包含匹配（去掉括号/空格干扰）
    def strip_symbol(s):
        return re.sub(r"[（）()\s]", "", s)

    for cand in candidates:
        k = strip_symbol(clean_col_name(cand))
        if not k:
            continue
        for key, raw in columns_cleaned_map.items():
            if k in strip_symbol(key):
                return raw
    return None


def is_writable(cell):
    """判断单元格是否可写（排除合并区间内的非锚点 MergedCell）。"""
    return not isinstance(cell, MergedCell)


def set_cell(cell, value):
    """安全写值：遇到 MergedCell 只读占位符则跳过。"""
    if isinstance(cell, MergedCell):
        return
    cell.value = value


def set_style(dst, src):
    """把 src 单元格的样式复制到 dst，不碰 value（避免 MergedCell 报错）。"""
    if isinstance(dst, MergedCell) or isinstance(src, MergedCell):
        return
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def _cell_is_empty(v):
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


# =========================================================
# 主转换（纯 pandas，返回按 OUTPUT_COLUMNS 顺序排列的 DataFrame）
# =========================================================
def convert_table(df: pd.DataFrame) -> pd.DataFrame:
    # 建立「清洗后列名 -> 原始列名」映射（去重，保留首个）
    cleaned_map = {}
    for orig in df.columns:
        nc = clean_col_name(orig)
        if nc and nc not in cleaned_map:
            cleaned_map[nc] = orig

    # 重命名列：原始列名 -> 清洗后列名（避免重复列名）
    rename = {orig: cleaned_map.get(clean_col_name(orig), orig)
              for orig in df.columns}
    df = df.rename(columns=rename).copy()

    # 序号列
    seq_real = find_real_column(cleaned_map, ["序号"])
    if seq_real is None:
        raise ValueError(f"未找到「序号」列。实际列名：{list(cleaned_map.values())}")

    df["__seq"] = df[seq_real].apply(lambda x: safe_int(x, default=None))
    df = df.dropna(subset=["__seq"]).copy()          # 去掉合计行（序号为空）
    if len(df) == 0:
        raise ValueError("序号列没有有效数值，请检查表一数据。")
    df["__seq"] = df["__seq"].astype(int)

    # 预解析输出列 -> 数据表真实列
    real_map = {}
    for out_col, candidates in COLUMN_MAP.items():
        real = find_real_column(cleaned_map, candidates)
        if real:
            real_map[out_col] = real

    def get_val(row, out_col, default=None):
        real = real_map.get(out_col)
        if not real:
            return default
        v = row[real]
        if _cell_is_empty(v):
            return default
        return v

    rows = []
    for _, r in df.sort_values("__seq").iterrows():
        total_price = safe_float(get_val(r, "总 价（美金）"))   # Q
        total_qty = safe_float(get_val(r, "总数量"))            # M
        irm = safe_float(get_val(r, "内贸金额"))                # W
        trr = safe_float(get_val(r, "退税率"))                  # X
        freight = safe_float(get_val(r, "运费"), 0) or 0
        premium = safe_float(get_val(r, "保费"), 0) or 0

        # 退税金额：数据表有则透传，否则 None（由公式 Y=W/1.13*X 计算）
        tax_refund = safe_float(get_val(r, "退税金额"))

        # 换汇成本 = (内贸金额 - 退税金额 + 运费 + 保费) / 总价（美金）
        cost = None
        if total_price and total_price > 0:
            net_tax = tax_refund if tax_refund is not None else (irm or 0) / 1.13 * (trr or 0)
            numerator = (irm or 0) - net_tax + freight + premium
            cost = round(numerator / total_price, 8)

        row = OrderedDict()
        row["序号"] = safe_int(r["__seq"])
        row["提单号"] = get_val(r, "提单号", "") or ""
        row["船次"] = "989船"
        row["外贸合同号"] = get_val(r, "外贸合同号", "") or ""
        row["内贸合同号"] = get_val(r, "内贸合同号", "") or ""
        row["品名"] = get_val(r, "品名", "") or ""
        row["英文品名"] = get_val(r, "英文品名", "") or ""
        row["打包后件数"] = safe_int(get_val(r, "打包后件数"), default=None)
        row["单位种类"] = get_val(r, "单位种类", "") or ""
        row["体积"] = safe_float(get_val(r, "体积"), default=None)        # J
        row["总毛重"] = safe_float(get_val(r, "总毛重"), default=None)     # K
        row["总净重"] = safe_float(get_val(r, "总净重"), default=None)
        row["总数量"] = safe_float(get_val(r, "总数量"), default=None)     # M
        row["单位"] = get_val(r, "单位", "") or ""
        row["法定单位"] = get_val(r, "法定单位", "") or ""
        row["单价（美金）"] = None        # 公式 P = Q / M
        row["总 价（美金）"] = total_price                                # Q
        row["换汇成本"] = cost                                            # R
        row["运费"] = None              # 公式
        row["保费"] = None              # 公式 T = Q * 1.1 * 0.000145
        row["报关单号"] = get_val(r, "报关单号", "") or ""
        row["出口/海关编码"] = get_val(r, "出口/海关编码", "") or ""
        row["内贸金额"] = irm                                             # W
        row["退税率"] = trr                                               # X
        row["退税金额"] = tax_refund                                      # Y（值 + 公式双保险）
        row["申报要素"] = get_val(r, "申报要素", "") or ""
        row["供应商"] = get_val(r, "供应商", "") or ""

        rows.append(row)

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


# =========================================================
# 导出 Excel
# =========================================================
def _col_letter(name):
    """OUTPUT_COLUMNS 中列名对应的字母。"""
    return get_column_letter(OUTPUT_COLUMNS.index(name) + 1)


def _merge_blank_up(ws, col_idx, start_row, end_row):
    """
    空格向上合并：把「连续为空的单元格」向上合并到该段第一个非空单元格。
    例：序号9~12 的体积均为空，序号8=渣浆泵有值 -> 把 9~12 的空格合并到序号8 的单元格下方区间。
    遇到非空值则重新开始一个新区间。
    关键点：合并前先 unmerge 该列已有合并，且只合并普通单元格区间（跳过 MergedCell）。
    """
    # 先解除本列在 [start_row, end_row] 内已有的合并区间，避免残留只读占位符
    existing = [rng for rng in ws.merged_cells.ranges
                if rng.min_col == col_idx and rng.min_row >= start_row and rng.max_row <= end_row]
    for rng in existing:
        ws.unmerge_cells(str(rng))

    r = start_row
    while r <= end_row:
        cell = ws.cell(row=r, column=col_idx)
        if is_writable(cell) and not _cell_is_empty(cell.value):
            # 找从 r+1 起连续的空格
            run_end = r
            while run_end + 1 <= end_row:
                nxt = ws.cell(row=run_end + 1, column=col_idx)
                if is_writable(nxt) and _cell_is_empty(nxt.value):
                    run_end += 1
                else:
                    break
            if run_end > r:
                ws.merge_cells(start_row=r, start_column=col_idx,
                               end_row=run_end, end_column=col_idx)
            r = run_end + 1
        else:
            r += 1


def to_excel_with_layout(result_df, path, template_path=None):
    """
    导出汇总表。
    - template_path 不为空时：以源表为模板，直接复刻其全部格式（列宽/字体/数字格式/
      边框/合并/冻结/筛选），仅重写数据值与公式。这是最稳妥的方式。
    - 否则：手工构建工作簿，应用完整格式规范。
    """
    if template_path:
        _write_from_template(result_df, path, template_path)
    else:
        _write_from_scratch(result_df, path)


# ---------- 方式一：以源表为模板（推荐） ----------
def _write_from_template(result_df, path, template_path):
    wb = load_workbook(template_path)
    ws = wb.active
    ws.title = "汇总表"

    n = len(result_df)
    data_first = 3          # 第1行空，第2行标题，第3行起数据
    data_last = 2 + n       # 数据结束行（Excel 行号）
    total_row = data_last + 1

    # 列字母
    L = _col_letter

    # 数据区已有合并会随单元格保留；我们先重写锚点单元格的值。
    # 写数据值（逐行，遇到 MergedCell 只读占位符则跳过）
    for offset, (_, row) in enumerate(result_df.iterrows()):
        excel_row = data_first + offset
        for col_name in OUTPUT_COLUMNS:
            col_idx = OUTPUT_COLUMNS.index(col_name) + 1
            cell = ws.cell(row=excel_row, column=col_idx)
            if not is_writable(cell):
                continue
            val = row[col_name]
            if val is None or (isinstance(val, str) and val == ""):
                cell.value = None
            else:
                cell.value = val

    # 写公式列（只写锚点单元格；合并区间锚点 = 区间首行，可写）
    for excel_row in range(data_first, data_last + 1):
        _safe_set_formula(ws, f"{L('单价（美金）')}{excel_row}",
                          f"={L('总 价（美金）')}{excel_row}/{L('总数量')}{excel_row}")
        _safe_set_formula(ws, f"{L('运费')}{excel_row}",
                          f"=MAX(ROUND(SUM({L('体积')}{excel_row}),2),"
                          f"SUM({L('总毛重')}{excel_row})/1000)*15")
        _safe_set_formula(ws, f"{L('保费')}{excel_row}",
                          f"={L('总 价（美金）')}{excel_row}*1.1*0.000145")
        _safe_set_formula(ws, f"{L('退税金额')}{excel_row}",
                          f"={L('内贸金额')}{excel_row}/1.13*{L('退税率')}{excel_row}")

    # 4 列空格向上合并（船次不在其中，天然不合并）
    for col_name in MERGE_COLUMNS:
        _merge_blank_up(ws, OUTPUT_COLUMNS.index(col_name) + 1,
                        data_first, data_last)

    # 合计行：删除原有内容，写 SUM 公式
    if total_row <= ws.max_row:
        for col_name in ["打包后件数", "体积", "总毛重", "总净重", "总数量",
                         "总 价（美金）", "内贸金额", "退税金额"]:
            cell = ws.cell(row=total_row, column=OUTPUT_COLUMNS.index(col_name) + 1)
            if is_writable(cell):
                letter = L(col_name)
                cell.value = f"=SUM({letter}{data_first}:{letter}{data_last})"

    wb.save(path)


def _safe_set_formula(ws, coordinate, formula):
    """在坐标处写公式，若该单元格是 MergedCell 则跳过（保留锚点的公式即可）。"""
    from openpyxl.utils.cell import coordinate_from_string
    coord_str, row = coordinate_from_string(coordinate)
    col = ws[coordinate].column
    cell = ws.cell(row=row, column=col)
    if is_writable(cell):
        cell.value = formula


# ---------- 方式二：手工构建（无模板时的完整规范） ----------
def _write_from_scratch(result_df, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "汇总表"

    n = len(result_df)
    data_first = 3
    data_last = 2 + n
    total_row = data_last + 1
    L = _col_letter

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数字格式（完全对齐源表）
    NUM_FMT = {
        "单价（美金）": r'\$#,##0.00;-\$#,##0.00',
        "总 价（美金）": r'\$#,##0.00;-\$#,##0.00',
        "运费":         r'\$#,##0.00;-\$#,##0.00',
        "保费":         r'\$#,##0.00;-\$#,##0.00',
        "换汇成本":     r'0.00_ ;[Red]\(0.00\)',
        "内贸金额":     r'"￥"#,##0.00;"￥"\-#,##0.00',
        "退税金额":     r'"￥"#,##0.00;"￥"\-#,##0.00',
        "退税率":       '0%',
    }

    # 列宽（完全对齐源表）
    COL_WIDTHS = {
        "提单号": 13.375, "船次": 8.375, "序号": 7.0,
        "外贸合同号": 19.258, "内贸合同号": 26.792,
        "品名": 20.125, "英文品名": 28.5,
        "打包后件数": 11.125, "单位种类": 9.0,
        "体积": 9.0, "总毛重": 11.75, "总净重": 11.5, "总数量": 10.0,
        "单位": 7.375, "法定单位": 7.625,
        "单价（美金）": 11.875, "总 价（美金）": 14.375, "换汇成本": 7.625,
        "运费": 12.5, "保费": 8.625,
        "报关单号": 21.625, "出口/海关编码": 16.5,
        "内贸金额": 16.5, "退税率": 10.5, "退税金额": 15.375,
        "申报要素": 19.875, "供应商": 20.625,
    }

    # 标题行（第2行）
    for col_name in OUTPUT_COLUMNS:
        col_idx = OUTPUT_COLUMNS.index(col_name) + 1
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        cell.font = Font(name="宋体", size=10, bold=True)
        cell.alignment = center_wrap
        cell.border = border

    # 数据行：先填值
    for offset, (_, row) in enumerate(result_df.iterrows()):
        excel_row = data_first + offset
        for col_name in OUTPUT_COLUMNS:
            col_idx = OUTPUT_COLUMNS.index(col_name) + 1
            cell = ws.cell(row=excel_row, column=col_idx)
            val = row[col_name]
            if not _cell_is_empty(val):
                cell.value = val

    # 公式列
    for excel_row in range(data_first, data_last + 1):
        ws[f"{L('单价（美金）')}{excel_row}"] = f"={L('总 价（美金）')}{excel_row}/{L('总数量')}{excel_row}"
        ws[f"{L('运费')}{excel_row}"] = (
            f"=MAX(ROUND(SUM({L('体积')}{excel_row}),2),"
            f"SUM({L('总毛重')}{excel_row})/1000)*15"
        )
        ws[f"{L('保费')}{excel_row}"] = f"={L('总 价（美金）')}{excel_row}*1.1*0.000145"
        ws[f"{L('退税金额')}{excel_row}"] = f"={L('内贸金额')}{excel_row}/1.13*{L('退税率')}{excel_row}"

    # 4 列空格向上合并（船次不合并）
    for col_name in MERGE_COLUMNS:
        _merge_blank_up(ws, OUTPUT_COLUMNS.index(col_name) + 1,
                        data_first, data_last)

    # 套用样式（合并后只遍历锚点单元格）
    for col_name in OUTPUT_COLUMNS:
        col_idx = OUTPUT_COLUMNS.index(col_name) + 1
        fmt = NUM_FMT.get(col_name, "General")
        is_long_text = col_name in {"外贸合同号", "内贸合同号", "英文品名", "品名",
                                    "出口/海关编码", "申报要素", "供应商", "报关单号"}
        align = Alignment(horizontal="center", vertical="center",
                          wrap_text=True) if is_long_text else center_wrap
        for excel_row in range(data_first, data_last + 1):
            cell = ws.cell(row=excel_row, column=col_idx)
            if not is_writable(cell):
                continue
            cell.alignment = align
            cell.border = border
            cell.number_format = fmt
            # 字体：中文用宋体，其余 Arial（简化为统一宋体，源表即如此）
            cell.font = Font(name="宋体", size=10)

    # 合计行
    for col_name in ["打包后件数", "体积", "总毛重", "总净重", "总数量",
                     "总 价（美金）", "内贸金额", "退税金额"]:
        cell = ws.cell(row=total_row, column=OUTPUT_COLUMNS.index(col_name) + 1)
        letter = L(col_name)
        cell.value = f"=SUM({letter}{data_first}:{letter}{data_last})"
        cell.alignment = center_wrap
        cell.border = border
        cell.font = Font(name="宋体", size=10, bold=True)
        cell.number_format = NUM_FMT.get(col_name, "General")

    # 列宽
    for col_name, width in COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(OUTPUT_COLUMNS.index(col_name) + 1)].width = width

    # 行高
    ws.row_dimensions[1].height = 20.1
    ws.row_dimensions[2].height = 24.95
    for r in range(data_first, data_last + 1):
        ws.row_dimensions[r].height = 20.0
    ws.row_dimensions[total_row].height = 18.0

    # 冻结首两行 + 自动筛选
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"B2:{get_column_letter(len(OUTPUT_COLUMNS))}{total_row}"

    wb.save(path)


# =========================================================
# 兼容旧接口
# =========================================================
def transform(df):
    return convert_table(df)

"""
Excel 明细表（数据表）-> 汇总表（表一格式）转换逻辑

核心规则：
- 数据表第1行空，第2行标题 → 读取用 header=1
- 每个「序号」恰好一行（51序号 = 51行），无需分组聚合
- 「单位种类」= 报关后单位
- 「单位」= 法定第一单位
- 换汇成本 = (内贸金额 - 退税金额 + 运费 + 保费) / 总价（美金）

27 列布局（A~AA）：
 A提单号 B船次 C序号 D外贸合同号 E内贸合同号 F品名 G英文品名
 H打包后件数 I单位种类 J体积 K总毛重 L总净重 M总数量 N单位 O法定单位
 P单价（美金） Q总 价（美金） R换汇成本 S运费 T保费
 U报关单号 V出口/海关编码 W内贸金额 X退税率 Y退税金额 Z申报要素 AA供应商

公式（均为本行引用）：
- 单价（美金） P = Q / M
- 运费       S = MAX(ROUND(SUM(J),2), SUM(K)/1000)*15
- 保费       T = Q * 1.1 * 0.000145
- 退税金额   Y = W / 1.13 * X

合并规则：仅「打包后件数/单位种类/体积/总毛重」4 列空格向上合并；船次不合并。
格式：完全复刻源表 —— 宋体/Arial 字体、居中、细边框、各列数字格式、列宽、行高、
      冻结首行、自动筛选，末尾追加合计行（SUM）。
"""

import re
from collections import OrderedDict
from copy import copy
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# 输出列定义（27列，完全对齐表一）
# =========================================================
OUTPUT_COLUMNS = [
    "提单号",          # A 1
    "船次",            # B 2
    "序号",            # C 3
    "外贸合同号",       # D 4
    "内贸合同号",       # E 5
    "品名",            # F 6
    "英文品名",         # G 7
    "打包后件数",       # H 8  合并
    "单位种类",         # I 9  合并
    "体积",            # J 10 合并
    "总毛重",          # K 11 合并
    "总净重",          # L 12
    "总数量",          # M 13
    "单位",            # N 14
    "法定单位",         # O 15
    "单价（美金）",      # P 16 公式 =Q/M
    "总 价（美金）",     # Q 17 公式引用
    "换汇成本",         # R 18
    "运费",            # S 19 公式
    "保费",            # T 20 公式 =Q*1.1 * 0.000145
    "报关单号",         # U 21
    "出口/海关编码",     # V 22
    "内贸金额",         # W 23
    "退税率",          # X 24
    "退税金额",         # Y 25 公式 =W/1.13*X
    "申报要素",         # Z 26
    "供应商",          # AA 27
]

# 需要做「空格向上合并」的列（仅这 4 列，船次不合并）
MERGE_COLUMNS = ["打包后件数", "单位种类", "体积", "总毛重"]


# 输出列 -> 数据表真实列的模糊候选
COLUMN_MAP = {
    "提单号":       ["提单号"],
    "船次":         [],                        # 固定值 989船
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
    "单价（美金）":   [],
    "总 价（美金）":  ["总价（美金）", "总 价（美金）", "总价"],
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
# 每列的数字格式（完全复刻源表）
# =========================================================
NUM_FMT = {
    "单价（美金）": r'\$#,##0.00;-\$#,##0.00',
    "总 价（美金）": r'\$#,##0.00;-\$#,##0.00',
    "运费":         r'\$#,##0.00;-\$#,##0.00',
    "保费":         r'\$#,##0.00;-\$#,##0.00',
    "换汇成本":     '0.00_);[Red]\(0.00\)',
    "内贸金额":     r'"￥"#,##0.00;"￥"\-#,##0.00',
    "退税金额":     r'"￥"#,##0.00;"￥"\-#,##0.00',
    "退税率":       '0%',
}

# 每列列宽（完全复刻源表，单位字符）
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

# 需要自动换行的列（长文本列）
WRAP_COLUMNS = {"外贸合同号", "内贸合同号", "英文品名", "品名",
                "出口/海关编码", "申报要素", "供应商", "报关单号"}


# =========================================================
# 工具函数
# =========================================================
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


def clean_col_name(c):
    if pd.isna(c) or str(c).strip() == "":
        return ""
    return re.sub(r"\s+", "", str(c)).strip()


def find_real_column(cleaned_map, candidates):
    if not candidates:
        return None
    for cand in candidates:
        ncand = clean_col_name(cand)
        if ncand in cleaned_map:
            return cleaned_map[ncand]

    def strip_symbol(s):
        return re.sub(r"[（）()\s]", "", s)

    for cand in candidates:
        k = strip_symbol(clean_col_name(cand))
        if not k:
            continue
        for key, raw in cleaned_map.items():
            if k in strip_symbol(key):
                return raw
    return None


def _cell_is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


# =========================================================
# 主转换
# =========================================================
def convert_table(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_map = {}
    for orig in df.columns:
        nc = clean_col_name(orig)
        if nc and nc not in cleaned_map:
            cleaned_map[nc] = orig

    df = df.copy()
    df.columns = [cleaned_map.get(clean_col_name(c), c) for c in df.columns]

    seq_real = find_real_column(cleaned_map, ["序号"])
    if seq_real is None:
        raise ValueError(f"未找到「序号」列。实际列名：{list(cleaned_map.values())}")

    df["__seq"] = df[seq_real].apply(lambda x: safe_int(x, default=None))
    df = df.dropna(subset=["__seq"]).copy()
    if len(df) == 0:
        raise ValueError("序号列没有有效数值。")
    df["__seq"] = df["__seq"].astype(int)

    real_map = {}
    for out_col, candidates in COLUMN_MAP.items():
        real = find_real_column(cleaned_map, candidates)
        if real:
            real_map[out_col] = real

    rows = []
    for _, r in df.sort_values("__seq").iterrows():
        def get_val(out_col, default=None):
            real = real_map.get(out_col)
            if not real:
                return default
            v = r[real]
            if _cell_is_empty(v):
                return default
            return v

        total_price = safe_float(get_val("总 价（美金）"))   # Q
        total_qty = safe_float(get_val("总数量"))            # M
        irm = safe_float(get_val("内贸金额"))                # W
        trr = safe_float(get_val("退税率"))                  # X

        # 退税金额：公式 Y = W/1.13*X（有源值则透传，否则留 None 由公式算）
        tax_refund = safe_float(get_val("退税金额"))
        if tax_refund is None and irm is not None and trr is not None:
            tax_refund = round(irm / 1.13 * trr, 6)

        # 换汇成本
        cost = safe_float(get_val("换汇成本"))
        if cost is None and total_price and total_price > 0:
            freight_tmp = safe_float(get_val("运费"), 0)
            premium_tmp = safe_float(get_val("保费"))
            if premium_tmp is None and total_price:
                premium_tmp = total_price * 1.1 * 0.000145
            numerator = (irm or 0) - (tax_refund or 0) + freight_tmp + (premium_tmp or 0)
            cost = round(numerator / total_price, 8)

        row = OrderedDict()
        row["序号"] = safe_int(r["__seq"])
        row["提单号"] = get_val("提单号", "") or ""
        row["船次"] = "989船"
        row["外贸合同号"] = get_val("外贸合同号", "") or ""
        row["内贸合同号"] = get_val("内贸合同号", "") or ""
        row["品名"] = get_val("品名", "") or ""
        row["英文品名"] = get_val("英文品名", "") or ""
        row["打包后件数"] = safe_int(get_val("打包后件数"), default=None)
        row["单位种类"] = get_val("单位种类", "") or ""
        row["体积"] = safe_float(get_val("体积"), default=None)
        row["总毛重"] = safe_float(get_val("总毛重"), default=None)
        row["总净重"] = safe_float(get_val("总净重"), default=None)
        row["总数量"] = safe_float(get_val("总数量"), default=None)
        row["单位"] = get_val("单位", "") or ""
        row["法定单位"] = get_val("法定单位", "") or ""
        row["单价（美金）"] = None   # 公式 P=Q/M
        row["总 价（美金）"] = total_price
        row["换汇成本"] = cost
        row["运费"] = None           # 公式
        row["保费"] = None           # 公式
        row["报关单号"] = get_val("报关单号", "") or ""
        row["出口/海关编码"] = get_val("出口/海关编码", "") or ""
        row["内贸金额"] = irm
        row["退税率"] = trr
        row["退税金额"] = tax_refund
        row["申报要素"] = get_val("申报要素", "") or ""
        row["供应商"] = get_val("供应商", "") or ""

        rows.append(row)

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


# =========================================================
# 导出 Excel：完全复刻源表格式
# =========================================================
def to_excel_with_layout(result_df, path, template_path=None):
    """
    template_path: 源表 Excel 路径。若提供，则直接从源表复制样式作为基准，
    保证列宽/字体/数字格式/冻结/筛选/打印设置 100% 一致。
    """
    col_idx_map = {name: idx + 1 for idx, name in enumerate(result_df.columns)}

    def L(name):
        return get_column_letter(col_idx_map[name])

    n = len(result_df)
    data_first = 3      # 第1行空，第2行标题，第3行起数据
    data_last = 2 + n   # 数据结束行
    total_row = data_last + 1   # 合计行

    # ---------- 优先：以源表为模板复制样式 ----------
    if template_path:
        _write_from_template(result_df, path, template_path, n)
        return path

    # ---------- 兜底：手动构造（无模板时） ----------
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "汇总表"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 标题行（第2行）
    header_font = Font(name="宋体", size=10, bold=True)
    for col_idx, name in enumerate(result_df.columns, start=1):
        cell = ws.cell(row=2, column=col_idx, value=name)
        cell.font = header_font
        cell.alignment = center_wrap
        cell.border = border

    # 数据行
    data_font = Font(name="Arial", size=10)
    for row_offset, (_, row) in enumerate(result_df.iterrows()):
        excel_row = data_first + row_offset
        for col_name in result_df.columns:
            col_idx = col_idx_map[col_name]
            cell = ws.cell(row=excel_row, column=col_idx)
            val = row[col_name]
            if val is None or (isinstance(val, str) and val == ""):
                continue
            cell.value = val
            cell.font = data_font
            cell.border = border
            cell.number_format = NUM_FMT.get(col_name, "General")
            cell.alignment = center_wrap if col_name in WRAP_COLUMNS else center

    # 公式列
    for excel_row in range(data_first, data_last + 1):
        ws[f"{L('单价（美金）')}{excel_row}"] = (
            f"={L('总 价（美金）')}{excel_row}/{L('总数量')}{excel_row}"
        )
        ws[f"{L('运费')}{excel_row}"] = (
            f"=MAX(ROUND(SUM({L('体积')}{excel_row}),2),"
            f"SUM({L('总毛重')}{excel_row})/1000)*15"
        )
        ws[f"{L('保费')}{excel_row}"] = (
            f"={L('总 价（美金）')}{excel_row}*1.1 * 0.000145"
        )
        ws[f"{L('退税金额')}{excel_row}"] = (
            f"={L('内贸金额')}{excel_row}/1.13*{L('退税率')}{excel_row}"
        )

    # 合计行
    total_font = Font(name="Arial", size=10, bold=True)
    for col_name in ["打包后件数", "体积", "总毛重", "总净重", "总数量",
                     "总 价（美金）", "内贸金额", "退税金额"]:
        cell = ws.cell(row=total_row, column=col_idx_map[col_name])
        letter = L(col_name)
        cell.value = f"=SUM({letter}{data_first}:{letter}{data_last})"
        cell.font = total_font
        cell.alignment = center
        cell.border = border
        cell.number_format = NUM_FMT.get(col_name, "General")
    ws.row_dimensions[total_row].height = 18

    # 空格向上合并（仅 4 列）
    for col_name in MERGE_COLUMNS:
        _merge_consecutive_blanks(ws, col_idx_map[col_name], data_first, data_last)

    # 列宽
    for col_name, width in COL_WIDTHS.items():
        ws.column_dimensions[L(col_name)].width = width

    # 行高
    ws.row_dimensions[1].height = 20.1
    ws.row_dimensions[2].height = 24.95
    for r in range(data_first, data_last + 1):
        ws.row_dimensions[r].height = 20.0

    # 冻结首行 + 自动筛选
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"B2:AC{total_row}"

    wb.save(path)
    return path


def _write_from_template(result_df, path, template_path, n):
    """
    以源表 Excel 为模板：
    - 复制其列宽、字体、数字格式、边框、冻结、筛选、打印设置
    - 清除其数据，写入新数据
    保证格式与源表 100% 一致。
    """
    wb = load_workbook(template_path)
    ws = wb.active
    ws.title = "汇总表"

    col_idx_map = {name: idx + 1 for idx, name in enumerate(result_df.columns)}
    data_first = 3
    data_last = 2 + n
    total_row = data_last + 1
    last_col = len(result_df.columns)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def L(name):
        return get_column_letter(col_idx_map[name])

    # ---------- 清除原有数据（保留第1行空、第2行标题） ----------
    # 关键：保留源表 H5:H6、H7:H8 ... 等合并结构，只重写锚点值
    for r in range(data_first, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            # 跳过合并区间内的非锚点单元格（MergedCell，read-only）
            if type(cell).__name__ == "MergedCell":
                continue
            cell.value = None

    # ---------- 重写标题行（第2行，沿用源表样式） ----------
    for col_idx, name in enumerate(result_df.columns, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = name
        if not cell.font or not cell.font.name:
            cell.font = Font(name="宋体", size=10, bold=True)
        if not cell.alignment or not cell.alignment.horizontal:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    # ---------- 写数据 ----------
    for row_offset, (_, row) in enumerate(result_df.iterrows()):
        excel_row = data_first + row_offset
        for col_name in result_df.columns:
            col_idx = col_idx_map[col_name]
            cell = ws.cell(row=excel_row, column=col_idx)
            if type(cell).__name__ == "MergedCell":
                continue  # 非锚点，跳过
            val = row[col_name]
            if val is None or (isinstance(val, str) and val == ""):
                continue
            cell.value = val

    # ---------- 空格向上合并（仅 4 列，船次除外） ----------
    # 必须在套样式【之前】完成：先定型合并结构，再逐单元格复制样式
    for col_name in MERGE_COLUMNS:
        _merge_consecutive_blanks(ws, col_idx_map[col_name], data_first, data_last)

    # ---------- 逐单元格套用源表样式（以第3行为模板按列复制） ----------
    # 合并区间内的非锚点单元格不单独设值，只对其锚点(top-left)套样式
    merged_anchors = {}
    for rng in ws.merged_cells.ranges:
        merged_anchors[(rng.min_row, rng.min_col)] = rng

    for col_idx in range(1, last_col + 1):
        template_cell = ws.cell(row=data_first, column=col_idx)
        t_font = copy(template_cell.font)
        t_align = copy(template_cell.alignment)
        t_fmt = template_cell.number_format
        for r in range(data_first, data_last + 1):
            cell = ws.cell(row=r, column=col_idx)
            rng = merged_anchors.get((r, col_idx))
            if rng and (r != rng.min_row or col_idx != rng.min_col):
                continue  # 非锚点，跳过
            cell.font = copy(t_font)
            cell.alignment = copy(t_align)
            cell.number_format = t_fmt
            cell.border = border

    # ---------- 公式列（逐行，只写锚点单元格） ----------
    for excel_row in range(data_first, data_last + 1):
        for col_name, formula in [
            ("单价（美金）", f"={L('总 价（美金）')}{excel_row}/{L('总数量')}{excel_row}"),
            ("运费", f"=MAX(ROUND(SUM({L('体积')}{excel_row}),2),SUM({L('总毛重')}{excel_row})/1000)*15"),
            ("保费", f"={L('总 价（美金）')}{excel_row}*1.1 * 0.000145"),
            ("退税金额", f"={L('内贸金额')}{excel_row}/1.13*{L('退税率')}{excel_row}"),
        ]:
            cell = ws.cell(row=excel_row, column=col_idx_map[col_name])
            if type(cell).__name__ == "MergedCell":
                continue
            cell.value = formula

    # ---------- 合计行 ----------
    total_font = Font(name="Arial", size=10, bold=True)
    center = Alignment(horizontal="center", vertical="center")
    for col_name in ["打包后件数", "体积", "总毛重", "总净重", "总数量",
                     "总 价（美金）", "内贸金额", "退税金额"]:
        cell = ws.cell(row=total_row, column=col_idx_map[col_name])
        letter = L(col_name)
        cell.value = f"=SUM({letter}{data_first}:{letter}{data_last})"
        cell.font = total_font
        cell.alignment = center
        cell.border = border
        cell.number_format = NUM_FMT.get(col_name, "General")
    ws.row_dimensions[total_row] = copy(ws.row_dimensions[data_first])
    ws.row_dimensions[total_row].height = 18

    # ---------- 恢复冻结 + 筛选范围 ----------
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"B2:AC{total_row}"

    wb.save(path)


def _merge_consecutive_blanks(ws, col_idx, start, end):
    """
    空格向上合并：对每个「非空单元格」，将其下方连续的空格合并到该单元格。
    例：序号9=1, 序号10~13=空 → 合并 H11:H15（H11 为锚点，H12~H15 并入）
    """
    r = start
    while r <= end:
        if not _cell_is_empty(ws.cell(row=r, column=col_idx).value):
            # 从 r+1 开始找连续空格
            merge_end = r
            while merge_end + 1 <= end and _cell_is_empty(
                ws.cell(row=merge_end + 1, column=col_idx).value
            ):
                merge_end += 1
            if merge_end > r:
                ws.merge_cells(
                    start_row=r, start_column=col_idx,
                    end_row=merge_end, end_column=col_idx,
                )
            r = merge_end + 1
        else:
            r += 1


def transform(df):
    """兼容旧接口"""
    return convert_table(df)

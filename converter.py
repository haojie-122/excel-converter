"""
Excel 明细表（数据表）-> 汇总表（表一格式）转换逻辑

核心规则：
- 数据表第1行空，第2行标题 → 读取用 header=1
- 每个「序号」恰好一行（51序号 = 51行），无需分组聚合
- 「单位种类.1」= 报关后单位 → 输出「单位种类」
- 「报关单位」= 法定第一单位 → 输出「单位」
- 换汇成本 = (内贸金额 - 退税金额 + 运费 + 保费) / 总价（美金）

27 列布局（A~AA）：
 A提单号 B船次 C序号 D外贸合同号 E内贸合同号 F品名 G英文品名
 H打包后件数 I单位种类 J体积 K总毛重 L总净重 M总数量 N单位 O法定单位
 P单价（美金） Q总 价（美金） R换汇成本 S运费 T保费
 U报关单号 V出口/海关编码 W内贸金额 X退税率 Y退税金额 Z申报要素 AA供应商

公式（均为本行引用，Excel 会自动按行计算）：
- 单价（美金） P = Q / M            (=总价/总数量)
- 运费       S = MAX(ROUND(SUM(J),2), SUM(K)/1000)*15  (J=体积, K=总毛重)
- 保费       T = Q * 1.1 * 0.000145  (Q=总价)
- 退税金额   Y = W / 1.13 * X        (W=内贸金额, X=退税率)

合并规则：仅「打包后件数/单位种类/体积/总毛重」4 列做空格向上合并；船次不合并。
对齐：所有单元格 水平居中 + 垂直居中 + 自动换行。
"""

import re
from collections import OrderedDict
import pandas as pd
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
    "体积",            # J 10 合并（运费公式引用）
    "总毛重",          # K 11 合并（运费公式引用）
    "总净重",          # L 12
    "总数量",          # M 13（单价公式引用）
    "单位",            # N 14
    "法定单位",         # O 15
    "单价（美金）",      # P 16 公式 =Q/M
    "总 价（美金）",     # Q 17（保费/单价公式引用）
    "换汇成本",         # R 18
    "运费",            # S 19 公式
    "保费",            # T 20 公式 =Q*1.1*0.000145
    "报关单号",         # U 21
    "出口/海关编码",     # V 22
    "内贸金额",         # W 23（退税金额公式引用）
    "退税率",          # X 24（退税金额公式引用）
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
        freight = safe_float(get_val("运费"), default=None)   # S（数据表无则用公式）
        premium = safe_float(get_val("保费"), default=None)   # T（数据表无则用公式）

        # 退税金额：数据表有则透传，无则用公式 W/1.13*X
        tax_refund = safe_float(get_val("退税金额"))
        if tax_refund is None and irm is not None and trr is not None:
            tax_refund = round(irm / 1.13 * trr, 6)

        # 换汇成本
        cost = safe_float(get_val("换汇成本"))
        if cost is None and total_price and total_price > 0:
            numerator = (irm or 0) - (tax_refund or 0) + (freight or 0) + (premium or 0)
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
        row["体积"] = safe_float(get_val("体积"), default=None)      # J
        row["总毛重"] = safe_float(get_val("总毛重"), default=None)   # K
        row["总净重"] = safe_float(get_val("总净重"), default=None)
        row["总数量"] = safe_float(get_val("总数量"), default=None)   # M
        row["单位"] = get_val("单位", "") or ""
        row["法定单位"] = get_val("法定单位", "") or ""
        # P 单价（美金）保留值（如有），同时也会写公式兜底
        row["单价（美金）"] = safe_float(get_val("单价（美金）"), default=None)
        row["总 价（美金）"] = total_price                            # Q
        row["换汇成本"] = cost
        row["运费"] = None   # 公式填充
        row["保费"] = None   # 公式填充
        row["报关单号"] = get_val("报关单号", "") or ""
        row["出口/海关编码"] = get_val("出口/海关编码", "") or ""
        row["内贸金额"] = irm                                        # W
        row["退税率"] = trr                                           # X
        row["退税金额"] = tax_refund                                  # Y（值 + 公式）
        row["申报要素"] = get_val("申报要素", "") or ""
        row["供应商"] = get_val("供应商", "") or ""

        rows.append(row)

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


# =========================================================
# 导出 Excel：第1行空，第2行标题，第3行起数据
# - 仅 4 列空格向上合并，船次不合并
# - 全部居中 + 细边框
# - 单价(P)/运费(S)/保费(T)/退税金额(Y) 写公式
# =========================================================
def to_excel_with_layout(result_df, path):
    col_idx_map = {name: idx + 1 for idx, name in enumerate(result_df.columns)}

    # 列字母（按 OUTPUT_COLUMNS 顺序）
    def L(name):
        return get_column_letter(col_idx_map[name])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        result_df.to_excel(
            writer, index=False, header=False, startrow=2, sheet_name="汇总表"
        )
        ws = writer.sheets["汇总表"]

        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        border = Border(*(Side(style="thin"),) * 4)

        n = len(result_df)
        data_first = 3              # 数据起始行（Excel）
        data_last = 2 + n           # 数据结束行（Excel）

        # ---------- 标题行（第2行） ----------
        for col_idx, name in enumerate(result_df.columns, start=1):
            cell = ws.cell(row=2, column=col_idx, value=name)
            cell.alignment = center
            cell.font = Font(bold=True, name="Microsoft YaHei", size=10)
            cell.border = border

        # ---------- 数据行：写值 + 居中 + 边框 ----------
        for excel_row in range(data_first, data_last + 1):
            for col_idx in range(1, len(result_df.columns) + 1):
                cell = ws.cell(row=excel_row, column=col_idx)
                cell.alignment = center
                cell.font = Font(name="Microsoft YaHei", size=10)
                cell.border = border

        # ---------- 公式列（逐行，本行引用） ----------
        for excel_row in range(data_first, data_last + 1):
            # P 单价（美金） = 总价(Q) / 总数量(M)
            ws[f"{L('单价（美金）')}{excel_row}"] = (
                f"={L('总 价（美金）')}{excel_row}/{L('总数量')}{excel_row}"
            )
            # S 运费 = MAX(ROUND(SUM(J),2), SUM(K)/1000)*15
            ws[f"{L('运费')}{excel_row}"] = (
                f"=MAX(ROUND(SUM({L('体积')}{excel_row}),2),"
                f"SUM({L('总毛重')}{excel_row})/1000)*15"
            )
            # T 保费 = 总价(Q) * 1.1 * 0.000145
            ws[f"{L('保费')}{excel_row}"] = (
                f"={L('总 价（美金）')}{excel_row}*1.1*0.000145"
            )
            # Y 退税金额 = 内贸金额(W) / 1.13 * 退税率(X)
            ws[f"{L('退税金额')}{excel_row}"] = (
                f"={L('内贸金额')}{excel_row}/1.13*{L('退税率')}{excel_row}"
            )

        # ---------- 空格向上合并（仅 MERGE_COLUMNS 4 列） ----------
        for col_name in MERGE_COLUMNS:
            _merge_consecutive_blanks(ws, col_idx_map[col_name], data_first, data_last)

        # ---------- 自动列宽 ----------
        for col_idx, name in enumerate(result_df.columns, start=1):
            letter = get_column_letter(col_idx)
            max_len = len(str(name))
            for row_idx in range(data_first, min(data_last + 1, data_first + 500)):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)[:40]))
            ws.column_dimensions[letter].width = min(max_len + 4, 42)

        # 行高
        ws.row_dimensions[2].height = 32
        for row_idx in range(data_first, data_last + 1):
            ws.row_dimensions[row_idx].height = 42

    return path


def _merge_consecutive_blanks(ws, col_idx, start, end):
    """
    对指定列，将每一段「连续为空的单元格」向上合并到该段的第一个非空单元格下方。
    规则：从第一个非空值之后开始，连续的空格合并成一个（合并到该段起始处）。
    具体：遇到非空值 A，其后连续空格合并到 A 所在单元格（即空格向上并入最近的非空值）。
    """
    # 找到每个非空值的位置
    nonempty_rows = [r for r in range(start, end + 1)
                     if not _cell_is_empty(ws.cell(row=r, column=col_idx).value)]
    if not nonempty_rows:
        return

    # 在每个非空值与其下一个非空值之间，若有空格则合并
    anchor = start  # 第一段从 start 开始（若开头就是空格，并入第一个非空值）
    # 在开头插入一个虚拟锚点，便于统一处理
    boundaries = [start - 1] + nonempty_rows + [end + 1]

    for i in range(1, len(boundaries) - 1):
        block_start = boundaries[i]      # 非空值行
        block_end = boundaries[i + 1] - 1  # 该非空值后连续空格的末尾
        if block_end > block_start:
            # block_start(非空) ~ block_end(最后一段空格) 合并
            ws.merge_cells(
                start_row=block_start, start_column=col_idx,
                end_row=block_end, end_column=col_idx,
            )


def transform(df):
    """兼容旧接口"""
    return convert_table(df)

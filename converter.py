"""
Excel 明细表（数据表）-> 汇总表（表一格式）转换逻辑

核心规则：
- 数据表第1行空，第2行标题 → 读取用 header=1
- 每个「序号」恰好一行（51序号 = 51行），无需分组聚合
- 「单位种类.1」= 报关后单位（件/个/套/台/吨）→ 输出「单位种类」
- 「报关单位」= 法定第一单位 → 输出「单位」
- 单价（美金）= 总价（美金） / 总数量（自动计算）
- 退税金额 = 内贸金额 / 1.13 × 退税率（自动计算）
- 保费 = 总价（美金） × 保险费率 0.01595%（自动计算）
- 运费 = 公式 MAX(ROUND(SUM(J行),2), SUM(K行)/1000)*15

输出格式：
- 第1行空，第2行标题，第3行起数据，按序号升序，共 27 列
- 空单元格合并（向上合并同类）
- 所有单元格：水平居中 + 垂直居中 + 自动换行
"""

import re
from collections import OrderedDict
import pandas as pd
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter


# =========================================================
# 可配置参数
# =========================================================
INSURANCE_RATE = 0.0001595   # 保险费率


# =========================================================
# 输出列定义（27列，完全对齐表一）
# =========================================================
OUTPUT_COLUMNS = [
    "提单号",
    "船次",
    "序号",
    "外贸合同号",
    "内贸合同号",
    "品名",
    "英文品名",
    "打包后件数",
    "单位种类",
    "体积",
    "总毛重",
    "总净重",
    "总数量",
    "单位",
    "法定单位",
    "单价（美金）",
    "总 价（美金）",
    "换汇成本",
    "运费",
    "保费",
    "报关单号",
    "出口/海关编码",
    "内贸金额",
    "退税率",
    "退税金额",
    "申报要素",
    "供应商",
]

# 需要合并的列（空单元格向上合并到最近的非空单元格）
MERGE_COLUMNS = [
    "提单号",
    "船次",
    "外贸合同号",
    "内贸合同号",
    "品名",
    "英文品名",
    "申报要素",
    "供应商",
]

NUMERIC_COLUMNS = {
    "序号", "打包后件数", "体积", "总毛重", "总净重", "总数量",
    "单价（美金）", "总 价（美金）", "换汇成本", "运费", "保费",
    "出口/海关编码", "内贸金额", "退税率", "退税金额",
}

# 输出列 -> 数据表真实列的模糊候选
COLUMN_MAP = {
    "提单号":       ["提单号"],
    "船次":         [],
    "序号":         ["序号", "商品序号"],
    "外贸合同号":    ["外贸合同号"],
    "内贸合同号":    ["内贸合同号"],
    "品名":         ["品名"],
    "英文品名":      ["英文品名"],
    "打包后件数":    ["打包后件数", "打包后 件数", "打包后\n件数"],
    "单位种类":      ["单位种类.1", "单位 种类.1", "单位种类", "种类"],
    "体积":         ["体积"],
    "总毛重":       ["总毛重"],
    "总净重":       ["总净重"],
    "总数量":       ["总数量", "货物最小单位数量"],
    "单位":         ["报关单位", "单位"],
    "法定单位":      ["法定单位"],
    "单价（美金）":   [],
    "总 价（美金）":  ["总价（美金）", "总 价（美金）", "总 价\n（美金）", "总价"],
    "换汇成本":      ["换汇成本"],
    "运费":         ["运费"],
    "保费":         ["保费"],
    "报关单号":      ["报关单号", "提运单号"],
    "出口/海关编码":  ["出口/海关编码", "海关编码", "出口/海关编码\n（供退税用）", "出口/海关编码\n（HS码）"],
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


# =========================================================
# 主转换（只接 DataFrame，不接 source_path）
# =========================================================
def convert_table(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_map = {}
    for orig in df.columns:
        nc = clean_col_name(orig)
        if nc and nc not in cleaned_map:
            cleaned_map[nc] = orig

    df = df.copy()
    df.columns = [cleaned_map.get(clean_col_name(c), c) for c in df.columns]

    seq_real = find_real_column(cleaned_map, ["序号", "商品序号"])
    if seq_real is None:
        raise ValueError(f"未找到「序号」列。实际列名：{list(cleaned_map.values())}")

    df["__seq"] = df[seq_real].apply(lambda x: safe_int(x, default=None))
    df = df[df["__seq"].notna()].copy()
    # 只保留序号 1~100（跳过后面配件明细如606）
    df = df[(df["__seq"] >= 1) & (df["__seq"] <= 100)].copy()
    if len(df) == 0:
        raise ValueError("序号列没有有效数值。")
    df = df.sort_values("__seq").reset_index(drop=True)

    real_map = {}
    for out_col, candidates in COLUMN_MAP.items():
        real = find_real_column(cleaned_map, candidates)
        if real:
            real_map[out_col] = real

    # 若「总数量」没匹配到，用「货物最小单位数量」兜底
    if "总数量" not in real_map:
        for orig in df.columns:
            if clean_col_name(orig) == "货物最小单位数量":
                real_map["总数量"] = orig
                break

    rows = []
    for _, r in df.iterrows():
        def get_val(out_col, default=None):
            real = real_map.get(out_col)
            if not real:
                return default
            v = r[real]
            if pd.isna(v) or (isinstance(v, str) and v.strip() == ""):
                return default
            return v

        total_price = safe_float(get_val("总 价（美金）"))
        total_qty = safe_float(get_val("总数量"))
        irm = safe_float(get_val("内贸金额"))
        trr = safe_float(get_val("退税率"))

        # 退税金额（自动计算）
        tax_refund = safe_float(get_val("退税金额"))
        if tax_refund is None and irm is not None and trr is not None:
            tax_refund = round(irm / 1.13 * trr, 6)

        # 保费（自动计算）
        premium = safe_float(get_val("保费"))
        if premium is None and total_price:
            premium = round(total_price * INSURANCE_RATE, 8)

        # 单价（美金）
        unit_price = round(total_price / total_qty, 4) if total_qty else None

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
        row["单价（美金）"] = unit_price
        row["总 价（美金）"] = total_price
        row["换汇成本"] = safe_float(get_val("换汇成本"), default=None)
        row["运费"] = None   # 由公式填充
        row["保费"] = premium
        row["报关单号"] = get_val("报关单号", "") or ""
        row["出口/海关编码"] = safe_float(get_val("出口/海关编码"), default=None)
        row["内贸金额"] = irm
        row["退税率"] = trr
        row["退税金额"] = tax_refund
        row["申报要素"] = get_val("申报要素", "") or ""
        row["供应商"] = get_val("供应商", "") or ""

        rows.append(row)

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


# =========================================================
# 导出 Excel（第1行空，第2行标题，第3行起数据）
# 格式：空单元格合并 + 全部居中 + 运费公式
# =========================================================
def to_excel_with_layout(result_df, path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        result_df.to_excel(
            writer, index=False, header=False, startrow=2, sheet_name="汇总表"
        )
        ws = writer.sheets["汇总表"]

        # ---- 写标题（第2行）----
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        for col_idx, name in enumerate(result_df.columns, start=1):
            cell = ws.cell(row=2, column=col_idx, value=name)
            cell.alignment = center_align
            cell.border = thin_border
            cell.font = Font(bold=True, name="Microsoft YaHei", size=10)

        # ---- 写数据 + 居中 + 运费公式 ----
        for row_idx in range(3, 3 + len(result_df)):
            excel_row = row_idx
            for col_idx, col_name in enumerate(result_df.columns, start=1):
                cell = ws.cell(row=excel_row, column=col_idx)
                cell.alignment = center_align
                cell.border = thin_border
                cell.font = Font(name="Microsoft YaHei", size=10)

                # 运费列（第19列 = "运费"）写公式
                if col_name == "运费":
                    # J = 体积(第10列), K = 总毛重(第11列)
                    formula = (
                        f"=MAX(ROUND(SUM(J{excel_row}),2),"
                        f"SUM(K{excel_row})/1000)*15"
                    )
                    cell.value = formula

        # ---- 合并空单元格（按列，向上合并连续相同/空值）----
        col_index_map = {name: idx + 1 for idx, name in enumerate(result_df.columns)}

        for col_name in MERGE_COLUMNS:
            if col_name not in col_index_map:
                continue
            col_idx = col_index_map[col_name]
            start_row = 3
            end_row = 2 + len(result_df)

            merge_start = start_row
            for r in range(start_row, end_row + 1):
                current_val = ws.cell(row=r, column=col_idx).value
                next_val = ws.cell(row=r + 1, column=col_idx).value if r < end_row else None

                if r == end_row:
                    if r > merge_start:
                        ws.merge_cells(
                            start_row=merge_start, start_column=col_idx,
                            end_row=r, end_column=col_idx
                        )
                elif _is_same_value(current_val, next_val):
                    continue
                else:
                    if r > merge_start:
                        ws.merge_cells(
                            start_row=merge_start, start_column=col_idx,
                            end_row=r, end_column=col_idx
                        )
                    merge_start = r + 1

        # ---- 自动列宽 ----
        for col_idx, name in enumerate(result_df.columns, start=1):
            letter = get_column_letter(col_idx)
            max_len = len(str(name))
            for row_idx in range(3, min(3 + len(result_df), 500)):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)[:40]))
            ws.column_dimensions[letter].width = min(max_len + 4, 40)

        # 第1行留空
        ws.row_dimensions[2].height = 30  # 标题行
        for row_idx in range(3, 3 + len(result_df)):
            ws.row_dimensions[row_idx].height = 40  # 数据行

    return path


def _is_same_value(a, b):
    """判断两个单元格值是否"相同"（都空也算相同，用于合并）"""
    a_empty = a is None or (isinstance(a, str) and a.strip() == "")
    b_empty = b is None or (isinstance(b, str) and b.strip() == "")
    if a_empty and b_empty:
        return True
    if a_empty or b_empty:
        return False
    return str(a).strip() == str(b).strip()


def transform(df):
    """兼容旧接口"""
    return convert_table(df)

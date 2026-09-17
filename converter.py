"""
Excel 明细表（表一）-> 汇总表（表二）转换逻辑

关键说明：
- 表一第1行空，第2行是标题（读取时用 header=1）
- 表一每列名可能含换行/多余空格，如「打包后\n件数」「总 价\n（美金）」，需清洗匹配
- 每个「序号」在表一中基本是一行，文本字段取首值、数值字段取首值
- 输出：第1行空，第2行标题，第3行起数据，按序号升序
"""

import re
from collections import OrderedDict
import pandas as pd


# =========================================================
# 1. 输出列定义（20列，严格按用户指定顺序）
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
]

# 输出列 -> 表一真实列的模糊候选（按优先级匹配）
COLUMN_MAP = {
    "提单号":       ["提单号"],
    "船次":         [],                        # 固定值 989船
    "序号":         ["序号"],
    "外贸合同号":    ["外贸合同号"],
    "内贸合同号":    ["内贸合同号"],
    "品名":         ["品名"],
    "英文品名":      ["英文品名"],
    "打包后件数":    ["打包后件数", "打包后 件数", "打包后\n件数"],
    "单位种类":      ["报关单位", "单位种类.1", "单位种类", "单位 种类", "单位\n种类.1"],
    "体积":         ["体积"],
    "总毛重":       ["总毛重"],
    "总净重":       ["总净重"],
    "总数量":       ["总数量"],
    "单位":         ["报关单位", "单位种类", "单位 种类", "货物最小单位数量"],
    "法定单位":      ["法定单位"],
    "单价（美金）":   [],                        # 由 总价/数量 计算
    "总 价（美金）":  ["总价（美金）", "总价 （美金）", "总 价（美金）", "总 价 （美金）", "总价"],
    "换汇成本":      ["换汇成本"],
    "运费":         [],
    "保费":         [],
}


# =========================================================
# 2. 工具函数
# =========================================================
def safe_float(x, default=0.0):
    """安全转 float，空字符串/None/NaN/占位符 -> default"""
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


def safe_int(x, default=0):
    v = safe_float(x, default)
    if v is None or v == 0:
        return default
    return int(v)


def clean_col_name(c):
    """列名归一化：去换行、合并内部空白、去两端空格。"""
    if pd.isna(c) or str(c).strip() == "":
        return ""
    return re.sub(r"\s+", "", str(c)).strip()


def find_real_column(df_columns, candidates):
    """在真实列名（已清洗）中按候选列表顺序匹配。"""
    if not candidates:
        return None
    cleaned = {clean_col_name(c): c for c in df_columns if clean_col_name(c)}

    # 1) 精确匹配
    for cand in candidates:
        ncand = clean_col_name(cand)
        if ncand in cleaned:
            return cleaned[ncand]

    # 2) 包含匹配（去掉括号干扰）
    def strip_symbol(s):
        return re.sub(r"[（）()\s]", "", s)

    for cand in candidates:
        k = strip_symbol(clean_col_name(cand))
        if not k:
            continue
        for key, raw in cleaned.items():
            if k in strip_symbol(key):
                return raw

    return None


# =========================================================
# 3. 主转换
# =========================================================
def convert_table(df: pd.DataFrame) -> pd.DataFrame:
    orig_columns = list(df.columns)
    df = df.copy()
    df.columns = [clean_col_name(c) for c in df.columns]

    # 找序号列
    seq_real = find_real_column(df.columns, ["序号"])
    if seq_real is None:
        raise ValueError(f"未找到「序号」列。实际列名：{list(orig_columns)}")

    # 提取有效序号行
    df["__seq"] = df[seq_real].apply(lambda x: safe_int(x, default=None))
    df = df.dropna(subset=["__seq"]).copy()
    if len(df) == 0:
        raise ValueError("序号列没有有效数值，请检查表一数据。")
    df["__seq"] = df["__seq"].astype(int)

    # 向下填充：关键字段在主件行填了，配件明细行要继承
    ffill_cols = []
    for cand_list in COLUMN_MAP.values():
        for c in cand_list:
            real = find_real_column(df.columns, [c])
            if real and real not in ffill_cols:
                ffill_cols.append(real)
    for c in df.columns:
        if c in [clean_col_name(orig) for orig in ffill_cols]:
            df[c] = df[c].ffill()

    # 预解析每个输出列对应的真实列
    real_map = {}
    for out_col, candidates in COLUMN_MAP.items():
        real = find_real_column(df.columns, candidates)
        if real:
            real_map[out_col] = real

    rows = []
    for seq, group in df.groupby("__seq", sort=True):
        def first_value(out_col):
            real = real_map.get(out_col)
            if not real:
                return None
            for v in group[real].dropna():
                if isinstance(v, str) and v.strip() == "":
                    continue
                return v
            return None

        row = OrderedDict()
        row["序号"] = int(seq)
        row["提单号"] = first_value("提单号") or ""
        row["船次"] = "989船"
        row["外贸合同号"] = first_value("外贸合同号") or ""
        row["内贸合同号"] = first_value("内贸合同号") or ""
        row["品名"] = first_value("品名") or ""
        row["英文品名"] = first_value("英文品名") or ""
        row["打包后件数"] = first_value("打包后件数")
        row["单位种类"] = first_value("单位种类") or ""
        row["体积"] = safe_float(first_value("体积"))
        row["总毛重"] = safe_float(first_value("总毛重"))
        row["总净重"] = safe_float(first_value("总净重"))
        row["总数量"] = safe_float(first_value("总数量"))
        row["单位"] = first_value("单位") or ""
        row["法定单位"] = first_value("法定单位") or ""

        total_price = safe_float(first_value("总 价（美金）"))
        total_qty = safe_float(first_value("总数量"))
        row["总 价（美金）"] = total_price
        row["单价（美金）"] = round(total_price / total_qty, 4) if total_qty else 0.0
        row["换汇成本"] = safe_float(first_value("换汇成本"))
        row["运费"] = safe_float(first_value("运费"))
        row["保费"] = safe_float(first_value("保费"))

        rows.append(row)

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


# =========================================================
# 4. 导出 Excel（第1行空，第2行标题，第3行起数据）
# =========================================================
def to_excel_with_layout(result_df, path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        result_df.to_excel(
            writer, index=False, header=False, startrow=2, sheet_name="汇总表"
        )
        ws = writer.sheets["汇总表"]
        # 第2行写标题
        for col_idx, name in enumerate(result_df.columns, start=1):
            ws.cell(row=2, column=col_idx, value=name)
        # 第1行留空（不写内容）
        # 自动列宽
        for col_idx, name in enumerate(result_df.columns, start=1):
            letter = ws.cell(row=2, column=col_idx).column_letter
            max_len = len(str(name))
            for row in range(3, min(3 + len(result_df), 500)):
                val = ws.cell(row=row, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)[:25]))
            ws.column_dimensions[letter].width = min(max_len + 4, 28)
    return path


def transform(df):
    """兼容旧接口"""
    return convert_table(df)

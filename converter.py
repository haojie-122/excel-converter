"""
Excel 明细表（数据表）-> 汇总表（表一格式）转换逻辑

核心规则：
- 数据表第1行空，第2行标题 → 读取用 header=1
- 每个「序号」恰好一行（51序号 = 51行），无需分组聚合
- 「单位种类.1」= 报关后单位（件/个/套/台/吨/米）→ 输出「单位种类」
- 「报关单位」= 法定第一单位（台/个/套）→ 输出「单位」
- 配件行（如序号4球阀）的 件数/体积/毛重 本就为空 → 输出也留空
- 单价（美金）= 总价（美金） / 总数量（自动计算）
- 退税金额 = 内贸金额 / 1.13 × 退税率（自动计算）
- 保费 = 总价（美金） × 保险费率 0.01595%（自动计算）
- 运费：数据表中无此列，由人工填写（代码保留数据表值，无则留空）

输出：第1行空，第2行标题，第3行起数据，按序号升序，共 27 列
"""

import re
from collections import OrderedDict
import pandas as pd


# =========================================================
# 可配置参数
# =========================================================
INSURANCE_RATE = 0.0001595   # 保险费率：保费 = 美金总价 × 此比例
# 运费：数据表中无此列，由人工填写；如需自动计算取消下面注释并调整费率
# FREIGHT_RATE = 0.03  # 元/kg，运费 = 总毛重 × FREIGHT_RATE


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
    "单位种类":      ["单位种类.1", "单位 种类.1"],       # 报关后单位
    "体积":         ["体积"],
    "总毛重":       ["总毛重"],
    "总净重":       ["总净重"],
    "总数量":       ["总数量"],
    "单位":         ["报关单位"],                  # 法定第一单位
    "法定单位":      ["法定单位"],
    "单价（美金）":   [],                           # 自动计算
    "总 价（美金）":  ["总价（美金）", "总 价（美金）", "总价"],
    "换汇成本":      ["换汇成本"],
    "运费":         ["运费"],                      # 无则留空
    "保费":         ["保费"],                      # 无则按保险率算
    "报关单号":      ["报关单号"],
    "出口/海关编码":  ["出口/海关编码", "海关编码"],
    "内贸金额":      ["内贸金额"],
    "退税率":       ["退税率"],
    "退税金额":      ["退税金额"],                  # 自动计算
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
    """cleaned_map: {清洗后列名: 原始列名}"""
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
# 主转换
# =========================================================
def convert_table(df: pd.DataFrame) -> pd.DataFrame:
    # 建立清洗映射
    cleaned_map = {}
    for orig in df.columns:
        nc = clean_col_name(orig)
        if nc and nc not in cleaned_map:
            cleaned_map[nc] = orig

    df = df.copy()
    df.columns = [cleaned_map.get(clean_col_name(c), c) for c in df.columns]

    # 序号列
    seq_real = find_real_column(cleaned_map, ["序号"])
    if seq_real is None:
        raise ValueError(f"未找到「序号」列。实际列名：{list(cleaned_map.values())}")

    df["__seq"] = df[seq_real].apply(lambda x: safe_int(x, default=None))
    df = df.dropna(subset=["__seq"]).copy()
    if len(df) == 0:
        raise ValueError("序号列没有有效数值。")
    df["__seq"] = df["__seq"].astype(int)

    # 预解析列映射
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
            if pd.isna(v) or (isinstance(v, str) and v.strip() == ""):
                return default
            return v

        # 基础数值
        total_price = safe_float(get_val("总 价（美金）"))
        total_qty = safe_float(get_val("总数量"))
        irm = safe_float(get_val("内贸金额"))
        trr = safe_float(get_val("退税率"))
        gross = safe_float(get_val("总毛重"))

        # 退税金额（自动计算）
        tax_refund = safe_float(get_val("退税金额"))
        if tax_refund is None and irm is not None and trr is not None:
            tax_refund = round(irm / 1.13 * trr, 6)

        # 运费（数据表无则留空，由人工填写）
        freight = safe_float(get_val("运费"), default=None)

        # 保费（数据表有则透传，无则按保险率算）
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
        row["打包后件数"] = safe_int(get_val("打包后件数"), default=None)  # 空则留空
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
        row["运费"] = freight
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
        # 第1行留空
        # 自动列宽
        for col_idx, name in enumerate(result_df.columns, start=1):
            letter = ws.cell(row=2, column=col_idx).column_letter
            max_len = len(str(name))
            for row in range(3, min(3 + len(result_df), 500)):
                val = ws.cell(row=row, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)[:30]))
            ws.column_dimensions[letter].width = min(max_len + 4, 35)
    return path


def transform(df):
    """兼容旧接口"""
    return convert_table(df)

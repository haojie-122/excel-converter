import pandas as pd
from collections import OrderedDict

NUMERIC_HINTS = [
    "体积", "毛重", "净重", "数量", "件", "单价", "总价", "价（美金）",
    "美金", "换汇", "成本", "运费", "保费", "序号"
]

MONEY_COLS = ["单价（美金）", "总 价（美金）", "总价（美金）", "换汇成本", "运费", "保费"]
QTY_COLS = ["体积", "总毛重", "总净重", "总数量", "打包后件"]
INT_HINTS = ["序号", "打包后件", "数量"]


def clean_col(c):
    if pd.isna(c):
        return ""
    return str(c).replace("\n", "").replace("\r", "").replace("\t", "").strip()


def safe_float(x, default=None):
    if x is None:
        return default
    if isinstance(x, float) and pd.isna(x):
        return default
    s = str(x).replace(",", "").replace("，", "").strip()
    if s in ("", "-", "—", "–", "null", "None", "NA", "N/A", "无", "空"):
        return default
    try:
        return float(s)
    except Exception:
        return default


def safe_int(x, default=None):
    v = safe_float(x, default)
    if v is None:
        return default
    if float(v) == int(v):
        return int(v)
    return int(v)


def looks_numeric(col):
    return any(k in col for k in NUMERIC_HINTS)


def first_non_empty(series):
    for x in series:
        if x is None:
            continue
        if isinstance(x, float) and pd.isna(x):
            continue
        s = str(x).strip()
        if s in ("", "-", "—", "–"):
            continue
        return x
    return ""


def normalize_columns(df):
    # 假设第2行是标题，第1行空；调用前可用 header=1
    df.columns = [clean_col(c) for c in df.columns]
    # 去掉完全空列名
    df = df.loc[:, df.columns != ""]
    return df


def convert_table(df):
    df = normalize_columns(df)

    # 调试：可以在桌面弹窗/命令行看到实际列名
    print("实际列名:", list(df.columns))

    # 序号处理
    if "序号" not in df.columns:
        raise ValueError(f"未找到“序号”列。实际列名：{list(df.columns)}")

    df["序号_数值"] = df["序号"].apply(lambda x: safe_int(x, default=None))
    df = df.dropna(subset=["序号_数值"]).copy()
    df["序号_数值"] = df["序号_数值"].astype(int)

    # 数字列安全转换
    for c in df.columns:
        if looks_numeric(c) and c != "序号":
            df[c + "_num"] = df[c].apply(lambda x: safe_float(x, 0))

    grouped = df.groupby("序号_数值", sort=True)

    out_rows = []
    for seq, g in grouped:
        row = OrderedDict()

        # 基础/文本字段优先保留
        text_pref_cols = ["提单号", "船次", "序号", "外贸合同", "内贸合同",
                          "品名", "英文品名", "单位", "种类", "法定单位",
                          "打包后件"]
        for c in text_pref_cols:
            if c in df.columns:
                row[c] = first_non_empty(g[c])
        if not row.get("序号"):
            row["序号"] = seq

        # 其他非数字列也保留首值
        for c in df.columns:
            if c in row:
                continue
            if c.endswith("_num"):
                continue
            if looks_numeric(c):
                continue
            if c == "序号_数值":
                continue
            row[c] = first_non_empty(g[c])

        # 数字列：数量/重量/体积/件数求和；金额类求和；其他可取首值
        for c in df.columns:
            if not looks_numeric(c) or c == "序号":
                continue
            num_col = c + "_num"
            if num_col in df.columns:
                v = g[num_col].sum()
            else:
                v = safe_float(first_non_empty(g[c]), 0)

            if any(k in c for k in ["单价", "换汇", "成本", "运费", "保费", "总价", "价（美金）"]):
                v = safe_float(first_non_empty(g[c]), 0)
                if v is None:
                    v = 0

            if any(k in c for k in INT_HINTS) and c not in MONEY_COLS:
                if float(v) == int(v):
                    row[c] = int(v) if v != 0 else ""
                else:
                    row[c] = v
            else:
                row[c] = v if v != 0 else ""

        out_rows.append(row)

    result = pd.DataFrame(out_rows)
    return result


def to_excel_with_layout(df, save_path):
    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=1)
        ws = writer.sheets["Sheet1"]
        ws.row_dimensions[1].height = 18
    return save_path

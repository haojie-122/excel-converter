import pandas as pd


def find_seq_col(columns):
    for col in columns:
        if "序号" in str(col):
            return col
    return columns[0]


def convert_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    seq_col = find_seq_col(df.columns)
    df = df.rename(columns={seq_col: "序号"})
    df = df[df["序号"].notna()]

    # 数值列：求和
    sum_cols = []
    for col in df.columns:
        for kw in ["体积", "毛重", "净重", "数量", "件数", "总价", "金额"]:
            if kw in str(col):
                sum_cols.append(col)
                break

    first_cols = [c for c in df.columns if c not in sum_cols + ["序号"]]

    agg = {}
    for c in sum_cols:
        agg[c] = "sum"
    for c in first_cols:
        agg[c] = "first"

    result = df.groupby("序号", as_index=False).agg(agg)

    # 填充船次
    if "船次" not in result.columns:
        result["船次"] = "989船"

    return result

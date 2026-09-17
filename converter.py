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

    # 按你指定的列顺序输出（20列）
    output_columns = [
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

    for col in output_columns:
        if col not in result.columns:
            result[col] = ""

    result = result[output_columns]

    # 计算单价
    if "总 价（美金）" in result.columns and "总数量" in result.columns:
        result["单价（美金）"] = (
            result["总 价（美金）"].astype(float)
            / result["总数量"].replace(0, pd.NA)
        ).round(4)

    # 填充船次
    result["船次"] = "989船"

    # 按序号升序排列
    result = result.sort_values("序号").reset_index(drop=True)

    return result


def to_excel_with_layout(result_df, path):
    """
    第1行空，第2行标题，第3行起数据，按序号升序
    """
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # 数据从第3行开始写（startrow=2 表示第3行，0-based）
        result_df.to_excel(
            writer,
            index=False,
            header=False,
            startrow=2,
            sheet_name="汇总表",
        )
        ws = writer.sheets["汇总表"]

        # 第2行写标题
        for col_idx, name in enumerate(result_df.columns, start=1):
            ws.cell(row=2, column=col_idx, value=name)

        # 第1行不写任何内容（自动为空）

        # 自动调整列宽
        for col_idx, name in enumerate(result_df.columns, start=1):
            max_len = len(str(name))
            for row in range(3, min(3 + len(result_df), 53)):
                val = ws.cell(row=row, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[ws.cell(row=2, column=col_idx).column_letter].width = min(max_len + 4, 30)

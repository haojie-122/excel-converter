import pandas as pd


def safe_float(x, default=0.0):
    if x is None:
        return default
    if isinstance(x, float) and pd.isna(x):
        return default
    if isinstance(x, str):
        s = x.strip()
        if s in ('', '-', '—', '–', 'null', 'None', 'NA', 'N/A', '无', '空'):
            return default
    try:
        return float(x)
    except (ValueError, TypeError):
        return default


def safe_int(x, default=0):
    v = safe_float(x, default)
    if v is None:
        return default
    return int(v)


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

    # 数值列列表
    num_cols = [
        "体积", "总毛重", "总净重", "总数量",
        "单价（美金）", "总 价（美金）", "换汇成本",
        "运费", "保费", "打包后件数",
    ]

    # 自动识别更多数值列（含"金额""件数""数量"等关键词）
    for col in df.columns:
        for kw in ["金额", "件数", "数量", "重量", "体积", "总价", "单价"]:
            if kw in str(col) and col not in num_cols and col != "序号":
                num_cols.append(col)
                break

    # 所有数值列先做安全转换
    for c in num_cols:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: safe_float(x, 0.0))

    # 序号列也做安全转换
    df["序号"] = df["序号"].apply(lambda x: safe_int(x, 0))

    # 去掉序号为0的行
    df = df[df["序号"] > 0]

    # 数值列：求和
    sum_cols = [c for c in num_cols if c in df.columns]

    # 非数值列：取每组第一个非空值
    first_cols = [c for c in df.columns if c not in sum_cols + ["序号"]]

    agg = {}
    for c in sum_cols:
        agg[c] = "sum"
    for c in first_cols:
        agg[c] = "first"

    result = df.groupby("序号", as_index=False).agg(agg)

    # 按指定列顺序输出（20列）
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

    # 计算单价（总价÷数量）
    result["单价（美金）"] = result.apply(
        lambda row: round(safe_float(row["总 价（美金）"], 0) / safe_float(row["总数量"], 1), 4)
        if safe_float(row["总数量"], 0) > 0
        else 0.0,
        axis=1,
    )

    # 填充船次
    result["船次"] = "989船"

    # 按序号升序
    result = result.sort_values("序号").reset_index(drop=True)

    return result


def to_excel_with_layout(result_df, path):
    """
    第1行空，第2行标题，第3行起数据，按序号升序
    """
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
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

        # 第1行空（不写内容）

        # 自动调整列宽
        for col_idx, name in enumerate(result_df.columns, start=1):
            max_len = len(str(name))
            for row in range(3, min(3 + len(result_df), 200)):
                val = ws.cell(row=row, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)[:20]))
            ws.column_dimensions[
                ws.cell(row=2, column=col_idx).column_letter
            ].width = min(max_len + 4, 30)

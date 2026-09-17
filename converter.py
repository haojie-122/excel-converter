"""
Excel 明细表（整合.xlsx）-> 汇总表（表一格式）转换逻辑

============================================================
【源表结构】(已通过原始单元格逐一核实)
- 第1行：空
- 第2行：标题行(合并单元格多，表头名不能直接作列索引)
- 第3行起：数据，每个 序号(1~51) = 一个商品
- 序号列(col3)有值 = 一条有效记录

【真实物理列映射】(1-based, 依据第3行原始单元格核实)
   col1=提运单号(报关单号)   col2=船次(989船)      col3=序号
   col4=外贸合同号           col5=内贸合同号       col6=品名
   col7=英文品名             col8=打包后件数       col9=(件数补充, 通常空)
   col10=体积                col11=总毛重          col12=总净重
   col13=总数量              col14=单位种类(报关后) col15=报关单位
   col16=(单价公式文本)       col17=(体积溢出,忽略)  col18=总 价（美金）(=Q)
   col19=(运费公式)          col20=(保费公式)       col21=法定单位
   col22=出口/海关编码        col23=内贸金额(=W)    col24=退税率(=X)
   col25=(退税金额公式)       col26=申报要素        col27=供应商

【输出 27 列 A~AA】
 A提单号 B船次 C序号 D外贸合同号 E内贸合同号 F品名 G英文品名
 H打包后件数 I单位种类 J体积 K总毛重 L总净重 M总数量 N单位 O法定单位
 P单价（美金） Q总 价（美金） R换汇成本 S运费 T保费
 U报关单号 V出口/海关编码 W内贸金额 X退税率 Y退税金额 Z申报要素 AA供应商

【公式】(本行引用, Excel 自动按行计算)
 P(单价) = Q/M
 S(运费) = MAX(ROUND(SUM(J),2), SUM(K)/1000)*15    J=体积, K=总毛重
 T(保费) = Q*1.1*0.000145
 Y(退税) = W/1.13*X
 R(换汇) = (W - Y + S + T)/Q

【格式】
- 第1行空, 第2行标题, 第3行起数据, 按序号升序
- 仅「打包后件数/单位种类/体积/总毛重」4列空格向上合并；船次不合并
- 全部：水平居中 + 垂直居中 + 自动换行
============================================================
"""

import pandas as pd
from collections import OrderedDict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter


# =========================================================
# 可配置参数
# =========================================================
MAX_SEQ = 51        # 只保留 序号 1~51 的主商品（606 等为配件明细）
DATA_START_ROW = 3  # 数据起始行(Excel)


# =========================================================
# 输出列（27列）
# =========================================================
OUTPUT_COLUMNS = [
    "提单号", "船次", "序号", "外贸合同号", "内贸合同号",
    "品名", "英文品名", "打包后件数", "单位种类", "体积", "总毛重",
    "总净重", "总数量", "单位", "法定单位", "单价（美金）", "总 价（美金）",
    "换汇成本", "运费", "保费", "报关单号", "出口/海关编码",
    "内贸金额", "退税率", "退税金额", "申报要素", "供应商",
]

LETTERS = [get_column_letter(i + 1) for i in range(len(OUTPUT_COLUMNS))]


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


def _is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _flatten_merges(ws):
    """
    扁平化合并单元格：先把每个合并区的值存入其覆盖的所有单元格，再解除合并。
    解除后所有单元格变为可写，值得以保留。
    """
    # 先记录每个合并区的值与覆盖范围
    info = []
    for rng in list(ws.merged_cells.ranges):
        top_left = ws.cell(row=rng.min_row, column=rng.min_col).value
        info.append((rng, top_left))

    # 解除所有合并
    for rng, _ in info:
        ws.unmerge_cells(str(rng))

    # 填充值到每个覆盖单元格（此时均已可写）
    for rng, value in info:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                ws.cell(row=r, column=c).value = value


# =========================================================
# 主转换
# =========================================================
def convert_table(df=None, source_path=None) -> pd.DataFrame:
    if source_path is None:
        raise ValueError("本实现需提供 source_path（源 xlsx 路径）")

    wb = load_workbook(source_path, data_only=True)
    ws = wb.active

    # 关键：先扁平化合并单元格
    _flatten_merges(ws)

    rows = []
    for r in range(DATA_START_ROW, ws.max_row + 1):
        seq = safe_int(ws.cell(row=r, column=3).value, default=None)  # col3 序号
        if seq is None or seq <= 0 or seq > MAX_SEQ:
            continue  # 只保留 序号 1~51

        def g(col, default=None):
            v = ws.cell(row=r, column=col).value
            return v if not _is_empty(v) else default

        total_price = safe_float(g(18))   # Q 总价（美金）
        total_qty = safe_float(g(13))     # M 总数量
        irm = safe_float(g(23))           # W 内贸金额
        trr = safe_float(g(24))           # X 退税率

        # 退税金额 = 内贸金额 / 1.13 * 退税率
        tax_refund = None
        if irm is not None and trr is not None:
            tax_refund = round(irm / 1.13 * trr, 6)

        # 换汇成本 = (内贸金额 - 退税金额 + 运费 + 保费) / 总价（美金）
        # 运费/保费按与主表一致公式计算（python 预填，Excel 公式为准）
        cost = None
        if total_price and total_price > 0 and irm is not None:
            j = safe_float(g(10)) or 0   # 体积
            k = safe_float(g(11)) or 0   # 总毛重
            freight = max(round(j, 2), k / 1000) * 15
            premium = round(total_price * 1.1 * 0.000145, 8)
            cost = round((irm - (tax_refund or 0) + freight + premium) / total_price, 8)

        d = OrderedDict()
        d["提单号"] = g(1) or ""               # col1 提运单号
        d["船次"] = "989船"                     # col2
        d["序号"] = int(seq)
        d["外贸合同号"] = g(4) or ""            # col4
        d["内贸合同号"] = g(5) or ""            # col5
        d["品名"] = g(6) or ""                 # col6
        d["英文品名"] = g(7) or ""              # col7
        d["打包后件数"] = safe_int(g(8), default=None)    # col8
        d["单位种类"] = g(14) or ""             # col14 报关后单位
        d["体积"] = safe_float(g(10), default=None)       # col10
        d["总毛重"] = safe_float(g(11), default=None)     # col11
        d["总净重"] = safe_float(g(12), default=None)     # col12
        d["总数量"] = safe_float(g(13), default=None)     # col13
        d["单位"] = g(15) or ""              # col15 报关单位
        d["法定单位"] = g(21) or ""            # col21 法定单位
        d["单价（美金）"] = None               # 公式 P = Q/M
        d["总 价（美金）"] = total_price       # col18
        d["换汇成本"] = cost
        d["运费"] = None                      # 公式
        d["保费"] = None                      # 公式
        d["报关单号"] = g(1) or ""            # col1
        d["出口/海关编码"] = g(22) or ""       # col22
        d["内贸金额"] = irm                   # col23
        d["退税率"] = trr                    # col24
        d["退税金额"] = tax_refund
        d["申报要素"] = g(26) or ""           # col26
        d["供应商"] = g(27) or ""             # col27
        rows.append(d)

    result = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    return result


# =========================================================
# 导出 Excel
# =========================================================
def to_excel_with_layout(result_df, path, template_path=None):
    n = len(result_df)
    data_first = 3
    data_last = 2 + n
    total_row = data_last + 1
    last_col = len(result_df.columns)
    col_letter = {name: LETTERS[i] for i, name in enumerate(result_df.columns)}

    def L(name):
        return col_letter[name]

    if template_path:
        _write_from_template(result_df, path, template_path, n)
        return path

    # ---- 手动构建（兜底）----
    wb = Workbook()
    ws = wb.active
    ws.title = "汇总表"

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(*(Side(style="thin"),) * 4)

    # 标题行(第2行)
    for i, name in enumerate(result_df.columns):
        cell = ws.cell(row=2, column=i + 1, value=name)
        cell.font = Font(bold=True, name="Microsoft YaHei", size=10)
        cell.alignment = center
        cell.border = border

    # 数据行
    for row_off, (_, row) in enumerate(result_df.iterrows()):
        er = data_first + row_off
        for name in result_df.columns:
            cell = ws.cell(row=er, column=result_df.columns.get_loc(name) + 1)
            val = row[name]
            if val is None or (isinstance(val, str) and val == ""):
                continue
            cell.value = val
            cell.alignment = center
            cell.border = border
            cell.font = Font(name="Microsoft YaHei", size=10)

    _write_formulas(ws, data_first, data_last, col_letter)

    # 合并（仅4列，船次不合并）
    for name in ["打包后件数", "单位种类", "体积", "总毛重"]:
        _merge_blanks(ws, result_df.columns.get_loc(name) + 1, data_first, data_last)

    # 合计行
    for name in ["打包后件数", "体积", "总毛重", "总净重", "总数量",
                 "总 价（美金）", "内贸金额", "退税金额"]:
        cell = ws.cell(row=total_row, column=result_df.columns.get_loc(name) + 1)
        cell.value = f"=SUM({L(name)}{data_first}:{L(name)}{data_last})"
        cell.font = Font(bold=True, name="Microsoft YaHei", size=10)
        cell.alignment = center
        cell.border = border

    # 列宽
    widths = {"提单号": 13, "船次": 8, "序号": 7, "外贸合同号": 19, "内贸合同号": 26,
              "品名": 20, "英文品名": 28, "打包后件数": 11, "单位种类": 9, "体积": 10,
              "总毛重": 12, "总净重": 11, "总数量": 10, "单位": 7, "法定单位": 8,
              "单价（美金）": 12, "总 价（美金）": 14, "换汇成本": 10, "运费": 12, "保费": 9,
              "报关单号": 16, "出口/海关编码": 16, "内贸金额": 16, "退税率": 9, "退税金额": 14,
              "申报要素": 42, "供应商": 24}
    for name, w in widths.items():
        if name in col_letter:
            ws.column_dimensions[L(name)].width = w

    ws.row_dimensions[2].height = 32
    for er in range(data_first, data_last + 1):
        ws.row_dimensions[er].height = 42
    ws.row_dimensions[total_row].height = 20
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{LETTERS[last_col-1]}{total_row}"

    wb.save(path)
    return path


def _write_formulas(ws, data_first, data_last, col_letter):
    L = col_letter
    for er in range(data_first, data_last + 1):
        ws[f"{L['单价（美金）']}{er}"] = f"={L['总 价（美金）']}{er}/{L['总数量']}{er}"
        ws[f"{L['运费']}{er}"] = (
            f"=MAX(ROUND(SUM({L['体积']}{er}),2),"
            f"SUM({L['总毛重']}{er})/1000)*15"
        )
        ws[f"{L['保费']}{er}"] = f"={L['总 价（美金）']}{er}*1.1*0.000145"
        ws[f"{L['退税金额']}{er}"] = f"={L['内贸金额']}{er}/1.13*{L['退税率']}{er}"


def _write_from_template(result_df, path, template_path, n):
    """以源文件为模板，复刻列宽/字体/边框/冻结/筛选，仅重写数据区"""
    wb = load_workbook(template_path)
    ws = wb.active
    ws.title = "汇总表"

    data_first = 3
    data_last = 2 + n
    total_row = data_last + 1
    last_col = len(result_df.columns)
    col_letter = {name: LETTERS[i] for i, name in enumerate(result_df.columns)}

    # 清除数据区旧合并（保留标题行第2行的合并结构）
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= data_first:
            ws.unmerge_cells(str(rng))

    # 清除旧数据值(保留第1行空、第2行标题)
    for er in range(data_first, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=er, column=c)
            if not type(cell).__name__ == "MergedCell":
                cell.value = None

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(*(Side(style="thin"),) * 4)

    # 重写标题(第2行)
    for i, name in enumerate(result_df.columns):
        cell = ws.cell(row=2, column=i + 1, value=name)
        cell.alignment = center
        cell.border = border

    # 数据行
    for row_off, (_, row) in enumerate(result_df.iterrows()):
        er = data_first + row_off
        for name in result_df.columns:
            cell = ws.cell(row=er, column=result_df.columns.get_loc(name) + 1)
            if type(cell).__name__ == "MergedCell":
                continue
            val = row[name]
            if val is None or (isinstance(val, str) and val == ""):
                continue
            cell.value = val
            cell.alignment = center
            cell.border = border

    _write_formulas(ws, data_first, data_last, col_letter)

    # 合并（仅4列，船次不合并）
    for name in ["打包后件数", "单位种类", "体积", "总毛重"]:
        _merge_blanks(ws, result_df.columns.get_loc(name) + 1, data_first, data_last)

    # 合计行
    for name in ["打包后件数", "体积", "总毛重", "总净重", "总数量",
                 "总 价（美金）", "内贸金额", "退税金额"]:
        cell = ws.cell(row=total_row, column=result_df.columns.get_loc(name) + 1)
        cell.value = f"=SUM({col_letter[name]}{data_first}:{col_letter[name]}{data_last})"
        cell.font = Font(bold=True, name="Microsoft YaHei", size=10)
        cell.alignment = center
        cell.border = border

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{LETTERS[last_col-1]}{total_row}"

    wb.save(path)


def _merge_blanks(ws, col_idx, start, end):
    """空格向上合并：连续空格并入其前一个非空单元格"""
    nonempty = [r for r in range(start, end + 1)
                if not _is_empty(ws.cell(row=r, column=col_idx).value)]
    boundaries = [start - 1] + nonempty + [end + 1]
    for i in range(1, len(boundaries) - 1):
        s = boundaries[i]
        e = boundaries[i + 1] - 1
        if e > s:
            ws.merge_cells(start_row=s, start_column=col_idx,
                           end_row=e, end_column=col_idx)


def transform(df):
    """兼容旧接口"""
    return convert_table(df)

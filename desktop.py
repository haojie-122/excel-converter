"""
Excel 表格转化工具 - 桌面版（tkinter）
双击运行：选择数据表 Excel -> 自动转换 -> 保存为汇总表
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from converter import convert_table, to_excel_with_layout
import os
import traceback


def do_convert(input_path, output_path):
    """执行转换，返回行数"""
    xls = pd.ExcelFile(input_path)
    sheet = xls.sheet_names[0]
    # 第1行空，第2行是标题
    df = pd.read_excel(xls, sheet_name=sheet, header=1)

    result = convert_table(df)
    # 传入 template_path = 源文件，直接复刻源表的全部格式
    to_excel_with_layout(result, output_path, template_path=input_path)
    return len(result)


def choose_file():
    path = filedialog.askopenfilename(
        title="选择表一 Excel 文件",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not path:
        return

    try:
        base_name = os.path.splitext(os.path.basename(path))[0]
        default_name = f"{base_name}_汇总表.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="保存汇总表"
        )
        if not save_path:
            return

        count = do_convert(path, save_path)

        messagebox.showinfo(
            "转换完成",
            f"共 {count} 行数据（27列）\n按序号升序排列\n\n"
            f"✅ 已复刻源表格式（列宽/字体/合并/冻结）\n"
            f"✅ 运费/保费/单价/退税金额 已写入公式，打开自动计算\n"
            f"✅ 末尾已加合计行（SUM）\n\n"
            f"已保存到：\n{save_path}"
        )
    except Exception as e:
        messagebox.showerror("转换失败", str(e) + "\n\n" + traceback.format_exc())


# ---------- UI ----------
root = tk.Tk()
root.title("Excel 表格转化工具")
root.geometry("480x280")
root.resizable(False, False)

# 居中
root.update_idletasks()
x = (root.winfo_screenwidth() // 2) - (480 // 2)
y = (root.winfo_screenheight() // 2) - (280 // 2)
root.geometry(f"480x280+{x}+{y}")

tk.Label(
    root,
    text="Excel 汇总表生成工具",
    font=("Microsoft YaHei", 16, "bold")
).pack(pady=18)

tk.Label(
    root,
    text="数据表 -> 汇总表（27列，格式完全复刻源表）",
    font=("Microsoft YaHei", 9),
    fg="gray"
).pack()

tk.Button(
    root,
    text="选择 Excel 并转换",
    command=choose_file,
    width=22,
    height=2,
    font=("Microsoft YaHei", 11)
).pack(pady=15)

tk.Label(
    root,
    text="输出格式：第1行空，第2行标题，第3行起数据（按序号升序）",
    font=("Microsoft YaHei", 8),
    fg="gray"
).pack(pady=5)

tk.Label(
    root,
    text="27列：提单号/船次/序号/合同号/品名/件数/单位/体积/毛重/净重/数量/单价/总价/运费/保费/退税...",
    font=("Microsoft YaHei", 7),
    fg="darkgray",
    wraplength=440,
    justify="center",
).pack(pady=5)

tk.Label(
    root,
    text="✅ 自动复刻源表格式：列宽/字体/对齐/数字格式/合并单元格/冻结窗格",
    font=("Microsoft YaHei", 8),
    fg="green",
).pack(pady=2)

root.mainloop()

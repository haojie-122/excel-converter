"""
Excel 表格转化工具 - 桌面版（tkinter）
双击运行：选择数据表 Excel -> 自动转换 -> 保存汇总表
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from converter import convert_table, to_excel_with_layout
import os
import traceback


def do_convert(input_path, output_path):
    xls = pd.ExcelFile(input_path)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet, header=1)
    result = convert_table(df)
    to_excel_with_layout(result, output_path)
    return len(result)


def choose_file():
    path = filedialog.askopenfilename(
        title="选择数据表 Excel 文件",
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
            f"共 {count} 行数据（27列）\n\n"
            f"✅ 空单元格已合并\n"
            f"✅ 全部居中对齐\n"
            f"✅ 运费公式已写入\n\n"
            f"已保存到：\n{save_path}"
        )
    except Exception as e:
        messageback = traceback.format_exc()
        messagebox.showerror("转换失败", str(e) + "\n\n" + traceback.format_exc())


# ---------- UI ----------
root = tk.Tk()
root.title("Excel 表格转化工具")
root.geometry("500x280")
root.resizable(False, False)

root.update_idletasks()
x = (root.winfo_screenwidth() // 2) - (500 // 2)
y = (root.winfo_screenheight() // 2) - (280 // 2)
root.geometry(f"500x280+{x}+{y}")

tk.Label(
    root,
    text="Excel 汇总表生成工具",
    font=("Microsoft YaHei", 16, "bold")
).pack(pady=18)

tk.Label(
    root,
    text="数据表 -> 汇总表（27列，合并+居中+运费公式）",
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
    text="输出格式：第1行空 | 第2行标题 | 第3行起数据（按序号升序）",
    font=("Microsoft YaHei", 8),
    fg="gray"
).pack(pady=5)

tk.Label(
    root,
    text="运费公式：=MAX(ROUND(SUM(Jn),2),SUM(Kn)/1000)*15",
    font=("Microsoft YaHei", 8),
    fg="blue"
).pack(pady=5)

root.mainloop()

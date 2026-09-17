import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from converter import convert_table, to_excel_with_layout
import os


def choose_file():
    path = filedialog.askopenfilename(
        title="选择 Excel 文件",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not path:
        return

    try:
        xls = pd.ExcelFile(path)
        sheet = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet)

        result = convert_table(df)

        base_name = os.path.splitext(os.path.basename(path))[0]
        default_name = f"{base_name}_汇总表.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="保存汇总表"
        )
        if save_path:
            to_excel_with_layout(result, save_path)
            messagebox.showinfo(
                "完成",
                f"转换成功！\n\n共 {len(result)} 行数据\n按序号升序排列\n第1行空行，第2行标题\n\n已保存到：\n{save_path}"
            )
    except Exception as e:
        messagebox.showerror("出错", str(e))


root = tk.Tk()
root.title("Excel 表格转化工具")
root.geometry("450x240")
root.resizable(False, False)

# 居中
root.update_idletasks()
x = (root.winfo_screenwidth() // 2) - (450 // 2)
y = (root.winfo_screenheight() // 2) - (240 // 2)
root.geometry(f"450x240+{x}+{y}")

tk.Label(
    root,
    text="Excel 序号聚合转换工具",
    font=("Microsoft YaHei", 16, "bold")
).pack(pady=25)

tk.Button(
    root,
    text="选择 Excel 并转换",
    command=choose_file,
    width=22,
    height=2,
    font=("Microsoft YaHei", 11)
).pack(pady=10)

tk.Label(
    root,
    text="输出格式：第1行空，第2行标题，第3行起数据（按序号升序）",
    font=("Microsoft YaHei", 9),
    fg="gray"
).pack(pady=15)

root.mainloop()

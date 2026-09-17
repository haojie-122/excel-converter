"""
Excel 汇总表生成工具 - 桌面 GUI 版
使用 tkinter 构建，无需浏览器
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

import pandas as pd
from converter import convert_table, to_excel_with_layout


class ExcelConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 汇总表生成工具")
        self.root.geometry("680x420")
        self.root.resizable(True, True)

        # 变量
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status = tk.StringVar(value="就绪")

        self._build_ui()

    def _build_ui(self):
        # 主框架
        main = ttk.Frame(self.root, padding=20)
        main.pack(fill="both", expand=True)

        # 标题
        ttk.Label(main, text="Excel 汇总表生成工具", font=("Microsoft YaHei", 16, "bold")).pack(pady=(0, 20))

        # 输入文件
        frm_input = ttk.Frame(main)
        frm_input.pack(fill="x", pady=8)
        ttk.Label(frm_input, text="源数据表：", width=12).pack(side="left")
        ttk.Entry(frm_input, textvariable=self.input_path, width=50).pack(side="left", padx=(0, 8))
        ttk.Button(frm_input, text="浏览...", command=self._browse_input).pack(side="left")

        # 输出文件
        frm_output = ttk.Frame(main)
        frm_output.pack(fill="x", pady=8)
        ttk.Label(frm_output, text="输出文件：", width=12).pack(side="left")
        ttk.Entry(frm_output, textvariable=self.output_path, width=50).pack(side="left", padx=(0, 8))
        ttk.Button(frm_output, text="浏览...", command=self._browse_output).pack(side="left")

        # 说明
        info = ttk.Label(
            main,
            text=(
                "说明：\n"
                "1. 选择源数据表（Excel 文件，第1行空、第2行标题）\n"
                "2. 选择输出路径（.xlsx）\n"
                "3. 点击「开始转换」\n"
                "4. 输出格式：第1行空、第2行标题、第3行起数据，共27列，自动合并"
            ),
            justify="left",
            font=("Microsoft YaHei", 9),
        )
        info.pack(fill="x", pady=20)

        # 转换按钮
        self.btn_convert = ttk.Button(main, text="开始转换", command=self._start_convert)
        self.btn_convert.pack(pady=10)

        # 进度条
        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 10))

        # 状态栏
        ttk.Label(main, textvariable=self.status, font=("Microsoft YaHei", 9), foreground="gray").pack()

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择源数据表",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)
            # 自动填充输出路径
            if not self.output_path.get():
                base = os.path.splitext(path)[0]
                self.output_path.set(base + "_汇总表.xlsx")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存汇总表",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=os.path.basename(self.output_path.get()) if self.output_path.get() else "汇总表.xlsx"
        )
        if path:
            self.output_path.set(path)

    def _start_convert(self):
        if not self.input_path.get():
            messagebox.showwarning("提示", "请先选择源数据表")
            return
        if not self.output_path.get():
            messagebox.showwarning("提示", "请先选择输出文件")
            return

        self.btn_convert.config(state="disabled")
        self.progress.start(10)
        self.status.set("转换中...")

        thread = threading.Thread(target=self._do_convert, daemon=True)
        thread.start()

    def _do_convert(self):
        try:
            df = pd.read_excel(self.input_path.get(), dtype=str)
            result = convert_table(df)
            to_excel_with_layout(result, self.output_path.get())

            self.root.after(0, lambda: self._on_success())
        except Exception as e:
            self.root.after(0, lambda: self._on_error(str(e)))

    def _on_success(self):
        self.progress.stop()
        self.btn_convert.config(state="normal")
        self.status.set("转换完成 ✓")
        messagebox.showinfo("完成", f"汇总表已生成：\n{self.output_path.get()}")

    def _on_error(self, msg):
        self.progress.stop()
        self.btn_convert.config(state="normal")
        self.status.set("转换失败 ✗")
        messagebox.showerror("错误", f"转换失败：\n{msg}")


def main():
    root = tk.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

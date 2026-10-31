#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx → Excel 汇总工具
功能：选择多个 .docx 文件，提取全文内容，汇总到一个 Excel 表格中。
"""

import os
import sys


# --------------------------------------------------------------------------
# 隐藏 CMD 控制台窗口
# --------------------------------------------------------------------------
def _hide_console() -> None:
    """双击 .py 运行时 Windows 会附带一个黑色控制台窗口。
    这里在程序启动的第一时间把它隐藏掉，只保留 tkinter 界面。
    若用 pythonw.exe 运行（本来就没有控制台），本函数什么都不做。
    """
    if os.name != "nt":
        return
    try:
        import ctypes

        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # 0 = SW_HIDE
    except Exception:
        pass


# 必须放在下面几个重量级导入之前，否则控制台会先闪一下再消失
_hide_console()


def _fatal(message: str) -> None:
    """控制台已隐藏，致命错误改用弹窗提示，免得用户看不到报错。"""
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, message, "启动失败", 0x10)
    except Exception:
        sys.stderr.write(message + "\n")
    sys.exit(1)


from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# 尝试导入依赖
try:
    import docx
except ImportError:
    _fatal("缺少依赖 python-docx\n\n请在命令行执行：\npip install python-docx")

try:
    import pandas as pd
except ImportError:
    _fatal("缺少依赖 pandas\n\n请在命令行执行：\npip install pandas")

try:
    from openpyxl import load_workbook
except ImportError:
    _fatal("缺少依赖 openpyxl\n\n请在命令行执行：\npip install openpyxl")


def extract_text_from_docx(docx_path: Path) -> str:
    """提取一个 docx 文件中的所有段落文本，用换行符连接"""
    try:
        doc = docx.Document(str(docx_path))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]  # 忽略空段
        return "\n".join(paragraphs)
    except Exception as e:
        return f"[读取失败: {e}]"


def summarize_to_excel(docx_files, output_excel_path, progress_callback=None):
    """
    汇总多个 docx 文件到 Excel
    docx_files: list of Path
    output_excel_path: Path
    progress_callback: 可选回调，接收 (当前索引, 总数)
    """
    data = []
    total = len(docx_files)
    for idx, file in enumerate(docx_files, 1):
        text = extract_text_from_docx(file)
        data.append({
            "文件名": file.name,
            "文档内容": text
        })
        if progress_callback:
            progress_callback(idx, total)

    # 写入 Excel
    df = pd.DataFrame(data)
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="汇总", index=False)
        # 调整列宽
        workbook = writer.book
        worksheet = writer.sheets["汇总"]
        worksheet.column_dimensions["A"].width = 40   # 文件名
        worksheet.column_dimensions["B"].width = 100  # 内容

    return len(data)


class Docx2ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("docx → Excel 汇总工具")
        # 取消固定尺寸，改为自动适应内容，并设置最小尺寸
        self.root.minsize(500, 300)

        # 文件列表存储
        self.docx_files = []

        # 主框架
        main = tk.Frame(root, padx=20, pady=20)
        main.pack(fill="both", expand=True)

        tk.Label(main, text="docx 文件汇总至 Excel", font=("微软雅黑", 18, "bold")).pack(pady=(0, 10))

        # 选择文件按钮
        btn_frame = tk.Frame(main)
        btn_frame.pack(fill="x", pady=5)
        tk.Button(btn_frame, text="选择 Word 文档 (多选)", command=self.select_files, bg="#3498db", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_files, bg="#e67e22", fg="white").pack(side="left", padx=5)

        # 文件列表显示
        list_frame = tk.Frame(main)
        list_frame.pack(fill="both", expand=True, pady=10)
        tk.Label(list_frame, text="已选文件：", anchor="w").pack(fill="x")
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        self.file_listbox.pack(fill="both", expand=True, side="left")
        scroll = tk.Scrollbar(list_frame, command=self.file_listbox.yview)
        scroll.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=scroll.set)

        # 状态条 / 进度
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(main, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=5)

        self.status_label = tk.Label(main, text="就绪", anchor="w", fg="#555")
        self.status_label.pack(fill="x")

        # 执行按钮
        tk.Button(main, text="开始汇总到 Excel", command=self.run_summary, bg="#27ae60", fg="white", font=("微软雅黑", 11, "bold")).pack(pady=15)

        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # -------- 关键修改：强制刷新布局，然后自动调整窗口大小并居中 --------
        self.root.update_idletasks()
        # 获取当前实际尺寸，确保不小于最小尺寸
        w = max(self.root.winfo_width(), 500)
        h = max(self.root.winfo_height(), 300)
        # 计算居中位置
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = (screen_w // 2) - (w // 2)
        y = (screen_h // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择 docx 文件",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if not files:
            return
        for f in files:
            p = Path(f)
            if p not in self.docx_files:
                self.docx_files.append(p)
                self.file_listbox.insert(tk.END, p.name)
        self.status_label.config(text=f"已选 {len(self.docx_files)} 个文件")

    def clear_files(self):
        self.docx_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.status_label.config(text="列表已清空")
        self.progress_var.set(0)

    def run_summary(self):
        if not self.docx_files:
            messagebox.showwarning("提示", "请先选择至少一个 docx 文件")
            return

        # 选择保存路径
        default_name = f"docx汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        excel_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not excel_path:
            return

        # 禁用按钮，防止重复点击
        for child in self.root.winfo_children():
            if isinstance(child, tk.Button) and child["text"].startswith("开始"):
                child.config(state="disabled")

        self.status_label.config(text="正在处理，请稍候...")
        self.progress_var.set(0)

        # 启动后台线程
        def worker():
            try:
                total = len(self.docx_files)

                def progress_cb(idx, total):
                    # 更新进度
                    pct = (idx / total) * 100
                    self.root.after(0, lambda: self.progress_var.set(pct))
                    self.root.after(0, lambda: self.status_label.config(text=f"处理中 {idx}/{total} 文件"))

                count = summarize_to_excel(
                    self.docx_files,
                    Path(excel_path),
                    progress_callback=progress_cb
                )
                self.root.after(0, lambda: self.on_finished(True, excel_path, count))
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: self.on_finished(False, str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def on_finished(self, success, info, count=0):
        # 恢复按钮
        for child in self.root.winfo_children():
            if isinstance(child, tk.Button) and child["text"].startswith("开始"):
                child.config(state="normal")
        if success:
            self.status_label.config(text=f"汇总完成！共 {count} 个文件")
            self.progress_var.set(100)
            messagebox.showinfo("完成", f"汇总成功！\n文件保存至：{info}")
        else:
            self.status_label.config(text="汇总失败")
            self.progress_var.set(0)
            messagebox.showerror("错误", f"处理过程中出错：\n{info}")

    def on_close(self):
        self.root.destroy()


if __name__ == "__main__":
    # -------- 添加 Windows DPI 感知，解决高缩放比下界面显示问题 --------
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    app = Docx2ExcelApp(root)
    root.mainloop()

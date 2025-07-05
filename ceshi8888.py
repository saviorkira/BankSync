import os
import sys
import json
import threading
import ttkbootstrap as ttk  # 确保导入 ttkbootstrap
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog, ttk
from datetime import datetime
from tkcalendar import DateEntry
from tkinter.scrolledtext import ScrolledText
from main import main as run_main, force_check_expiration_local

class BankExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("宁波银行数据导出工具")
        self.root.geometry("800x600")
        self.project_root = os.path.abspath(os.path.dirname(__file__))

        # 检查程序是否过期
        force_check_expiration_local(self.project_root)

        # 主框架
        self.main_frame = ttk.Frame(root, padding=10)
        self.main_frame.pack(fill=BOTH, expand=True)

        # Excel 文件选择
        self.excel_label = ttk.Label(self.main_frame, text="Excel 文件：")
        self.excel_label.grid(row=0, column=0, sticky=W, pady=5)
        self.excel_path_var = ttk.StringVar()
        self.excel_entry = ttk.Entry(self.main_frame, textvariable=self.excel_path_var, width=50)
        self.excel_entry.grid(row=0, column=1, sticky=EW, pady=5)
        self.excel_button = ttk.Button(self.main_frame, text="浏览", command=self.browse_excel)
        self.excel_button.grid(row=0, column=2, padx=5)

        # 下载路径选择
        self.download_label = ttk.Label(self.main_frame, text="下载路径：")
        self.download_label.grid(row=1, column=0, sticky=W, pady=5)
        self.download_path_var = ttk.StringVar()
        self.download_entry = ttk.Entry(self.main_frame, textvariable=self.download_path_var, width=50)
        self.download_entry.grid(row=1, column=1, sticky=EW, pady=5)
        self.download_button = ttk.Button(self.main_frame, text="浏览", command=self.browse_download)
        self.download_button.grid(row=1, column=2, padx=5)

        # 日期选择
        self.date_frame = ttk.LabelFrame(self.main_frame, text="日期范围", padding=10)
        self.date_frame.grid(row=2, column=0, columnspan=3, sticky=EW, pady=10)

        self.start_date_label = ttk.Label(self.date_frame, text="开始日期：")
        self.start_date_label.grid(row=0, column=0, sticky=W, padx=5)
        self.start_date_entry = DateEntry(self.date_frame, date_pattern="yyyy-mm-dd")
        self.start_date_entry.grid(row=0, column=1, padx=5)

        self.end_date_label = ttk.Label(self.date_frame, text="结束日期：")
        self.end_date_label.grid(row=0, column=2, sticky=W, padx=5)
        self.end_date_entry = DateEntry(self.date_frame, date_pattern="yyyy-mm-dd")
        self.end_date_entry.grid(row=0, column=3, padx=5)

        # 日志显示区域
        self.log_frame = ttk.LabelFrame(self.main_frame, text="运行日志", padding=10)
        self.log_frame.grid(row=3, column=0, columnspan=3, sticky=NSEW, pady=10)
        self.log_text = ScrolledText(self.log_frame, height=15, width=80, wrap=WORD)
        self.log_text.grid(row=0, column=0, sticky=NSEW)
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)

        # 按钮区域
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=4, column=0, columnspan=3, pady=10)
        self.start_button = ttk.Button(self.button_frame, text="开始导出", style="primary.TButton", command=self.start_export)
        self.start_button.grid(row=0, column=0, padx=5)
        self.clear_button = ttk.Button(self.button_frame, text="清空日志", style="secondary.TButton", command=self.clear_log)
        self.clear_button.grid(row=0, column=1, padx=5)

        # 配置主框架权重
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(3, weight=1)

        # 日志回调函数
        self.log_callback = self.append_log

    def browse_excel(self):
        """选择 Excel 文件"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if file_path:
            self.excel_path_var.set(file_path)

    def browse_download(self):
        """选择下载路径"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.download_path_var.set(folder_path)

    def append_log(self, message):
        """将日志追加到文本框"""
        self.log_text.insert(END, f"{datetime.now()}: {message}\n")
        self.log_text.see(END)
        self.root.update()

    def clear_log(self):
        """清空日志文本框"""
        self.log_text.delete(1.0, END)

    def start_export(self):
        """开始导出流程"""
        download_path = self.download_path_var.get()
        excel_path = self.excel_path_var.get()
        start_date = self.start_date_entry.get()
        end_date = self.end_date_entry.get()

        if not download_path or not excel_path:
            Messagebox.show_error("请填写所有字段，包括 Excel 文件路径和下载路径。", title="输入错误")
            return

        if not os.path.exists(excel_path):
            Messagebox.show_error("Excel 文件不存在，请重新选择。", title="文件错误")
            return

        if not os.path.exists(download_path):
            Messagebox.show_error("下载路径不存在，请重新选择。", title="路径错误")
            return

        # 禁用按钮防止重复点击
        self.start_button.config(state=DISABLED)
        self.append_log("开始导出流程...")

        # 在线程中运行 main 函数
        def run_in_thread():
            try:
                sys.argv = [sys.argv[0], download_path, excel_path, start_date, end_date]
                run_main()
                self.root.after(0, lambda: Messagebox.show_info("导出完成！", title="成功"))
            except Exception as e:
                error_message = json.loads(e.args[0])["message"] if e.args[0].startswith('{"status":"error"') else str(e)
                self.root.after(0, lambda: Messagebox.show_error(f"导出失败：{error_message}", title="错误"))
            finally:
                self.root.after(0, lambda: self.start_button.config(state=NORMAL))

        threading.Thread(target=run_in_thread, daemon=True).start()

if __name__ == "__main__":
    root = ttk.Window(themename="darkly")
    app = BankExportApp(root)
    root.mainloop()
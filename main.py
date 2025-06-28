import os
import sys
import threading
import pandas as pd
import webview
from ningbo_bank import run_ningbo_bank
from utils import force_check_expiration_local, log


class Api:
    def __init__(self):
        self.excel_data = []
        self.base_path = r"D:\Desktop"
        self.is_running = False
        self.window = None

    def set_window(self, window):
        self.window = window

    def log(self, message):
        if self.window:
            self.window.evaluate_js(f"window.updateLog('{message.replace("'", "\\'")}')")
        log(message, self.base_path)

    def import_excel(self):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
            if not file_path:
                self.log("错误: 未选择 Excel 文件")
                return {"success": False, "message": "未选择 Excel 文件"}
            df = pd.read_excel(file_path, header=0)
            if df.empty or len(df.columns) < 2:
                self.log("错误: Excel 文件格式错误，至少需要两列（项目名称和银行账号）")
                return {"success": False, "message": "Excel 文件格式错误"}
            self.excel_data = [(str(project).strip(), str(account).strip()) for project, account in
                               zip(df.iloc[:, 0], df.iloc[:, 1]) if str(project).strip() and str(account).strip()]
            self.log(f"成功导入 Excel 文件：{file_path}")
            return {"success": True, "data": self.excel_data}
        except Exception as e:
            self.log(f"导入 Excel 失败：{str(e)}")
            return {"success": False, "message": str(e)}

    def select_base_path(self):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            path = filedialog.askdirectory()
            if not path:
                self.log("错误: 未选择下载路径")
                return {"success": False, "message": "未选择下载路径"}
            self.base_path = path
            self.log(f"下载路径设置为：{path}")
            return {"success": True, "path": path}
        except Exception as e:
            self.log(f"选择下载路径失败：{str(e)}")
            return {"success": False, "message": str(e)}

    def run_export(self, bank, start_date, end_date):
        if self.is_running:
            self.log("提示: 导出进程正在运行，请等待")
            return {"success": False, "message": "导出进程正在运行"}
        if not bank:
            self.log("错误: 请先选择银行")
            return {"success": False, "message": "请先选择银行"}
        if bank != "Ningbo Bank":
            self.log("错误: 仅支持宁波银行")
            return {"success": False, "message": "仅支持宁波银行"}
        if not self.excel_data:
            self.log("提示: 请先导入 Excel 文件")
            return {"success": False, "message": "请先导入 Excel 文件"}
        if not (start_date and end_date and self.base_path):
            self.log("提示: 请填写完整日期和下载路径")
            return {"success": False, "message": "请填写完整日期和下载路径"}
        self.is_running = True
        self.log("开始运行宁波银行导出...")

        def worker():
            try:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as playwright:
                    run_ningbo_bank(playwright, self.base_path, self.excel_data, start_date, end_date,
                                    log_callback=self.log)
                self.log("导出完成。")
            except Exception as e:
                self.log(f"执行出错: {str(e)}")
            finally:
                self.is_running = False
                if self.window:
                    self.window.evaluate_js("window.updateRunButton(false, '运行')")

        threading.Thread(target=worker, daemon=True).start()
        return {"success": True, "message": "导出已开始"}


if __name__ == "__main__":
    force_check_expiration_local("2026-06-01")
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}")
    if not os.path.exists(os.environ["PLAYWRIGHT_BROWSERS_PATH"]):
        log(f"错误: Playwright 浏览器路径不存在: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}")
        sys.exit(1)

    api = Api()
    window = webview.create_window(
        "银行流水回单导出",
        "./frontend/dist/index.html",
        js_api=api,
        width=800,
        height=600,
        resizable=True
    )
    api.set_window(window)
    try:
        webview.start()
    except Exception as e:
        log(f"应用启动失败: {str(e)}")
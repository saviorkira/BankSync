import os
import sys
import threading
import configparser
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QTableWidget, QPushButton,
    QFileDialog, QMessageBox, QLineEdit, QFormLayout, QTableWidgetItem,
    QTextEdit
)
from PySide6.QtCore import Qt
from playwright.sync_api import sync_playwright, Playwright
from datetime import datetime

def force_check_expiration_local(expire_date_str="2025-06-30"):
    """使用本地系统时间判断是否过期，过期则显示弹窗退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        QMessageBox.critical(None, "错误", f"程序已过期（截止日期为 {expire_date_str}）。")
        sys.exit(0)

def read_bank_config():
    config_path = get_resource_path("../../config.txt")
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')
    if "ningbo_bank" not in config:
        raise KeyError("配置文件中缺少 [ningbo_bank] 配置项")
    bank_conf = config["ningbo_bank"]
    username = bank_conf.get("username")
    password = bank_conf.get("password")
    login_url = bank_conf.get("login_url")
    if not all([username, password, login_url]):
        raise ValueError("ningbo_bank 配置不完整，请检查 config.txt")
    return username, password, login_url

def run_ningbo_bank(playwright: Playwright, base_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    log(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    username, password, login_url = read_bank_config()
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    try:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto(login_url)
        page.get_by_role("textbox", name="用户名").fill(username)
        page.get_by_role("textbox", name="请输入您的密码").fill(password)
        page.wait_for_selector('text=账户管理', timeout=60000)
        page.get_by_role("link", name="账户管理").click()
        page.get_by_role("link", name="账户明细").click()

        for xiangmu, account in projects_accounts:
            log(f"处理项目：{xiangmu}，银行账号：{account}")
            try:
                duizhang_path = os.path.join(base_path, xiangmu, "银行流水")
                huidan_path = os.path.join(base_path, xiangmu, "银行回单")
                os.makedirs(duizhang_path, exist_ok=True)
                os.makedirs(huidan_path, exist_ok=True)

                page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
                page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").fill(account)
                log(f"使用银行账号查询：{account}")
                page.get_by_role("listitem").filter(has_text=xiangmu).locator("span").nth(2).click()
                page.get_by_text("展开").first.click()
                page.get_by_role("textbox", name="开始日期").fill(kaishiriqi)
                page.get_by_role("textbox", name="开始日期").press("Enter")
                page.get_by_role("textbox", name="结束日期").fill(jieshuriqi)
                page.get_by_role("textbox", name="结束日期").press("Enter")
                page.get_by_role("button", name=" 查询").click()
                page.get_by_role("checkbox", name="Toggle Selection of All Rows").check()
                page.get_by_role("button", name="导出 ").click()
                try:
                    page.wait_for_timeout(1000)
                    menus = page.locator("css=[id^='dropdown-menu-']")
                    count = menus.count()
                    found = False
                    for i in range(count):
                        item = menus.nth(i).filter(has_text="凭证导出")
                        if item.is_visible():
                            with page.expect_download() as download_info:
                                item.click()
                            download = download_info.value
                            filename = f"{xiangmu}_银行回单_{kaishiriqi}_{jieshuriqi}.pdf"
                            download.save_as(os.path.join(huidan_path, filename))
                            log(f"银行回单导出完成：{filename}")
                            found = True
                            break
                    if not found:
                        log(f"未找到‘凭证导出’菜单项（项目：{xiangmu}）")
                        page.screenshot(path=f"error_menu_{xiangmu}.png")
                        log(f"已保存错误截图：error_menu_{xiangmu}.png")
                except Exception as e:
                    log(f"导出银行回单失败（项目：{xiangmu}）：{str(e)}")
                    page.screenshot(path=f"error_menu_{xiangmu}.png")
                    log(f"已保存错误截图：error_menu_{xiangmu}.png")
                page.get_by_role("button", name="导出 ").click()
                with page.expect_download() as download_info:
                    page.get_by_text("对账单导出", exact=True).click()
                download = download_info.value
                filename = f"{xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
                download.save_as(os.path.join(duizhang_path, filename))
                log(f"银行流水导出完成：{filename}")

                # # 导航回账户明细页面，为下一个项目做准备
                # page.get_by_role("link", name="账户管理").click()
                # page.get_by_role("link", name="账户明细").click()
            except Exception as e:
                with open(log_path, "a", encoding="utf-8") as f:
                    f.write(f"导出失败：{xiangmu} {kaishiriqi}~{jieshuriqi} 错误：{str(e)}\n")
                log(f"导出失败（项目：{xiangmu}）：{str(e)}")
                continue  # 继续处理下一个项目

        # os.startfile(os.path.join(base_path, xiangmu))
        context.close()
        browser.close()
    except Exception as e:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"初始化失败：{str(e)}\n")
        log(f"初始化失败：{str(e)}")

def get_resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("银行流水回单导出")
        self.resize(600, 400)
        self.data = []  # 存储 Excel 数据

        layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.import_btn = QPushButton("导入 Excel")
        self.import_btn.clicked.connect(self.import_excel)
        form_layout.addWidget(self.import_btn)

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["项目名称", "银行账号"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 300)
        self.table.setColumnWidth(1, 150)
        form_layout.addRow("导入数据:", self.table)

        self.start_date_edit = QLineEdit()
        self.start_date_edit.setPlaceholderText("开始日期 (例如 2025-03-01)")
        self.start_date_edit.setText("2025-03-01")
        form_layout.addRow("开始日期:", self.start_date_edit)

        self.end_date_edit = QLineEdit()
        self.end_date_edit.setPlaceholderText("结束日期 (例如 2025-03-31)")
        self.end_date_edit.setText("2025-03-31")
        form_layout.addRow("结束日期:", self.end_date_edit)

        self.base_path_edit = QLineEdit()
        self.base_path_edit.setText(r"D:\Desktop")
        form_layout.addRow("下载路径:", self.base_path_edit)

        self.select_path_btn = QPushButton("选择下载路径")
        self.select_path_btn.clicked.connect(self.select_base_path)
        form_layout.addWidget(self.select_path_btn)

        layout.addLayout(form_layout)

        self.run_btn = QPushButton("宁波银行")
        self.run_btn.clicked.connect(self.run_ningbo_bank)
        layout.addWidget(self.run_btn)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        layout.addWidget(self.log_edit)

        self.setLayout(layout)

    def import_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)"
        )
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path, header=0)
            if df.empty or len(df.columns) < 2:
                QMessageBox.warning(self, "错误", "Excel 文件格式错误，至少需要两列（项目名称和银行账号）")
                return
            self.data = [(str(project).strip(), str(account).strip()) for project, account in zip(df.iloc[:, 0], df.iloc[:, 1]) if str(project).strip() and str(account).strip()]
            self.table.setRowCount(len(self.data))
            for row, (project, account) in enumerate(self.data):
                self.table.setItem(row, 0, QTableWidgetItem(project))
                self.table.setItem(row, 1, QTableWidgetItem(account))
            self.table.resizeColumnsToContents()
            self.log(f"成功导入 Excel 文件：{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导入 Excel 失败：{str(e)}")

    def select_base_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择下载路径", self.base_path_edit.text())
        if path:
            self.base_path_edit.setText(path)

    def log(self, msg):
        self.log_edit.append(msg)
        print(msg)

    def run_ningbo_bank(self):
        if not self.data:
            QMessageBox.warning(self, "提示", "请先导入 Excel 文件")
            return
        kaishi = self.start_date_edit.text().strip()
        jieshu = self.end_date_edit.text().strip()
        base_path = self.base_path_edit.text().strip()
        if not (kaishi and jieshu and base_path):
            QMessageBox.warning(self, "提示", "请填写完整日期和下载路径")
            return
        def worker():
            self.log("开始运行宁波银行导出...")
            try:
                with sync_playwright() as playwright:
                    run_ningbo_bank(playwright, base_path, self.data, kaishi, jieshu, log_callback=self.log)
                self.log("导出完成。")
            except Exception as e:
                self.log(f"执行出错: {str(e)}")
        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    force_check_expiration_local("2026-06-01")
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_dir, "playwright-browsers")
    print("设置 PLAYWRIGHT_BROWSERS_PATH:", os.environ["PLAYWRIGHT_BROWSERS_PATH"])
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
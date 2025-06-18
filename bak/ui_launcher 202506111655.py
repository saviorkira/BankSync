import sys
import os
import subprocess
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QTextEdit,
    QPushButton, QFileDialog, QMessageBox, QLineEdit, QFormLayout
)


def get_resource_path(relative_path):
    """获取资源的绝对路径，兼容PyInstaller打包后路径"""
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)




def ensure_playwright_browsers_installed():
    """检测并安装 Playwright 浏览器内核"""
    try:
        import playwright
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
        print("Playwright浏览器内核已安装。")
    except Exception:
        print("未检测到Playwright浏览器内核，开始安装...")
        result = subprocess.run([sys.executable, "-m", "playwright", "install"], capture_output=True, text=True)
        if result.returncode == 0:
            print("浏览器内核安装完成。")
        else:
            print("浏览器内核安装失败，请手动运行：")
            print(f"{sys.executable} -m playwright install")
            QMessageBox.critical(None, "错误", "Playwright浏览器内核安装失败，请手动运行:\n"
                                               f"{sys.executable} -m playwright install\n"
                                               f"错误信息：{result.stderr}")
            sys.exit(1)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("银行流水回单导出")
        self.resize(600, 400)

        layout = QVBoxLayout()

        form_layout = QFormLayout()

        self.xiangmu_edit = QTextEdit()
        self.xiangmu_edit.setPlaceholderText("请输入项目名称（多项目请换行）")
        self.xiangmu_edit.setPlainText("重信·北极22026·华睿精选9号集合资金信托计划")
        form_layout.addRow("项目名称:", self.xiangmu_edit)

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
        form_layout.addRow("基础路径:", self.base_path_edit)

        self.select_path_btn = QPushButton("选择基础路径")
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

    def select_base_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择基础路径", self.base_path_edit.text())
        if path:
            self.base_path_edit.setText(path)

    def log(self, msg):
        self.log_edit.append(msg)
        print(msg)

    def run_ningbo_bank(self):
        xiangmu = self.xiangmu_edit.toPlainText().strip()
        if not xiangmu:
            QMessageBox.warning(self, "提示", "请输入项目名称")
            return
        kaishi = self.start_date_edit.text().strip()
        jieshu = self.end_date_edit.text().strip()
        base_path = self.base_path_edit.text().strip()
        if not (kaishi and jieshu and base_path):
            QMessageBox.warning(self, "提示", "请填写完整日期和基础路径")
            return

        # 改为使用兼容打包环境的资源路径获取函数
        script_path = get_resource_path("ningbo_bank.py")
        if not os.path.exists(script_path):
            self.log("未找到ningbo_bank.py文件！")
            QMessageBox.critical(self, "错误", "未找到ningbo_bank.py文件！")
            return

        import threading

        def worker():
            self.log("开始运行宁波银行导出...")

            env = os.environ.copy()
            # 浏览器内核目录也使用兼容函数
            env["PLAYWRIGHT_BROWSERS_PATH"] = get_resource_path("playwright-browsers")

            cmd = [
                sys.executable,
                script_path,
                xiangmu,
                kaishi,
                jieshu,
                base_path
            ]

            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, env=env)
            for line in proc.stdout:
                self.log(line.strip())
            proc.wait()
            if proc.returncode == 0:
                self.log("导出完成。")
            else:
                self.log(f"导出失败，退出码 {proc.returncode}")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    # 启动时先检测浏览器内核
    ensure_playwright_browsers_installed()

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

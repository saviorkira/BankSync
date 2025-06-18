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

        script_path = get_resource_path("ningbo_bank.py")
        if not os.path.exists(script_path):
            self.log("未找到ningbo_bank.py文件！")
            QMessageBox.critical(self, "错误", "未找到ningbo_bank.py文件！")
            return

        import threading

        def worker():
            self.log("开始运行宁波银行导出...")

            env = os.environ.copy()
            env["PLAYWRIGHT_BROWSERS_PATH"] = get_resource_path("../playwright-browsers")

            cmd = [
                sys.executable,
                script_path,
                xiangmu,
                kaishi,
                jieshu,
                base_path
            ]

            self.log(f"运行命令: {' '.join(cmd)}")

            try:
                proc = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",  # ✅ 解决 UnicodeDecodeError 的关键
                    env=env
                )
                for line in proc.stdout:
                    self.log(line.strip())
                proc.wait()
                if proc.returncode == 0:
                    self.log("导出完成。")
                else:
                    self.log(f"导出失败，退出码 {proc.returncode}")
            except Exception as e:
                self.log(f"执行出错: {str(e)}")

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
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

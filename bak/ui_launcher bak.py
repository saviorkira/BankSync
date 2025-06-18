import sys
import os
import subprocess
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPlainTextEdit,
    QLineEdit, QPushButton, QMessageBox
)
from PySide6.QtCore import Qt

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("银行流水导出工具 - PySide6版本")

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("请输入项目名称（多项目请换行）："))
        self.project_textedit = QPlainTextEdit()
        self.project_textedit.setPlainText("重信·北极22026·华睿精选9号集合资金信托计划")
        layout.addWidget(self.project_textedit)

        layout.addWidget(QLabel("开始日期（YYYY-MM-DD）："))
        self.start_date_edit = QLineEdit("2025-03-01")
        layout.addWidget(self.start_date_edit)

        layout.addWidget(QLabel("结束日期（YYYY-MM-DD）："))
        self.end_date_edit = QLineEdit("2025-03-31")
        layout.addWidget(self.end_date_edit)

        layout.addWidget(QLabel("选择基础路径（下载文件保存根目录）："))
        self.base_path_edit = QLineEdit(r"D:\Desktop")
        layout.addWidget(self.base_path_edit)

        self.run_button = QPushButton("运行宁波银行导出")
        layout.addWidget(self.run_button)
        self.run_button.clicked.connect(self.run_ningbo_bank)

        layout.addWidget(QLabel("输出日志："))
        self.log_textedit = QPlainTextEdit()
        self.log_textedit.setReadOnly(True)
        layout.addWidget(self.log_textedit)

        self.resize(600, 600)

    def run_ningbo_bank(self):
        xiangmu = self.project_textedit.toPlainText().strip()
        kaishiriqi = self.start_date_edit.text().strip()
        jieshuriqi = self.end_date_edit.text().strip()
        base_path = self.base_path_edit.text().strip()

        if not all([xiangmu, kaishiriqi, jieshuriqi, base_path]):
            QMessageBox.warning(self, "提示", "请填写所有参数！")
            return

        script_path = os.path.join(os.path.dirname(__file__), "ningbo_bank.py")
        if not os.path.exists(script_path):
            QMessageBox.critical(self, "错误", f"未找到脚本文件：{script_path}")
            return

        python_exe = sys.executable
        projects = [p.strip() for p in xiangmu.splitlines() if p.strip()]
        self.log_textedit.appendPlainText(f"开始执行 {len(projects)} 个项目...\n")

        for project in projects:
            cmd = [
                python_exe,
                script_path,
                project,
                kaishiriqi,
                jieshuriqi,
                base_path
            ]
            self.log_textedit.appendPlainText(f"运行命令：{' '.join(cmd)}")
            self.log_textedit.repaint()

            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            stdout, stderr = proc.communicate()

            if proc.returncode == 0:
                self.log_textedit.appendPlainText(f"项目【{project}】执行成功:\n{stdout}\n")
            else:
                self.log_textedit.appendPlainText(f"项目【{project}】执行失败:\n{stderr}\n")

        self.log_textedit.appendPlainText("所有项目执行完成。")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

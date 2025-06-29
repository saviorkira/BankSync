import os
import sys
import threading
import pandas as pd
from playwright.sync_api import sync_playwright
import flet as ft
from flet import Page, FilePicker, FilePickerResultEvent
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_overwrite_dialog, handle_save_dialog
from ningbo_bank import run_ningbo_bank
from datetime import datetime

os.environ["PYTHONIOENCODING"] = "utf-8"  # 防止路径乱码

def force_check_expiration_local(expire_date_str="2026-06-01"):
    """使用本地系统时间判断是否过期，过期则退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期为 {expire_date_str}）。", project_root)
        sys.exit(0)

def main(page: Page):
    """Flet桌面应用主函数"""
    page.title = "银行流水回单导出"
    page.window_width = 800
    page.window_height = 600
    page.window_resizable = True

    # UI状态变量
    excel_data = []
    download_path = r"D:\Desktop"  # 用户选择的下载路径
    project_root = os.path.abspath(os.path.dirname(__file__))  # 项目根目录
    start_date = ft.TextField(label="开始日期", value="2025-03-01", width=200)
    end_date = ft.TextField(label="结束日期", value="2025-03-31", width=200)
    bank_dropdown = ft.Dropdown(
        label="选择银行",
        options=[ft.dropdown.Option("Ningbo Bank", "宁波银行")],
        value=None,  # 默认空白
        width=200,
    )
    data_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("项目名称")),
            ft.DataColumn(ft.Text("银行账号")),
        ],
        rows=[],
        expand=True,
    )
    log_area = ft.TextField(
        label="日志",
        multiline=True,
        min_lines=5,
        max_lines=10,
        read_only=True,
        expand=True,
    )
    run_button = ft.ElevatedButton("运行", disabled=False)
    is_running = [False]  # 使用列表以在闭包中修改

    def update_log(msg):
        log_area.value += f"{msg}\n"
        log_area.update()

    def on_bank_select(e):
        if bank_dropdown.value:
            update_log(f"已选择银行: {bank_dropdown.options[0].text if bank_dropdown.value == 'Ningbo Bank' else '未知'}")
        else:
            update_log("银行选择已清空")

    bank_dropdown.on_change = on_bank_select

    def import_excel(e: FilePickerResultEvent):
        if e.files and any(f.name.endswith(('.xlsx', '.xls')) for f in e.files):
            file_path = e.files[0].path
            try:
                df = pd.read_excel(file_path, header=0)
                if df.empty or len(df.columns) < 2:
                    update_log("错误: Excel 文件格式错误，至少需要两列（项目名称和银行账号）")
                    return
                nonlocal excel_data
                excel_data.clear()
                excel_data.extend([(str(project).strip(), str(account).strip()) for project, account in zip(df.iloc[:, 0], df.iloc[:, 1]) if str(project).strip() and str(account).strip()])
                data_table.rows = [
                    ft.DataRow(cells=[
                        ft.DataCell(ft.Text(project)),
                        ft.DataCell(ft.Text(account)),
                    ]) for project, account in excel_data
                ]
                update_log(f"成功导入 Excel 文件：{file_path}")
                page.update()
            except Exception as ex:
                update_log(f"导入 Excel 失败：{str(ex)}")
        else:
            update_log("错误: 请选择有效的 Excel 文件")

    def select_download_path(e: FilePickerResultEvent):
        if e.path:
            nonlocal download_path
            download_path = e.path
            base_path_text.value = f"下载路径: {download_path}"
            update_log(f"下载路径设置为：{download_path}")
            page.update()

    def run_export(e):
        nonlocal is_running
        if is_running[0]:
            update_log("提示: 导出进程正在运行，请等待")
            return
        if not bank_dropdown.value:
            update_log("错误: 请先选择银行")
            return
        if bank_dropdown.value != "Ningbo Bank":
            update_log("错误: 仅支持宁波银行")
            return
        if not excel_data:
            update_log("提示: 请先导入 Excel 文件")
            return
        kaishi = start_date.value.strip()
        jieshu = end_date.value.strip()
        if not (kaishi, jieshu, download_path):
            update_log("提示: 请填写完整日期和下载路径")
            return
        is_running[0] = True
        run_button.disabled = True
        run_button.text = "运行中..."
        run_button.update()
        update_log("开始运行宁波银行导出...")
        def worker():
            try:
                with sync_playwright() as playwright:
                    run_ningbo_bank(playwright, project_root, download_path, excel_data, kaishi, jieshu, log_callback=update_log)
                update_log("导出完成。")
            except Exception as ex:
                update_log(f"执行出错: {str(ex)}")
            finally:
                nonlocal is_running
                is_running[0] = False
                run_button.disabled = False
                run_button.text = "运行"
                run_button.update()
        threading.Thread(target=worker, daemon=True).start()

    run_button.on_click = run_export

    # 文件选择器
    file_picker = FilePicker(on_result=import_excel)
    dir_picker = FilePicker(on_result=select_download_path)
    page.overlay.extend([file_picker, dir_picker])

    # UI布局
    base_path_text = ft.Text(f"下载路径: {download_path}")
    page.add(
        ft.Column(
            [
                bank_dropdown,
                ft.Row([
                    ft.ElevatedButton("导入 Excel", on_click=lambda _: file_picker.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls"])),
                    ft.ElevatedButton("选择下载路径", on_click=lambda _: dir_picker.get_directory_path()),
                ]),
                base_path_text,
                ft.Row([start_date, end_date]),
                run_button,
                ft.Text("导入数据:"),
                ft.Container(data_table, expand=True),
                ft.Text("日志:"),
                ft.Container(log_area, expand=True),
            ],
            spacing=10,
            expand=True,
        )
    )

if __name__ == "__main__":
    force_check_expiration_local("2026-06-01")
    project_root = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(project_root, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}", project_root)
    try:
        ft.app(target=main)
    except Exception as e:
        log(f"应用启动失败: {str(e)}", project_root)
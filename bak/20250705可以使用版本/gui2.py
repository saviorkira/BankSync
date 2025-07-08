import os
import sys
import threading
import configparser
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright
import flet as ft
from flet import (
    Page, FilePicker, FilePickerResultEvent, Theme, Container, Column, Row, Text,
    ElevatedButton, Dropdown, DataTable, DataColumn, DataRow, DataCell, TextField, ListView
)
from ningbo_bank import run_ningbo_bank
from utils import log, read_bank_config, get_resource_path

def main(page: Page):
    """Flet 桌面应用主函数，带美化界面"""
    # 设置窗口和主题
    page.title = "BankSync"
    page.window_max_width = 800
    page.window_max_height = 600
    page.window_resizable = True
    page.theme = Theme(
        color_scheme=ft.ColorScheme(
            primary=ft.Colors.BLUE_700,
            primary_container=ft.Colors.BLUE_100,
            secondary=ft.Colors.GREEN_600,
            background=ft.Colors.GREY_50,
        ),
        visual_density=ft.VisualDensity.COMPACT,
        font_family="FZLanTingHei",
    )
    page.padding = 10
    page.bgcolor = ft.Colors.GREY_50

    # 加载自定义字体
    project_root = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath(os.path.dirname(__file__))
    font_path = os.path.join(project_root, "data", "方正兰亭准黑_GBK.ttf")
    page.fonts = {"FZLanTingHei": font_path}
    page.update()

    # UI 状态变量
    excel_data = []
    base_path = r"D:\Desktop"
    is_running = [False]

    # UI 组件
    bank_dropdown = ft.Dropdown(
        label="选择银行",
        options=[ft.dropdown.Option("Ningbo Bank", "宁波银行")],
        value=None,
        width=490,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        fill_color=ft.Colors.WHITE,  # 强制背景为白色
        content_padding=10,
        border_color=ft.Colors.GREY_300,
        color=ft.Colors.BLACK,
        text_style=ft.TextStyle(color=ft.Colors.BLACK, size=14),
        label_style=ft.TextStyle(color=ft.Colors.BLACK, size=14),
        tooltip="请选择要操作的银行",
    )

    start_date = ft.TextField(
        label="开始日期",
        value="2025-06-01",
        width=240,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        text_style=ft.TextStyle(size=14),
        label_style=ft.TextStyle(size=14),
    )

    end_date = ft.TextField(
        label="结束日期",
        value="2025-07-04",
        width=240,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        text_style=ft.TextStyle(size=14),
        label_style=ft.TextStyle(size=14),
    )

    data_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("项目名称", weight=ft.FontWeight.BOLD, size=14)),
            ft.DataColumn(ft.Text("银行账号", weight=ft.FontWeight.BOLD, size=14)),
        ],
        rows=[],
        expand=True,
        border=ft.border.all(1, ft.Colors.GREY_300),
        bgcolor=ft.Colors.WHITE,
        heading_row_color=ft.Colors.BLUE_50,
    )

    log_area = ft.TextField(
        multiline=True,
        min_lines=4,
        max_lines=6,
        read_only=True,
        expand=True,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        text_style=ft.TextStyle(size=14),
    )

    run_button = ft.ElevatedButton(
        text="运行导出",
        icon=ft.Icons.PLAY_CIRCLE,
        disabled=False,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(size=14),  # 使用 text_style
        ),
        tooltip="开始导出银行流水和回单",
        width=490,
    )

    base_path_text = ft.Text(
        f"下载路径: {base_path}",
        size=14,
        color=ft.Colors.GREY_700,
    )

    select_path_button = ft.ElevatedButton(
        text="选择下载路径",
        icon=ft.Icons.FOLDER_OPEN,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(size=14),
        ),
        tooltip="选择保存导出文件的目录",
        width=490,
    )

    import_excel_button = ft.ElevatedButton(
        text="导入项目数据",
        icon=ft.Icons.UPLOAD_FILE,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(size=14),
        ),
        tooltip="选择包含项目名称和银行账号的 Excel 文件",
        width=490,
    )

    # 日志更新函数
    def update_log(msg):
        log_area.value += f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: {msg}\n"
        log_area.update()
        page.scroll_to(key="log_area", duration=500)

    # 银行选择事件
    def on_bank_select(e):
        if bank_dropdown.value:
            update_log(f"已选择银行: {bank_dropdown.options[0].text if bank_dropdown.value == 'Ningbo Bank' else '未知'}")
        else:
            update_log("银行选择已清空")
        page.update()

    bank_dropdown.on_change = on_bank_select

    # 导入 Excel 文件
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
                        ft.DataCell(ft.Text(project, size=14)),
                        ft.DataCell(ft.Text(account, size=14)),
                    ]) for project, account in excel_data
                ]
                update_log(f"成功导入 Excel 文件：{file_path}")
                page.update()
            except Exception as ex:
                update_log(f"导入 Excel 失败：{str(ex)}")
        else:
            update_log("错误: 请选择有效的 Excel 文件 (.xlsx 或 .xls)")

    # 选择下载路径
    def select_base_path(e: FilePickerResultEvent):
        if e.path:
            nonlocal base_path
            base_path = e.path
            base_path_text.value = f"下载路径: {base_path}"
            update_log(f"下载路径设置为：{base_path}")
            page.update()

    # 运行导出任务
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
        if not (kaishi and jieshu and base_path):
            update_log("提示: 请填写完整日期和下载路径")
            return
        try:
            datetime.strptime(kaishi, "%Y-%m-%d")
            datetime.strptime(jieshu, "%Y-%m-%d")
            if kaishi > jieshu:
                update_log("错误: 开始日期不能晚于结束日期")
                return
        except ValueError:
            update_log("错误: 日期格式不正确，应为 YYYY-MM-DD")
            return
        is_running[0] = True
        run_button.disabled = True
        run_button.text = "运行中..."
        run_button.icon = ft.Icons.HOURGLASS_TOP
        run_button.update()
        update_log("开始运行宁波银行导出...")

        def worker():
            try:
                project_root = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath(os.path.dirname(__file__))
                if not os.path.exists(os.path.join(project_root, "playwright-browsers")):
                    update_log("错误: playwright-browsers 文件夹不存在")
                    return
                if not os.path.exists(os.path.join(project_root, "config.txt")):
                    update_log("错误: config.txt 文件不存在")
                    return
                if not os.path.exists(os.path.join(project_root, "seek")):
                    update_log("错误: seek 文件夹不存在")
                    return
                if not os.path.exists(base_path) or not os.access(base_path, os.W_OK):
                    update_log(f"错误: 下载路径不可访问或不可写: {base_path}")
                    return
                with sync_playwright() as playwright:
                    run_ningbo_bank(playwright, project_root, base_path, excel_data, kaishi, jieshu, log_callback=update_log)
                update_log("导出完成。")
            except Exception as ex:
                update_log(f"执行出错: {str(ex)}")
            finally:
                is_running[0] = False
                run_button.disabled = False
                run_button.text = "运行导出"
                run_button.icon = ft.Icons.PLAY_CIRCLE
                run_button.update()

        threading.Thread(target=worker, daemon=True).start()

    run_button.on_click = run_export
    import_excel_button.on_click = lambda _: file_picker.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls"])
    select_path_button.on_click = lambda _: dir_picker.get_directory_path()

    # 文件选择器
    file_picker = FilePicker(on_result=import_excel)
    dir_picker = FilePicker(on_result=select_base_path)
    page.overlay.extend([file_picker, dir_picker])

    # UI 布局
    page.add(
        ft.Container(
            content=ft.Column(
                [
                    bank_dropdown,
                    ft.Row([start_date, end_date], spacing=10),
                    select_path_button,
                    base_path_text,
                    import_excel_button,
                    ft.Container(
                        content=ft.ListView(
                            controls=[data_table],
                            auto_scroll=False,
                        ),
                        padding=5,
                        border_radius=8,
                        bgcolor=ft.Colors.WHITE,
                        shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
                        width=490,
                        height=100,
                    ),
                    run_button,
                    ft.Container(
                        content=log_area,
                        padding=5,
                        border_radius=8,
                        bgcolor=ft.Colors.WHITE,
                        shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
                        width=490,
                        height=100,
                        key="log_area",
                    ),
                ],
                spacing=10,
                scroll=ft.ScrollMode.AUTO,
            ),
            padding=10,
            border_radius=10,
            bgcolor=ft.Colors.WHITE,
            shadow=ft.BoxShadow(
                blur_radius=10,
                spread_radius=2,
                color=ft.Colors.GREY_400,
            ),
            margin=5,
        )
    )

if __name__ == "__main__":
    force_check_expiration_local = lambda project_root, expire_date_str="2026-06-01": (
        sys.exit(1) if datetime.now() > datetime.strptime(expire_date_str, "%Y-%m-%d")
        else None
    )
    force_check_expiration_local(os.path.abspath(os.path.dirname(__file__)))
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}", base_path)
    try:
        ft.app(target=main)
    except Exception as e:
        log(f"应用启动失败: {str(e)}", base_path)
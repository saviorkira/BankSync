import os
import sys
import threading
import time
import json
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from playwright.sync_api import sync_playwright, Playwright
import flet as ft
from flet import (
    Page, FilePicker, FilePickerResultEvent, Theme, Container, Column, Row, Text,
    ElevatedButton, Dropdown, DataTable, DataColumn, DataRow, DataCell, TextField, ListView,
    NavigationRail, NavigationRailDestination, Ref, AnimatedSwitcher, Checkbox, IconButton,
    FloatingActionButton, Tabs, Tab, OutlinedButton, Image
)
from login_manager import load_site_icons, login_site
from todo_manager import TodoApp
from statement_processor import process_bank_statements, update_with_match_data
from ningbo_bank import run_ningbo_bank
from hangzhou_bank import run_hangzhou_bank
from pingan_bank import run_pingan_bank
from shanghai_bank import run_shanghai_bank
from zheshang_bank import run_zheshang_bank
from zhongxin_bank import run_zhongxin_bank
from utils import log, read_bank_config, get_resource_path, check_expiration_with_ntp
from AI import update_ai_output, send_ai_message

def main(page: Page):
    """Flet 桌面应用主函数，带固定 NavigationRail 和美化界面"""
    # 设置窗口和主题
    page.title = "BankSync"
    project_root = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath(os.path.dirname(__file__))
    icon_path = get_resource_path("S.ico", project_root, subfolder="data")
    page.window.icon = icon_path

    page.window.title_bar_hidden = True
    page.window.title_bar_buttons_hidden = True
    page.padding = 0
    page.window.maximizable = False
    page.window.left = 100
    page.window.top = 100
    page.window.width = 380
    page.window.height = 520
    page.window_resizable = True
    page.window.min_width = 380
    page.window.min_height = 520
    page.window.max_width = 580
    page.window.max_height = 520

    page.theme = Theme(
        color_scheme=ft.ColorScheme(
            primary=ft.Colors.BLUE_700,
            primary_container=ft.Colors.BLUE_100,
            secondary=ft.Colors.GREEN_600,
            background=ft.Colors.WHITE,
        ),
        visual_density=ft.VisualDensity.COMPACT,
        font_family="sansr",
    )
    page.bgcolor = ft.Colors.WHITE

    # 加载自定义字体
    # 加载自定义字体
    font_path_SourceHanSansRegular = os.path.join(project_root, "data", "font", "SourceHanSansSC-Regular.otf")
    font_path_SourceHanSansMedium = os.path.join(project_root, "data", "font", "SourceHanSansSC-Medium.otf")

    page.fonts = {
        "sansr": font_path_SourceHanSansRegular,
        "sansm": font_path_SourceHanSansMedium,
    }

    page.update()

    # UI 状态变量
    excel_data = []
    download_path = r"D:\Desktop"
    is_running = [False]
    selected_index = ft.Ref()
    selected_index.current = 0
    log_messages = []
    is_maximized = [False]
    is_minimized = [True]
    statement_folder = [r"D:\Desktop"]  # 流水文件夹路径，列表形式以支持闭包修改
    export_path = [r"D:\Desktop"]  # 导出路径，列表形式以支持闭包修改
    excel_file_path = [None]  # 新增：存储 Excel 文件路径

    # 页面尺寸配置
    page_sizes = {
        0: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        1: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        2: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        3: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        4: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
    }

    # 银行处理函数映射，仅用于下载
    BANK_HANDLERS = {
        "ningbo_bank": run_ningbo_bank,
        "hangzhou_bank": run_hangzhou_bank,
        "pingan_bank": run_pingan_bank,
        # "shanghai_bank": run_shanghai_bank,
        # "zheshang_bank": run_zheshang_bank,
        "zhongxin_bank": run_zhongxin_bank,
    }

    # 银行名称映射
    BANK_NAMES = {
        "ningbo_bank": "宁波银行",
        "hangzhou_bank": "杭州银行",
        "pingan_bank": "平安银行",
        "shanghai_bank": "上海银行",
        "zheshang_bank": "浙商银行",
        "zhongxin_bank": "中信银行",
    }

    # 计算上个月的默认日期
    current_date = datetime.now()
    last_month = current_date - relativedelta(months=1)
    start_date_value = last_month.replace(day=1).strftime("%Y-%m-%d")
    last_day_of_last_month = (current_date.replace(day=1) - timedelta(days=1)).strftime("%Y-%m-%d")

    # UI 组件
    bank_dropdown = ft.Dropdown(
        label="选择银行",
        options=[
            ft.dropdown.Option(key=key, text=BANK_NAMES.get(key, key))
            for key in BANK_HANDLERS.keys()
        ],
        value=None,
        width=page.window.width-70,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        fill_color=ft.Colors.WHITE,
        content_padding=10,
        border_color=ft.Colors.GREY_300,
        color=ft.Colors.BLACK,
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
        tooltip="请选择要操作的银行",
    )

    start_date = ft.TextField(
        label="开始日期",
        value=start_date_value,  # 上个月第一天
        width=(page.window.width-80)/2,
        border_radius=8,
        expand=True,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        tooltip="日期范围通常3个月以内",
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
    )

    end_date = ft.TextField(
        label="结束日期（<3个月）",
        value=last_day_of_last_month,  # 上个月最后一天
        width=(page.window.width-80)/2,
        border_radius=8,
        expand=True,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        tooltip="日期范围通常3个月以内",
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
    )

    data_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("产品编号", weight=ft.FontWeight.BOLD, size=14, font_family="sansr")),
            ft.DataColumn(ft.Text("产品名称", weight=ft.FontWeight.BOLD, size=14, font_family="sansr")),
            ft.DataColumn(ft.Text("托管账户", weight=ft.FontWeight.BOLD, size=14, font_family="sansr")),
        ],
        rows=[],
        expand=True,
        border=ft.border.all(1, ft.Colors.GREY_300),
        bgcolor=ft.Colors.WHITE,
        heading_row_color=ft.Colors.BLUE_50,
        heading_row_height=30,
        data_row_min_height=30,
        data_row_max_height=30,
    )

    log_area = ft.TextField(
        multiline=True,
        min_lines=20,
        max_lines=20,
        read_only=True,
        expand=True,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        text_style=ft.TextStyle(font_family="sansr", size=14),
    )

    ai_input = ft.TextField(
        label="输入消息",
        expand=True,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="输入消息后点击发送",
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
        multiline=True,
        min_lines=2,
        max_lines=2,
        on_submit=lambda e: send_ai_message(e, ai_input, ai_output, update_log, selected_index, project_root),
    )

    ai_submit_button = ft.FloatingActionButton(
        icon=ft.Icons.SEND,
        on_click=lambda e: send_ai_message(e, ai_input, ai_output, update_log, selected_index, project_root),
        bgcolor=ft.Colors.BLUE_700,
        tooltip="发送消息",
    )

    ai_output = ft.TextField(
        multiline=True,
        min_lines=16,
        max_lines=16,
        read_only=True,
        expand=True,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        text_style=ft.TextStyle(font_family="sansr", size=14),
    )

    run_bankdownloader_button = ft.ElevatedButton(
        text="开始下载",
        icon=ft.Icons.PLAY_CIRCLE,
        disabled=False,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="sansr", size=14),
        ),
        tooltip="开始下载银行流水、回单、对账单",
        width=page.window.width-70,
    )

    download_path_text = ft.Text(
        f"下载路径: {download_path}",
        size=14,
        color=ft.Colors.GREY_700,
        font_family="sansr",
    )

    select_path_button = ft.ElevatedButton(
        text="选择下载路径",
        icon=ft.Icons.FOLDER_OPEN,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="sansr", size=14),
        ),
        tooltip="选择保存下载文件的目录",
        width=page.window.width-70,
    )

    import_excel_button = ft.ElevatedButton(
        text="导入项目数据",
        icon=ft.Icons.UPLOAD_FILE,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="sansr", size=14),
        ),
        tooltip="选择包含项目信息的Excel 文件，参照模板",
        width=page.window.width-70,
    )

    # 数据页面组件
    statement_folder_text = ft.Text(
        f"流水文件夹: {statement_folder[0]}",
        size=14,
        color=ft.Colors.GREY_700,
        font_family="sansr",
    )

    select_statement_folder_button = ft.ElevatedButton(
        text="选择银行流水文件夹",
        icon=ft.Icons.FOLDER_OPEN,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="sansr", size=14),
        ),
        tooltip="选择包含银行流水的文件夹",
        width=page.window.width-70,
        on_click=lambda _: statement_dir_picker.get_directory_path(),
    )

    statement_dropdown = ft.Dropdown(
        label="选择要筛选的银行流水项目",
        options=[
            ft.dropdown.Option(key="bank_interest", text="银行利息"),
            # 可扩展其他选项
        ],
        value=None,
        width=page.window.width-70,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        fill_color=ft.Colors.WHITE,
        content_padding=10,
        border_color=ft.Colors.GREY_300,
        color=ft.Colors.BLACK,
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
    )

    def on_statement_dropdown_change(e):
        if statement_dropdown.value:
            update_log(f"已选择银行流水项目: {statement_dropdown.value}")
        else:
            update_log("银行流水项目已清空")

    statement_dropdown.on_change = on_statement_dropdown_change

    start_date_statement = ft.TextField(
        label="开始月份",
        value=datetime.now().strftime('%Y-%m'),
        width=page.window.width-70,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM",
        tooltip="输入银行流水处理月份",
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
        on_change=lambda e: update_log(f"已设置开始月份: {e.control.value}"),
    )

    export_path_text = ft.Text(
        f"导出文件夹: {export_path[0]}",
        size=14,
        color=ft.Colors.GREY_700,
        font_family="sansr",
    )

    export_statement_button = ft.ElevatedButton(
        text="选择导出文件夹",
        icon=ft.Icons.FOLDER_OPEN,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="sansr", size=14),
        ),
        tooltip="选择导出路径并导出流水",
        width=page.window.width-70,
        on_click=lambda _: export_dir_picker.get_directory_path(),
    )

    start_export_button = ft.ElevatedButton(
        text="导出筛选",
        icon=ft.Icons.PLAY_CIRCLE,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="sansr", size=14),
        ),
        tooltip="开始筛选的银行流水",
        width=page.window.width - 70,
    )

    select_match_excel_button = ft.ElevatedButton(
        text="导出筛选并生成上传模板",
        icon=ft.Icons.PLAY_CIRCLE,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="sansr", size=14),
        ),
        tooltip="选择包含项目信息的 Excel 文件，参照模板",
        width=page.window.width - 70,
    )

    # 日志更新函数
    def update_log(msg):
        log_messages.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: {msg}\n")
        log_area.value = "".join(log_messages)
        if selected_index.current == 4 and tools_content.selected_index == 1 and hasattr(log_area, 'page') and log_area.page is not None:
            log_area.update()
            page.scroll_to(key="log_area", duration=500)
        log(msg, project_root)

    # 银行选择事件
    def on_bank_select(e):
        nonlocal excel_data  # 移动到函数开头
        if bank_dropdown.value:
            update_log(f"已选择银行: {BANK_NAMES.get(bank_dropdown.value, bank_dropdown.value)}")
            # 如果已导入 Excel 文件，重新加载对应银行的 Sheet 数据
            if excel_file_path[0] and os.path.exists(excel_file_path[0]):
                try:
                    sheet_name = BANK_NAMES.get(bank_dropdown.value, bank_dropdown.value)
                    df = pd.read_excel(excel_file_path[0], sheet_name=sheet_name, header=0)
                    required_columns = ["产品编号", "产品名称", "托管账户"]
                    if df.empty or not all(col in df.columns for col in required_columns):
                        update_log(f"错误: Excel 文件 '{sheet_name}' 工作表格式错误，至少需要列：{required_columns}")
                        excel_data.clear()
                        data_table.rows = []
                        page.update()
                        return
                    excel_data.clear()
                    excel_data.extend([(str(xiangmuid).strip(), str(xiangmu).strip(), str(account).strip()) for
                                       xiangmuid, xiangmu, account in
                                       zip(df["产品编号"], df["产品名称"], df["托管账户"]) if
                                       str(xiangmuid).strip() and str(xiangmu).strip() and str(account).strip()])
                    data_table.rows = [
                        ft.DataRow(cells=[
                            ft.DataCell(ft.Text(xiangmuid, size=14, font_family="sansr")),
                            ft.DataCell(ft.Text(xiangmu, size=14, font_family="sansr")),
                            ft.DataCell(ft.Text(account, size=14, font_family="sansr")),
                        ]) for xiangmuid, xiangmu, account in excel_data
                    ]
                    update_log(f"已更新数据表：{sheet_name}，包含 {len(excel_data)} 条记录")
                    page.update()
                except ValueError as ve:
                    update_log(f"错误: 未找到 '{sheet_name}' 工作表：{str(ve)}")
                    excel_data.clear()
                    data_table.rows = []
                    page.update()
                except Exception as ex:
                    update_log(f"更新数据表失败：{str(ex)}")
                    excel_data.clear()
                    data_table.rows = []
                    page.update()
        else:
            update_log("银行选择已清空")
            excel_data.clear()
            data_table.rows = []
            page.update()

    bank_dropdown.on_change = on_bank_select

    # 导入 Excel 文件
    def import_excel(e: ft.FilePickerResultEvent):
        if not bank_dropdown.value:
            update_log("错误: 请先选择银行")
            return
        if e.files and any(f.name.endswith(('.xlsx', '.xls')) for f in e.files):
            file_path = e.files[0].path
            excel_file_path[0] = file_path  # 保存 Excel 文件路径
            try:
                # 获取 Sheet 名称
                sheet_name = BANK_NAMES.get(bank_dropdown.value, bank_dropdown.value)
                # 读取指定 Sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
                # 检查必需列
                required_columns = ["产品编号", "产品名称", "托管账户"]
                if df.empty or not all(col in df.columns for col in required_columns):
                    update_log(f"错误: Excel 文件 '{sheet_name}' 工作表格式错误，至少需要列：{required_columns}")
                    return
                nonlocal excel_data
                excel_data.clear()
                excel_data.extend([(str(xiangmuid).strip(), str(xiangmu).strip(), str(account).strip()) for
                                   xiangmuid, xiangmu, account in
                                   zip(df["产品编号"], df["产品名称"], df["托管账户"]) if
                                   str(xiangmuid).strip() and str(xiangmu).strip() and str(account).strip()])
                data_table.rows = [
                    ft.DataRow(cells=[
                        ft.DataCell(ft.Text(xiangmuid, size=14, font_family="sansr")),
                        ft.DataCell(ft.Text(xiangmu, size=14, font_family="sansr")),
                        ft.DataCell(ft.Text(account, size=14, font_family="sansr")),
                    ]) for xiangmuid, xiangmu, account in excel_data
                ]
                update_log(f"成功导入 Excel 文件：{file_path}，工作表：{sheet_name}，包含 {len(excel_data)} 条记录")
                page.update()
            except ValueError as ve:
                update_log(f"错误: 未找到 '{sheet_name}' 工作表：{str(ve)}")
                excel_data.clear()
                data_table.rows = []
                page.update()
            except Exception as ex:
                update_log(f"导入 Excel 失败：{str(ex)}")
                excel_data.clear()
                data_table.rows = []
                page.update()
        else:
            update_log("错误: 请选择有效的 Excel 文件 (.xlsx 或 .xls)")

    # 选择下载路径##e: FilePickerResultEvent是对的不要改
    def select_download_path(e: FilePickerResultEvent):
        if e.path:
            nonlocal download_path
            download_path = e.path
            download_path_text.value = f"下载路径: {download_path}"
            update_log(f"下载路径设置为：{download_path}")
            page.update()

    # 选择流水文件夹
    def select_statement_folder(e:FilePickerResultEvent):
        if e.path and os.path.exists(e.path):
            statement_folder[0] = e.path
            statement_folder_text.value = f"流水文件夹: {statement_folder[0]}"
            update_log(f"流水文件夹设置为：{statement_folder[0]}")
            page.update()

    # 选择导出路径
    def select_export_path(e: ft.FilePickerResultEvent):
        if e.path and os.path.exists(e.path):
            export_path[0] = e.path
            export_path_text.value = f"导出文件夹: {export_path[0]}"
            update_log(f"导出路径设置为：{export_path[0]}")
            # 这里可以添加导出逻辑，例如处理流水文件夹中的文件
            update_log(f"开始导出流水到：{export_path[0]}")
            page.update()
        else:
            update_log("错误: 请选择有效的文件夹路径")

    # 开始下载任务
    def run_bankdownloader(e):
        nonlocal is_running
        if is_running[0]:
            update_log("提示: 下载进程正在运行，请等待")
            return
        if not bank_dropdown.value:
            update_log("错误: 请先选择银行")
            return
        if bank_dropdown.value not in BANK_HANDLERS:
            update_log("错误: 仅支持已注册的银行")
            return
        if not excel_data:
            update_log("提示: 请先导入 Excel 文件")
            return
        kaishi = start_date.value.strip()
        jieshu = end_date.value.strip()
        if not (kaishi and jieshu and download_path):
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
        run_bankdownloader_button.disabled = True
        run_bankdownloader_button.text = "下载中..."
        run_bankdownloader_button.icon = ft.Icons.HOURGLASS_TOP
        run_bankdownloader_button.update()
        update_log(
            f"开始运行 {BANK_NAMES.get(bank_dropdown.value, bank_dropdown.value)} 下载，包含 {len(excel_data)} 条记录...")

        def worker():
            try:
                if not os.path.exists(os.path.join(project_root, "playwright-browsers")):
                    update_log("错误: playwright-browsers 文件夹不存在")
                    return
                if not os.path.exists(os.path.join(project_root, "config.json")):
                    update_log("错误: config.json 文件不存在")
                    return
                if not os.path.exists(os.path.join(project_root, "data", "cv2")):
                    update_log("错误: data/cv2 文件夹不存在")
                    return
                if not os.path.exists(download_path) or not os.access(download_path, os.W_OK):
                    update_log(f"错误: 下载路径不可访问或不可写: {download_path}")
                    return
                with sync_playwright() as playwright:
                    BANK_HANDLERS[bank_dropdown.value](playwright, project_root, download_path, excel_data, kaishi, jieshu,
                                                       log_callback=update_log)
                update_log("下载完成。")
            except Exception as ex:
                update_log(f"执行出错：{str(ex)}")
            finally:
                is_running[0] = False
                run_bankdownloader_button.disabled = False
                run_bankdownloader_button.text = "开始下载"
                run_bankdownloader_button.icon = ft.Icons.PLAY_CIRCLE
                run_bankdownloader_button.update()

        threading.Thread(target=worker, daemon=True).start()

    # 开始导出任务
    def start_export(e):
        nonlocal is_running
        if is_running[0]:
            update_log("提示: 导出进程正在运行，请等待")
            return
        if not statement_folder[0] or not os.path.exists(statement_folder[0]):
            update_log("错误: 请先选择有效的银行流水文件夹")
            return
        if not export_path[0] or not os.path.exists(export_path[0]):
            update_log("错误: 请先选择有效的导出文件夹")
            return
        if not start_date_statement.value:
            update_log("错误: 请填写开始月份")
            return
        if not statement_dropdown.value:
            update_log("错误: 请先选择银行流水项目")
            return
        is_running[0] = True
        start_export_button.disabled = True
        start_export_button.text = "导出中..."
        start_export_button.icon = ft.Icons.HOURGLASS_TOP
        start_export_button.update()
        update_log(f"开始导出银行流水，月份: {start_date_statement.value}...")

        def worker():
            try:
                if statement_dropdown.value == "bank_interest":
                    success, output_file = process_bank_statements(statement_folder[0], start_date_statement.value, export_path[0], update_log, dropdown_value=statement_dropdown.value or "全部")
                    if success:
                        update_log(f"流水导出完成到：{output_file}")
                    else:
                        update_log("流水导出失败：未找到符合条件的利息数据")
                else:
                    update_log(f"不支持的银行流水项目: {statement_dropdown.value}")
            except Exception as ex:
                update_log(f"导出失败：{str(ex)}")
            finally:
                is_running[0] = False
                start_export_button.disabled = False
                start_export_button.text = "开始导出"
                start_export_button.icon = ft.Icons.PLAY_CIRCLE
                start_export_button.update()

        threading.Thread(target=worker, daemon=True).start()

    # 新函数: select_match_excel (作为 FilePicker 的 on_result)
    def select_match_excel(e: ft.FilePickerResultEvent):
        if not e.files or not any(f.name.endswith(('.xlsx', '.xls')) for f in e.files):
            update_log("错误: 请选择有效的 Excel 文件 (.xlsx 或 .xls)")
            return

        match_file_path = e.files[0].path
        update_log(f"已选择匹配 Excel 文件: {match_file_path}")

        # 先检查必要条件（类似于 start_export）
        if not statement_folder[0] or not os.path.exists(statement_folder[0]):
            update_log("错误: 请先选择有效的银行流水文件夹")
            return
        if not export_path[0] or not os.path.exists(export_path[0]):
            update_log("错误: 请先选择有效的导出文件夹")
            return
        if not start_date_statement.value:
            update_log("错误: 请填写开始月份")
            return
        if not statement_dropdown.value:
            update_log("错误: 请先选择银行流水项目")
            return

        # 标记运行中
        nonlocal is_running
        if is_running[0]:
            update_log("提示: 导出进程正在运行，请等待")
            return
        is_running[0] = True
        select_match_excel_button.disabled = True
        select_match_excel_button.text = "处理中..."
        select_match_excel_button.icon = ft.Icons.HOURGLASS_TOP
        select_match_excel_button.update()
        update_log(f"开始导出并匹配银行流水，月份: {start_date_statement.value}...")

        def worker():
            try:
                if statement_dropdown.value == "bank_interest":
                    success, output_file = process_bank_statements(statement_folder[0], start_date_statement.value, export_path[0], update_log, dropdown_value=statement_dropdown.value or "全部")
                    if success:
                        update_log(f"流水导出完成到：{output_file}")
                        # 调用 statement_processor.py 的新函数进行匹配
                        success = update_with_match_data(output_file, match_file_path, update_log)
                        if success:
                            update_log(f"匹配完成，已更新 {output_file} 添加托管账户和入息账套列")
                        else:
                            update_log("匹配失败：请检查日志")
                    else:
                        update_log("流水导出失败：未找到符合条件的利息数据")
                else:
                    update_log(f"不支持的银行流水项目: {statement_dropdown.value}")
            except Exception as ex:
                update_log(f"导出并匹配失败：{str(ex)}")
            finally:
                is_running[0] = False
                select_match_excel_button.disabled = False
                select_match_excel_button.text = "选择匹配Excel并导出"
                select_match_excel_button.icon = ft.Icons.UPLOAD_FILE
                select_match_excel_button.update()

        threading.Thread(target=worker, daemon=True).start()

    run_bankdownloader_button.on_click = run_bankdownloader
    start_export_button.on_click = start_export
    import_excel_button.on_click = lambda _: file_picker.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls"])
    select_path_button.on_click = lambda _: dir_picker.get_directory_path()
    select_match_excel_button.on_click = lambda _: match_file_picker.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls"])

    # 文件选择器
    file_picker = ft.FilePicker(on_result=import_excel)
    dir_picker = ft.FilePicker(on_result=select_download_path)
    statement_dir_picker = ft.FilePicker(on_result=select_statement_folder)
    export_dir_picker = ft.FilePicker(on_result=select_export_path)
    match_file_picker = ft.FilePicker(on_result=select_match_excel)
    page.overlay.extend([file_picker, dir_picker, statement_dir_picker, export_dir_picker, match_file_picker])
    page.update()  # 确保 FilePicker 控件初始化

    # 切换窗口大小按钮
    toggle_size_button = ft.IconButton(
        icon=ft.Icons.FULLSCREEN_EXIT,
        on_click=lambda _: toggle_window_size(),
        icon_size=15,
        style=ft.ButtonStyle(
            color=ft.Colors.BLUE_600,
            bgcolor=ft.Colors.WHITE,
        ),
        tooltip="切换窗口大小",
    )

    # 关闭按钮
    close_button = ft.IconButton(
        icon=ft.Icons.CLOSE,
        on_click=lambda _: page.window.close(),
        icon_size=15,
        style=ft.ButtonStyle(
            color=ft.Colors.RED_600,
            bgcolor=ft.Colors.WHITE,
        ),
    )

    # 切换窗口大小函数
    def toggle_window_size():
        nonlocal is_minimized
        current_page = selected_index.current
        sizes = page_sizes.get(current_page, page_sizes[0])
        steps = 15
        duration = 0.0133
        start_width = page.window.width
        start_height = page.window.height
        if is_minimized[0]:
            target_width = sizes["max_width"]
            target_height = sizes["max_height"]
            toggle_size_button.icon = ft.Icons.FULLSCREEN
        else:
            target_width = sizes["min_width"]
            target_height = sizes["min_height"]
            toggle_size_button.icon = ft.Icons.FULLSCREEN_EXIT

        for i in range(steps + 1):
            t = i / steps
            page.window.width = start_width + (target_width - start_width) * t
            page.window.height = start_height + (target_height - start_height) * t
            time.sleep(duration)

        # 统一按钮宽度
        button_width = page.window.width - 70
        bank_dropdown.width = button_width
        start_date.width = (page.window.width - 80) / 2
        end_date.width = (page.window.width - 80) / 2
        run_bankdownloader_button.width = button_width
        select_path_button.width = button_width
        import_excel_button.width = button_width
        bank_export_content.controls[1].width = button_width
        bank_export_content.controls[5].width = button_width
        ai_content.controls[0].width = button_width
        ai_content.controls[1].width = button_width
        tools_content.width = button_width
        statement_content.width = button_width
        statement_content.controls[0].width = button_width
        statement_content.controls[1].width = button_width
        statement_content.controls[2].width = button_width
        statement_content.controls[4].width = button_width
        statement_content.controls[5].width = button_width
        statement_content.controls[6].width = button_width
        statement_content.controls[7].width = button_width
        # 更新 tools_content 内部的子控件宽度
        todo_content.controls[0].width = button_width
        settings_content.controls[0].width = button_width
        new_page_content.controls[0].width = button_width
        home_content.controls[0].width = button_width
        main_content.width = button_width
        page.update()

        is_minimized[0] = not is_minimized[0]
        toggle_size_button.update()

    # 页面内容
    home_content = ft.Column(
        [
            # ft.Text("网站图标区域", size=14, font_family="sansr"),
            ft.Row(
                controls=load_site_icons(project_root, update_log, lambda name: login_site(name, project_root, update_log, last_click_time=[0])),
                wrap=True,
                spacing=5,
                run_spacing=5,
                alignment=ft.MainAxisAlignment.SPACE_EVENLY,
                width=page.window.width-70,
                scroll=ft.ScrollMode.AUTO,
            ),
        ],
        spacing=10,
        # padding=0,
        alignment=ft.MainAxisAlignment.START,
    )

    bank_export_content = ft.Column(
        [
            bank_dropdown,
            ft.Row(
                [
                    start_date,
                    end_date,
                ],
                spacing=10,
                width=page.window.width-70,
                expand=True,
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            ),
            select_path_button,
            download_path_text,
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
                width=page.window.width-70,
                height=173,
            ),
            run_bankdownloader_button,
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
    )

    ai_content = ft.Column(
        [
            ft.Container(
                content=ai_output,
                padding=5,
                border_radius=8,
                bgcolor=ft.Colors.WHITE,
                shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
                width=page.window.width-70,
                height=page.window.height-150,
                key="ai_output",
                alignment=ft.alignment.top_left,
            ),
            ft.Row(
                controls=[
                    ai_input,
                    ai_submit_button,
                ],
                spacing=10,
                width=page.window.width-70,
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            ),
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
    )

    todo_content = ft.Column(
        [
            ft.Container(
                content=TodoApp(project_root),
                padding=5,
                border_radius=8,
                bgcolor=ft.Colors.WHITE,
                shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
                width=page.window.width-70,
                height=page.window.height-70,
                alignment=ft.alignment.top_left,
                expand=True,  # 确保容器填充可用空间
            ),
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
    )

    new_page_content = ft.Column(
        [
            ft.Container(
                content=ai_output,
                padding=5,
                border_radius=8,
                bgcolor=ft.Colors.WHITE,
                shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
                width=page.window.width-70,
                height=page.window.height-70,
                alignment=ft.alignment.top_left,
                expand=True,  # 确保容器填充可用空间
            ),
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
    )

    settings_content = ft.Column(
        [
            ft.Container(
                content=log_area,
                padding=5,
                border_radius=8,
                bgcolor=ft.Colors.WHITE,
                shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
                width=page.window.width-70,
                height=page.window.height-70,
                key="log_area",
                alignment=ft.alignment.top_left,
                expand=True,  # 确保容器填充可用空间
            ),
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
    )

    statement_content = ft.Column(
        [
            statement_dropdown,
            start_date_statement,
            select_statement_folder_button,
            statement_folder_text,
            export_statement_button,
            export_path_text,
            start_export_button,
            select_match_excel_button,
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
        # width=page.window.width-70,
    )

    tools_content = ft.Tabs(
        selected_index=0,
        # animation_duration=0,
        tabs=[
            ft.Tab(
                text="待办",
                content=todo_content,
                icon=ft.Icons.TASK_OUTLINED,
            ),
            ft.Tab(
                text="日志",
                content=settings_content,
                icon=ft.Icons.SETTINGS_OUTLINED,
            ),
            ft.Tab(
                text="新页面",
                content=new_page_content,
                icon=ft.Icons.NEW_RELEASES_OUTLINED,
            ),
        ],
        expand=True,
        tab_alignment=ft.TabAlignment.CENTER,
        indicator_color=ft.Colors.BLUE_700,
        label_color=ft.Colors.BLUE_700,
        unselected_label_color=ft.Colors.GREY_700,
        width=page.window.width-70,
        on_change=lambda e: [
            setattr(drag_area_title, "value", ["待办", "日志", "新页面"][e.control.selected_index]),
            drag_area_title.update(),
            page.update(),
        ],
    )

    pages = [
        {"icon": ft.Icons.HOME_OUTLINED, "selected_icon": ft.Icons.HOME, "label": "登录", "content": home_content},
        {"icon": ft.Icons.DOWNLOAD_OUTLINED, "selected_icon": ft.Icons.DOWNLOAD, "label": "下载", "content": bank_export_content},
        {"icon": ft.Icons.DATA_USAGE_OUTLINED, "selected_icon": ft.Icons.DATA_USAGE, "label": "流水", "content": statement_content},
        {"icon": ft.Icons.CHAT_OUTLINED, "selected_icon": ft.Icons.CHAT, "label": "AI", "content": ai_content},
        {"icon": ft.Icons.APPS_OUTLINED, "selected_icon": ft.Icons.APPS, "label": "其他", "content": tools_content},
    ]

    destinations = [
        ft.NavigationRailDestination(
            icon=page["icon"],
            selected_icon=page["selected_icon"],
            label=page["label"],
            label_content=ft.Text(page["label"], font_family="sansm", size=14),
        ) for page in pages
    ]

    drag_area_title = ft.Text(pages[0]["label"], size=16, weight=ft.FontWeight.BOLD, font_family="sansr")
    drag_area = ft.WindowDragArea(
        content=ft.Container(
            content=drag_area_title,
            padding=10,
            bgcolor=ft.Colors.BLUE_50,
            alignment=ft.alignment.center,
        ),
        height=40,
        expand=False,
    )

    rail = ft.NavigationRail(
        selected_index=selected_index.current,
        label_type=ft.NavigationRailLabelType.ALL,
        min_width=50,
        min_extended_width=50,
        expand=True,
        bgcolor=ft.Colors.WHITE,
        destinations=destinations,
        on_change=lambda e: [
            setattr(selected_index, "current", e.control.selected_index),
            setattr(content_ref.current, "content", get_content()),
            setattr(drag_area_title, "value", pages[e.control.selected_index]["label"]),
            drag_area_title.update(),
            page.update(),
        ],
    )

    rail_container = ft.Container(
        content=ft.Column(
            [
                rail,
                ft.Container(
                    content=ft.Column(
                        [
                            toggle_size_button,
                            close_button,
                        ],
                        spacing=5,
                        alignment=ft.MainAxisAlignment.END,
                    ),
                    alignment=ft.alignment.bottom_left,
                    padding=10,
                    bgcolor=ft.Colors.WHITE,
                ),
            ],
            expand=True,
        ),
        width=50,
        height=page.window.height,
        bgcolor=ft.Colors.WHITE,
        border_radius=ft.border_radius.only(top_right=10, bottom_right=10),
        shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
        expand=False,
    )

    content_ref = ft.Ref[ft.AnimatedSwitcher]()
    def get_content():
        return pages[selected_index.current]["content"]

    main_content = ft.Container(
        content=ft.AnimatedSwitcher(
            content=get_content(),
            ref=content_ref,
            transition=ft.AnimatedSwitcherTransition.FADE,
            duration=0,
            reverse_duration=0,
            switch_in_curve=ft.AnimationCurve.EASE_IN_OUT,
            switch_out_curve=ft.AnimationCurve.EASE_IN_OUT,
        ),
        padding=ft.padding.only(left=10, right=10, top=10, bottom=10),
        bgcolor=ft.Colors.WHITE,
        alignment=ft.alignment.top_left,
        width=page.window.width-70,
        expand=True,
    )

    page.add(
        ft.Column(
            [
                drag_area,
                ft.Row(
                    [
                        rail_container,
                        main_content,
                    ],
                    expand=True,
                    spacing=4,
                ),
            ],
            expand=True,
            spacing=0,
        )
    )

if __name__ == "__main__":
    project_root = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath(os.path.dirname(__file__))
    check_expiration_with_ntp(project_root)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(project_root, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}", project_root)
    try:
        ft.app(target=main)
    except Exception as e:
        log(f"应用启动失败：{str(e)}", project_root)
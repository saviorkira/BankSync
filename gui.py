import os
import sys
import threading
import time
import json
import pandas as pd
from datetime import datetime
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
from ningbo_bank import run_ningbo_bank
from utils import log, read_bank_config, get_resource_path
from AI import update_ai_output, send_ai_message

def main(page: Page):
    """Flet 桌面应用主函数，带固定 NavigationRail 和美化界面"""
    # 设置窗口和主题
    page.title = "BankSync"
    page.window.title_bar_hidden = True
    page.window.title_bar_buttons_hidden = True
    page.padding = 0
    page.window.maximizable = False
    page.window.left = 100
    page.window.top = 100
    page.window.width = 580
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
        font_family="FZLanTingHei",
    )
    page.bgcolor = ft.Colors.WHITE

    # 加载自定义字体
    project_root = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath(os.path.dirname(__file__))
    font_path = os.path.join(project_root, "data", "方正兰亭准黑_GBK.ttf")
    page.fonts = {"FZLanTingHei": font_path}
    page.update()

    # UI 状态变量
    excel_data = []
    base_path = r"D:\Desktop"
    is_running = [False]
    selected_index = ft.Ref()
    selected_index.current = 0
    log_messages = []
    is_maximized = [True]

    # 页面尺寸配置
    page_sizes = {
        0: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        1: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        2: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        3: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
        4: {"min_width": 380, "min_height": 520, "max_width": 580, "max_height": 520},
    }

    # 银行处理函数映射，仅用于导出
    BANK_HANDLERS = {
        "ningbo_bank": run_ningbo_bank,
    }

    # UI 组件
    bank_dropdown = ft.Dropdown(
        label="选择银行",
        options=[ft.dropdown.Option(key="ningbo_bank", text="宁波银行")],
        value=None,
        width=page.window.width-70,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        fill_color=ft.Colors.WHITE,
        content_padding=10,
        border_color=ft.Colors.GREY_300,
        color=ft.Colors.BLACK,
        text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        label_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        tooltip="请选择要操作的银行",
    )

    start_date = ft.TextField(
        label="开始日期",
        value="2025-06-01",
        width=(page.window.width-80)/2,
        border_radius=8,
        expand=True,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        label_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
    )

    end_date = ft.TextField(
        label="结束日期",
        value="2025-07-04",
        width=(page.window.width-80)/2,
        border_radius=8,
        expand=True,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        label_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
    )

    data_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("项目名称", weight=ft.FontWeight.BOLD, size=14, font_family="FZLanTingHei")),
            ft.DataColumn(ft.Text("银行账号", weight=ft.FontWeight.BOLD, size=14, font_family="FZLanTingHei")),
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
        text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
    )

    ai_input = ft.TextField(
        label="输入消息",
        expand=True,
        border_radius=8,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="输入消息后点击发送或按 Enter",
        text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        label_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
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
        text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
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
            text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        ),
        tooltip="开始导出银行流水和回单",
        width=page.window.width-70,
    )

    base_path_text = ft.Text(
        f"下载路径: {base_path}",
        size=14,
        color=ft.Colors.GREY_700,
        font_family="FZLanTingHei",
    )

    select_path_button = ft.ElevatedButton(
        text="选择下载路径",
        icon=ft.Icons.FOLDER_OPEN,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=10,
            text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        ),
        tooltip="选择保存导出文件的目录",
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
            text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        ),
        tooltip="选择包含项目名称和银行账号的 Excel 文件",
        width=page.window.width-70,
    )

    # 日志更新函数
    def update_log(msg):
        log_messages.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: {msg}\n")
        log_area.value = "".join(log_messages)
        if selected_index.current == 4 and hasattr(log_area, 'page') and log_area.page is not None:
            log_area.update()
            page.scroll_to(key="log_area", duration=500)
            log(f"log_area 更新: 日志页面已渲染", project_root)
        log(msg, project_root)

    # 银行选择事件
    def on_bank_select(e):
        if bank_dropdown.value:
            update_log(f"已选择银行: 宁波银行")
        else:
            update_log("银行选择已清空")

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
                        ft.DataCell(ft.Text(project, size=14, font_family="FZLanTingHei")),
                        ft.DataCell(ft.Text(account, size=14, font_family="FZLanTingHei")),
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
        if bank_dropdown.value not in BANK_HANDLERS:
            update_log("错误: 仅支持已注册的银行")
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
        update_log(f"开始运行 宁波银行 导出...")

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
                if not os.path.exists(base_path) or not os.access(base_path, os.W_OK):
                    update_log(f"错误: 下载路径不可访问或不可写: {base_path}")
                    return
                with sync_playwright() as playwright:
                    BANK_HANDLERS[bank_dropdown.value](playwright, project_root, base_path, excel_data, kaishi, jieshu, log_callback=update_log)
                update_log("导出完成。")
            except Exception as ex:
                update_log(f"执行出错：{str(ex)}")
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
    file_picker = ft.FilePicker(on_result=import_excel)
    dir_picker = ft.FilePicker(on_result=select_base_path)
    page.overlay.extend([file_picker, dir_picker])

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
        nonlocal is_maximized
        current_page = selected_index.current
        sizes = page_sizes.get(current_page, page_sizes[0])
        steps = 15
        duration = 0.0133
        start_width = page.window.width
        start_height = page.window.height
        if is_maximized[0]:
            target_width = sizes["min_width"]
            target_height = sizes["min_height"]
            toggle_size_button.icon = ft.Icons.FULLSCREEN
        else:
            target_width = sizes["max_width"]
            target_height = sizes["max_height"]
            toggle_size_button.icon = ft.Icons.FULLSCREEN_EXIT

        for i in range(steps + 1):
            t = i / steps
            page.window.width = start_width + (target_width - start_width) * t
            page.window.height = start_height + (target_height - start_height) * t
            time.sleep(duration)

        bank_dropdown.width = page.window.width - 70
        start_date.width = (page.window.width - 80) / 2
        end_date.width = (page.window.width - 80) / 2
        run_button.width = page.window.width - 70
        select_path_button.width = page.window.width - 70
        import_excel_button.width = page.window.width - 70
        bank_export_content.controls[1].width = page.window.width - 70
        bank_export_content.controls[5].width = page.window.width - 70
        ai_content.controls[0].width = page.window.width - 70
        ai_content.controls[1].width = page.window.width - 70
        todo_content.controls[0].width = page.window.width - 70
        settings_content.controls[0].width = page.window.width - 70
        main_content.width = page.window.width - 70
        page.update()

        is_maximized[0] = not is_maximized[0]
        toggle_size_button.update()

    # 登录页面
    home_content = ft.Column(
        [
            ft.Text("网站图标区域", size=14, font_family="FZLanTingHei"),
            ft.Row(
                controls=load_site_icons(project_root, update_log, lambda name: login_site(name, project_root, update_log, last_click_time=[0])),
                wrap=True,
                spacing=10,
                run_spacing=10,
                alignment=ft.MainAxisAlignment.CENTER,
            ),
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
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
                width=page.window.width-70,
                height=173,
            ),
            run_button,
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
            ),
        ],
        spacing=10,
        scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
    )

    pages = [
        {"icon": ft.Icons.HOME_OUTLINED, "selected_icon": ft.Icons.HOME, "label": "登录", "content": home_content},
        {"icon": ft.Icons.DOWNLOAD_OUTLINED, "selected_icon": ft.Icons.DOWNLOAD, "label": "导出", "content": bank_export_content},
        {"icon": ft.Icons.CHAT_OUTLINED, "selected_icon": ft.Icons.CHAT, "label": "AI", "content": ai_content},
        {"icon": ft.Icons.TASK_OUTLINED, "selected_icon": ft.Icons.TASK_ROUNDED, "label": "待办", "content": todo_content},
    ]

    destinations = [
        ft.NavigationRailDestination(
            icon=page["icon"],
            selected_icon=page["selected_icon"],
            label=page["label"],
            label_content=ft.Text(page["label"], font_family="FZLanTingHei", size=14),
        ) for page in pages
    ]
    destinations.append(
        ft.NavigationRailDestination(
            icon=ft.Icons.SETTINGS_OUTLINED,
            selected_icon=ft.Icons.SETTINGS,
            label="日志",
            label_content=ft.Text("日志", font_family="FZLanTingHei", size=14),
        )
    )

    drag_area_title = ft.Text(pages[0]["label"], size=16, weight=ft.FontWeight.BOLD, font_family="FZLanTingHei")
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
            setattr(drag_area_title, "value", "日志" if e.control.selected_index == 4 else pages[e.control.selected_index]["label"]),
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
        if selected_index.current < len(pages):
            return pages[selected_index.current]["content"]
        else:
            return settings_content

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
        padding=ft.padding.symmetric(vertical=10, horizontal=10),
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
        log(f"应用启动失败：{str(e)}", base_path)
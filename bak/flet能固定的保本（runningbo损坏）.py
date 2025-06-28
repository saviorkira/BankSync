import os
import sys
import threading
import configparser
import pandas as pd
import numpy as np
import cv2
import time
from datetime import datetime
from playwright.sync_api import sync_playwright, Playwright
import pyautogui
from pywinauto import Desktop
from pywinauto.application import Application
import flet as ft
from flet import Page, FilePicker, FilePickerResultEvent, Container, border, Colors, Column, Row, Text, ElevatedButton, Dropdown, DataTable, DataColumn, DataRow, DataCell, TextField, ProgressRing, SnackBar, Icons
import win32gui
import win32con
import ctypes

# 设置 DPI 感知
ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 每监视器 DPI 感知

def force_check_expiration_local(expire_date_str="2026-01-01"):
    """检查程序是否过期，过期则退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期：{expire_date_str}）")
        sys.exit()

def log(msg: str, base_path: str = r"D:\Data", log_callback=None):
    """记录日志到文件和界面"""
    print(msg)
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")
    if log_callback:
        log_callback(msg)

def get_resource_path(relative_path: str, subfolder: str = "data") -> str:
    """获取资源文件路径，支持开发和打包环境"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    full_path = os.path.join(base_path, subfolder, relative_path)
    log(f"检查路径: {full_path}, 是否存在: {os.path.exists(full_path)}", base_path)
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    full_path = os.path.join(base_path, relative_path)
    log(f"尝试备用路径: {full_path}, 是否存在: {os.path.exists(full_path)}", base_path)
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    return full_path

def read_bank_config(base_path: str):
    """读取宁波银行配置"""
    config_path = get_resource_path("../config.txt")
    log(f"加载配置文件: {config_path}", base_path)
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件未找到: {config_path}")
    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')
    if "ningbo_bank" not in config:
        raise KeyError("配置文件缺少 [ningbo_bank] 部分")
    bank_conf = config["ningbo_bank"]
    username = bank_conf.get("username")
    password = bank_conf.get("password")
    login_url = bank_conf.get("login_url")
    if not all([username, password, login_url]):
        raise ValueError("宁波银行配置不完整")
    return username, password, login_url, config_path

def find_and_click_image(template_path: str, offset_x: int = 0, offset_y: int = 0, threshold: float = 0.5,
                         max_attempts: int = 10) -> tuple:
    """通过模板匹配找到并点击屏幕上的图片"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        if not os.path.exists(template_path):
            log(f"模板图片不存在: {template_path}")
            return None
        try:
            template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception as e:
            log(f"加载模板图片失败: {template_path}, 错误: {e}")
            return None
        if template is None:
            log(f"无法加载模板图片: {template_path}")
            return None
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            x, y = max_loc
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)
            log(f"尝试 {attempt + 1} 成功，点击位置: ({x + offset_x}, {y + offset_y})")
            return (x, y)
        time.sleep(1)
    log(f"未找到模板: {template_path}, 尝试次数: {max_attempts}")
    return None

def handle_overwrite_dialog():
    """处理文件覆盖对话框"""
    try:
        app = Desktop(backend="win32")
        dialog = app.window(title_re=".*文件已存在.*|.*确认保存.*|.*确认另存为.*|.*Confirm Save.*|.*Replace.*")
        dialog.wait("exists ready", timeout=5)
        dialog.set_focus()
        replace_btn = dialog.child_window(title_re="是.*|替换.*|Yes.*|Replace.*", class_name="Button")
        replace_btn.wait("exists enabled visible ready", timeout=3)
        replace_btn.click()
        time.sleep(1)
    except Exception as e:
        log(f"未检测到覆盖对话框或点击失败: {e}")

def handle_save_dialog(save_path: str, pdf_filename: str):
    """处理另存为对话框"""
    full_path = os.path.join(save_path, pdf_filename)
    log(f"尝试处理另存文件对话框，目标路径: {full_path}")
    try:
        desktop = Desktop(backend="win32")
        dialogs = desktop.windows(title_re="^另存为$")
        if not dialogs:
            raise Exception("未找到 '另存为' 对话框")
        for i, dialog in enumerate(dialogs):
            try:
                handle = dialog.handle
                app = Application(backend="win32").connect(handle=handle)
                dlg = app.window(handle=handle)
                dlg.set_focus()
                time.sleep(0.3)
                edit = dlg.child_window(class_name="Edit")
                edit.set_focus()
                edit.set_edit_text(full_path)
                log(f"对话框 {i + 1} 设置路径成功: {full_path}")
                time.sleep(0.5)
                save_btn = dlg.child_window(class_name="Button", title_re="保存|Save")
                save_btn.click()
                log("已点击保存按钮")
                time.sleep(2)
                handle_overwrite_dialog()
                return
            except Exception as inner_e:
                log(f"对话框 {i + 1} 处理失败: {inner_e}")
                continue
        raise Exception("找不到可用的 '另存为' 对话框")
    except Exception as e:
        log(f"处理另存对话框失败: {e}")
        raise

def run_ningbo_bank(playwright: Playwright, base_path: str, projects_accounts: list, kaishiriqi: str, jieshuriqi: str,
                    log_callback=None):
    """执行宁波银行流水、回单和对账单导出"""
    log("进入 run_ningbo_bank 函数", base_path, log_callback)

    def log_local(msg):
        log(msg, base_path, log_callback)

    log_local("开始宁波银行处理...")
    log_local(f"Playwright 浏览器路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(base_path)
        log_local(f"已加载配置文件: {config_path}")
    except Exception as e:
        log_local(f"加载 config.txt 失败: {str(e)}")
        raise
    browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
    if not os.path.exists(browser_path):
        log_local(f"Playwright 浏览器路径不存在: {browser_path}")
        raise FileNotFoundError(f"浏览器路径不存在: {browser_path}")
    try:
        log_local("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(90000)
        log_local(f"登录页面导航: {login_url}")
        page.goto(login_url)
        log_local("输入用户名和密码...")
        page.get_by_role("textbox", name="用户名").fill(username)
        page.get_by_role("textbox", name="请输入您的密码").fill(password)
        log_local("等待账户管理页面加载...")
        page.wait_for_selector('text=账户管理', timeout=90000)
        page.get_by_role("link", name="账户管理").click()
        page.get_by_role("link", name="账户明细").click()
        previous_xiangmu = None
        for index, (xiangmu, account) in enumerate(projects_accounts):
            log_local(f"处理项目: {xiangmu}, 账户: {account}")
            try:
                duizhang_path = os.path.join(base_path, xiangmu, "银行流水")
                huidan_path = os.path.join(base_path, xiangmu, "银行回单")
                duizhangdan_path = os.path.join(base_path, xiangmu, "银行对账单")
                os.makedirs(duizhang_path, exist_ok=True)
                os.makedirs(huidan_path, exist_ok=True)
                os.makedirs(duizhangdan_path, exist_ok=True)
                if index == 0:
                    page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
                    page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").fill(account)
                    log_local(f"搜索账户: {account}")
                    page.wait_for_timeout(1000)
                    try:
                        page.get_by_role("listitem").filter(has_text=xiangmu).locator("span").nth(2).click()
                    except Exception as e:
                        log_local(f"点击搜索结果失败 (项目: {xiangmu}, 账户: {account}): {str(e)}")
                        page.screenshot(path=os.path.join(base_path, f"error_select_{xiangmu}.png"))
                        continue
                    page.get_by_text("展开").first.click()
                    page.get_by_role("textbox", name="开始日期").fill(kaishiriqi)
                    page.get_by_role("textbox", name="开始日期").press("Enter")
                    page.get_by_role("textbox", name="结束日期").fill(jieshuriqi)
                    page.get_by_role("textbox", name="结束日期").press("Enter")
                else:
                    # 后续项目：简化三段式逻辑
                    page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
                    page.get_by_role("textbox", name=f"- {previous_xiangmu}").fill(account)
                    log(f"使用银行账号查询：{account}")
                    page.wait_for_timeout(1000)  # 等待结果加载
                    try:
                        page.get_by_role("link", name=account).click()
                    except Exception as e:
                        log(f"点击链接失败（项目：{xiangmu}，账号：{account}）：{str(e)}")
                        page.screenshot(path=f"error_select_{xiangmu}.png")
                        log(f"已保存选择错误截图：error_select_{xiangmu}.png")
                        continue
                # 查询
                page.get_by_role("button", name=" 查询").click()
                page.wait_for_timeout(1000)  # 等待查询结果加载

                if not checkbox.is_visible() and checkbox.is_enabled():
                    log_local(f"无回单或流水数据，跳过 (项目: {xiangmu}, 账户: {account})")
                    page.screenshot(path=os.path.join(base_path, f"error_no_data_{xiangmu}.png"))
                    continue
                try:
                    if not checkbox.is_checked():
                        checkbox.check()
                    log_local("已勾选复选框")
                except Exception as e:
                    log_local(f"复选框选择失败 (项目: {xiangmu}, 账户: {account}): {e}")
                    page.screenshot(path=os.path.join(base_path, f"error_{xiangmu}_checkbox.png"))
                    continue
                # 导出回单
                page.get_by_role("button", name="export").click()
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
                            filename = f"{xiangmu}_银行_回单_{kaishiriqi}_{jieshuriqi}.pdf"
                            download.save_as(os.path.join(huidan_path, filename))
                            log_local(f"已导出银行回单: {filename}")
                            found = True
                            break
                    if not found:
                        log_local(f"未找到 '凭证导出' 菜单项 (项目: {xiangmu})")
                        page.screenshot(path=os.path.join(base_path, f"error_menu_{xiangmu}.png"))
                except Exception as e:
                    log_local(f"银行回单导出失败: {str(e)}")
                # 导出流水
                page.get_by_role("button", name="导出").click()
                with page.expect_download() as download_info:
                    page.get_by_text("对账单导出", exact=True).click()
                download = download_info.value
                filename = f"{xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
                download.save_as(os.path.join(duizhang_path, filename))
                log_local(f"已导出流水: {filename}")
                # 打印对账单为 PDF
                try:
                    page.get_by_role("button", name="打印").click()
                    page.wait_for_timeout(2000)
                    menus = page.locator("css=[id^='dropdown-menu-']")
                    count = menus.count()
                    found = False
                    for i in range(count):
                        item = menus.nth(i).filter(has_text="对账单打印")
                        if item.is_visible():
                            item.click()
                            log_local(f"已点击对账单打印菜单项")
                            found = True
                            break
                    if not found:
                        log_local(f"未找到 '对账单打印' 菜单项 (项目: {xiangmu})")
                        page.screenshot(path=os.path.join(base_path, f"error_print_menu_{xiangmu}.png"))
                        continue
                    log_local("等待 Chrome 打印对话框...")
                    time.sleep(5)
                    target_printer_path = get_resource_path("target_printer.bmp", subfolder="seek")
                    save_as_pdf_default_path = os.path.join(base_path, "save_as_pdf_default.jpg")
                    log_local(f"检查默认保存PDF路径: {save_as_pdf_default_path}")
                    save_as_pdf_hover_path = os.path.join(base_path, "save_as_pdf_hover.jpg")
                    save_button_path = os.path.join(base_path, "save_button.jpg")
                    if not (os.path.exists(target_printer_path) and os.path.exists(save_as_pdf_default_path) and os.path.exists(save_as_pdf_hover_path) and os.path.exists(save_button_path)):
                        log_local("模板图片缺失或为空")
                        raise FileNotFoundError("请在 seek 文件夹准备模板图片")
                    log_local("定位 'Target Printer'...")
                    target_pos = find_and_click_image(target_printer_path)
                    if not target_pos:
                        log_local("未找到 'Target Printer' 文字")
                        pyautogui.screenshot(os.path.join(base_path, f"error_target_printer_{xiangmu}.png"))
                        continue
                    x_target, y_target = target_pos
                    x_offset = 250
                    pyautogui.click(x_target + x_offset, y_target)
                    log_local(f"点击偏移位置: ({x_target + x_offset}, {y_target})")
                    time.sleep(2)
                    pdf_clicked = False
                    for attempt in range(3):
                        pyautogui.moveTo(x_target + x_offset, y_target + 20)
                        time.sleep(0.5)
                        if find_and_click_image(save_as_pdf_default_path):
                            log_local("已点击默认 'Save as PDF' 按钮")
                            pdf_clicked = True
                            time.sleep(2)
                            break
                        elif find_and_click_image(save_as_pdf_hover_path):
                            log_local("已点击悬停 'Save as PDF' 按钮")
                            pdf_clicked = True
                            time.sleep(2)
                            break
                    if not pdf_clicked:
                        log_local("未找到 'Save as PDF' 按钮")
                        pyautogui.screenshot(os.path.join(base_path, f"error_save_as_pdf_{xiangmu}.png"))
                        continue
                    if find_and_click_image(save_button_path):
                        log_local("已点击 'Save' 按钮")
                        time.sleep(2)
                    else:
                        log_local("未找到 'Save' 按钮")
                        pyautogui.screenshot(os.path.join(base_path, f"error_save_button_{xiangmu}.png"))
                        continue
                    pdf_filename = f"{xiangmu}_对账单打印_{kaishiriqi}_{jieshuriqi}.pdf"
                    handle_save_dialog(duizhangdan_path, pdf_filename)
                    pdf_path = os.path.join(duizhangdan_path, pdf_filename)
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                        log_local(f"已打印对账单 PDF: {pdf_path}")
                    else:
                        log_local(f"PDF 文件生成失败或为空: {pdf_path}")
                except Exception as e:
                    log_local(f"打印对账单 PDF 失败 (项目: {xiangmu}): {str(e)}")
                    page.screenshot(path=os.path.join(base_path, f"error_print_{xiangmu}.png"))
                previous_xiangmu = xiangmu
            except Exception as e:
                log_local(f"导出失败 (项目: {xiangmu}): {str(e)}")
                continue
        context.close()
        browser.close()
    except Exception as e:
        log_local(f"初始化失败: {str(e)}")
        raise
    finally:
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()

def fix_window_size(page: Page):
    """通过 Windows API 强制固定窗口大小"""
    log("应用 fix_window_size 以固定窗口大小")
    time.sleep(0.5)  # 延迟确保窗口初始化
    hwnd = win32gui.GetForegroundWindow()
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    style &= ~(win32con.WS_SIZEBOX | win32con.WS_MAXIMIZEBOX)
    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
    win32gui.MoveWindow(hwnd, 0, 0, 600, 700, True)
    page.window_width = 600
    page.window_height = 700
    page.update()

def main(page: Page):
    """Flet 桌面应用主函数"""
    log(f"Python 版本: {sys.version}")
    page.title = "银行流水回单导出"
    page.window_width = 100
    page.window_height = 100
    page.window_resizable = False
    page.window_maximizable = False
    page.window_minimizable = True
    page.window_min_width = 100
    page.window_min_height = 100
    page.window_max_width = 100
    page.window_max_height = 100
    page.padding = 10
    page.theme = ft.Theme(color_scheme=ft.ColorScheme(primary=ft.Colors.BLUE_700, secondary=ft.Colors.BLUE_100))
    page.scroll = ft.ScrollMode.AUTO

    # 绑定窗口调整事件
    def on_resize(e):
        fix_window_size(page)
    page.on_resize = on_resize

    # ========== 字体加载 ==========
    def get_font_path():
        if getattr(sys, 'frozen', False):
            return os.path.join(sys._MEIPASS, "data", "方正兰亭准黑_GBK.ttf")
        else:
            return os.path.join(os.path.abspath(os.path.dirname(__file__)), "data", "方正兰亭准黑_GBK.ttf")

    font_path = get_font_path()
    if os.path.exists(font_path):
        page.fonts = {"FangZhengLanTingZhunHei": font_path}
        log(f"加载字体文件: {font_path}")
        page.font = "FangZhengLanTingZhunHei"
    else:
        log(f"字体文件未找到: {font_path}，使用默认字体")

    # ========== 图标加载 ==========
    def load_icon_or_default(path: str, fallback_icon: str):
        """如果图标文件存在，则加载图片，否则使用默认图标"""
        if os.path.exists(path):
            return ft.Image(src=path, width=20, height=20)
        else:
            return ft.Icon(fallback_icon, size=20)

    excel_icon = load_icon_or_default(get_resource_path("excel_icon.png"), ft.Icons.TABLE_CHART)
    folder_icon = load_icon_or_default(get_resource_path("folder_icon.png"), ft.Icons.FOLDER)

    # ========== UI 状态变量 ==========
    excel_data = []
    base_path = r"D:\Data"

    start_date = ft.TextField(
        label="开始日期",
        value="2025-03-01",
        width=175,
        height=40,
        border_radius=6,
        tooltip="输入开始日期 (格式: YYYY-MM-DD)",
        hint_text="YYYY-MM-DD",
    )
    end_date = ft.TextField(
        label="结束日期",
        value="2025-03-31",
        width=175,
        height=40,
        border_radius=6,
        tooltip="输入结束日期 (格式: YYYY-MM-DD)",
        hint_text="YYYY-MM-DD",
    )

    bank_dropdown = ft.Dropdown(
        label="选择银行",
        options=[ft.dropdown.Option("Ningbo Bank", "宁波银行")],
        value=None,
        width=360,
        border_radius=6,
        dense=True,
        tooltip="选择要导出的银行",
    )

    data_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("项目名称", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("银行账户", weight=ft.FontWeight.BOLD)),
        ],
        rows=[],
        border=ft.border.all(1, ft.Colors.GREY_400),
        width=360,
    )

    log_area = ft.TextField(
        multiline=True,
        min_lines=4,
        max_lines=6,
        read_only=True,
        width=360,
        border_radius=6,
        bgcolor=ft.Colors.GREY_100,
        tooltip="显示导出过程中的日志信息",
    )

    run_button = ft.ElevatedButton(
        text="运行",
        disabled=False,
        width=360,
        height=40,
        color=ft.Colors.WHITE,
        bgcolor=ft.Colors.BLUE_700,
        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6), elevation=2),
        tooltip="点击开始导出银行流水和回单",
    )

    progress_ring = ft.ProgressRing(visible=False, width=20, height=20, color=ft.Colors.BLUE_700)

    is_running = [False]

    # ========== UI 交互逻辑 ==========
    def show_snackbar(message, is_error=False):
        page.snack_bar = SnackBar(
            content=Text(message, color=ft.Colors.WHITE if is_error else ft.Colors.BLACK),
            bgcolor=ft.Colors.RED_600 if is_error else ft.Colors.GREEN_600,
            duration=3000,
        )
        page.snack_bar.open = True
        page.update()

    def update_log(msg):
        log_area.value += f"{msg}\n"
        log_area.update()

    def on_bank_select(e):
        if bank_dropdown.value:
            update_log(f"已选择银行: {bank_dropdown.options[0].text if bank_dropdown.value == 'Ningbo Bank' else '未知'}")
            show_snackbar(f"已选择: {bank_dropdown.options[0].text}", is_error=False)
        else:
            update_log("银行选择已清空")
            show_snackbar("银行选择已清空", is_error=True)

    bank_dropdown.on_change = on_bank_select

    def import_excel(e: FilePickerResultEvent):
        if e.files and any(f.name.endswith(('.xlsx', '.xls')) for f in e.files):
            file_path = e.files[0].path
            try:
                df = pd.read_excel(file_path, header=0)
                if df.empty or len(df.columns) < 2:
                    update_log("错误: Excel 文件格式错误，至少需要两列（项目名称和银行账户）")
                    show_snackbar("Excel 格式错误", is_error=True)
                    return
                nonlocal excel_data
                excel_data.clear()
                excel_data.extend([(str(project).strip(), str(account).strip()) for project, account in
                                   zip(df.iloc[:, 0], df.iloc[:, 1]) if str(project).strip() and str(account).strip()])
                data_table.rows = [
                    ft.DataRow(cells=[
                        ft.DataCell(ft.Text(project)),
                        ft.DataCell(ft.Text(account)),
                    ]) for project, account in excel_data
                ]
                update_log(f"成功导入 Excel 文件: {file_path}")
                show_snackbar(f"成功导入: {file_path}", is_error=False)
                page.update()
            except Exception as ex:
                update_log(f"导入 Excel 失败: {str(ex)}")
                show_snackbar(f"导入失败: {str(ex)}", is_error=True)
        else:
            update_log("错误: 请选择有效的 Excel 文件")
            show_snackbar("请选择有效的 Excel 文件", is_error=True)

    def select_base_path(e: FilePickerResultEvent):
        if e.path:
            nonlocal base_path
            base_path = e.path
            base_path_field.value = base_path
            update_log(f"已选择路径: {base_path}")
            show_snackbar(f"已选择路径: {base_path}", is_error=False)
            page.update()

    def run_export(e):
        nonlocal is_running
        if is_running[0]:
            update_log("提示: 导出进程正在运行，请等待")
            show_snackbar("导出进程正在运行", is_error=True)
            return
        if not bank_dropdown.value:
            update_log("错误: 请先选择银行")
            show_snackbar("请先选择银行", is_error=True)
            return
        if bank_dropdown.value != "Ningbo Bank":
            update_log("错误: 仅支持宁波银行")
            show_snackbar("仅支持宁波银行", is_error=True)
            return
        if not excel_data:
            update_log("提示: 请先导入 Excel 文件")
            show_snackbar("请先导入 Excel 文件", is_error=True)
            return
        kaishi = start_date.value.strip()
        jieshu = end_date.value.strip()
        nonlocal base_path
        base_path = base_path_field.value.strip()
        if not (kaishi, jieshu, base_path):
            update_log("提示: 请填写完整日期和下载路径")
            show_snackbar("请填写完整日期和路径", is_error=True)
            return
        is_running[0] = True
        run_button.disabled = True
        run_button.text = "运行中..."
        progress_ring.visible = True
        run_button.content = Row([Text("运行中..."), progress_ring], alignment=ft.MainAxisAlignment.CENTER)
        run_button.update()
        update_log("开始运行宁波银行导出...")

        def worker():
            try:
                with sync_playwright() as playwright:
                    run_ningbo_bank(playwright, base_path, excel_data, kaishi, jieshu, log_callback=update_log)
                update_log("导出完成")
                show_snackbar("导出完成", is_error=False)
            except Exception as ex:
                update_log(f"执行出错: {str(ex)}")
                show_snackbar(f"执行出错: {str(ex)}", is_error=True)
            finally:
                nonlocal is_running
                is_running[0] = False
                run_button.disabled = False
                run_button.text = "运行"
                run_button.content = Text("运行")
                progress_ring.visible = False
                run_button.update()

        threading.Thread(target=worker, daemon=True).start()

    run_button.on_click = run_export

    # 文件选择器
    file_picker = FilePicker(on_result=import_excel)
    dir_picker = FilePicker(on_result=select_base_path)
    page.overlay.extend([file_picker, dir_picker])

    # 下载路径文本框
    base_path_field = ft.TextField(
        label="下载路径",
        value=base_path,
        width=360,
        height=40,
        border_radius=6,
        tooltip="输入或选择下载路径",
    )

    # 创建统一样式的图标按钮
    def create_icon_button(label, icon, on_click, width):
        return ft.ElevatedButton(
            content=ft.Row([icon, ft.Text(label, size=12)], spacing=6),
            width=width,
            height=40,
            style=ft.ButtonStyle(
                shape=ft.RoundedRectangleBorder(radius=6),
                bgcolor=ft.Colors.BLUE_50,
                elevation=2,
            ),
            tooltip=f"点击{label}",
            on_click=on_click,
        )

    # 页面主内容
    page.add(
        ft.Card(
            content=Container(
                content=Column(
                    controls=[
                        bank_dropdown,
                        Row(
                            controls=[
                                create_icon_button(
                                    "导入 Excel",
                                    excel_icon,
                                    lambda _: file_picker.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls"]),
                                    width=180
                                ),
                                create_icon_button(
                                    "下载路径",
                                    folder_icon,
                                    lambda _: dir_picker.get_directory_path(),
                                    width=170
                                ),
                            ],
                            spacing=10,
                            alignment=ft.MainAxisAlignment.START,
                        ),
                        base_path_field,
                        Row(
                            controls=[
                                start_date,
                                end_date,
                            ],
                            spacing=10,
                            alignment=ft.MainAxisAlignment.START,
                        ),
                        run_button,
                        ft.Text("项目列表", size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                        data_table,
                        ft.Text("导出日志", size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                        log_area,
                    ],
                    spacing=10,
                    alignment=ft.MainAxisAlignment.START,
                    horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                ),
                padding=10,
            ),
            elevation=5,
            surface_tint_color=ft.Colors.WHITE,
            margin=5,
        )
    )

    page.update()
    fix_window_size(page)

if __name__ == "__main__":
    force_check_expiration_local()
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}")
    try:
        ft.app(target=main)
    except Exception as e:
        log(f"应用启动失败: {e}")
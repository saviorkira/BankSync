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
import win32gui
import win32con
import ctypes
import platform
from flask import Flask, request, jsonify, send_from_directory
from flask_sock import Sock
import webview
import tempfile

# 设置 DPI 感知
ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 每监视器 DPI 感知

app = Flask(__name__, static_folder='data', template_folder='data')
sock = Sock(app)

# WebSocket for real-time logging
connected_clients = []

@sock.route('/ws')
def websocket(ws):
    connected_clients.append(ws)
    try:
        while True:
            ws.receive()  # Keep connection alive
    except:
        connected_clients.remove(ws)

def broadcast_log(msg):
    for client in connected_clients:
        try:
            client.send(msg)
        except:
            connected_clients.remove(client)

def force_check_expiration_local(expire_date_str="2026-01-01"):
    """检查程序是否过期，过期则退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期：{expire_date_str})")
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
    config_path = get_resource_path("config.txt")
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
                    page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
                    page.get_by_role("textbox", name=f"- {previous_xiangmu}").fill(account)
                    log_local(f"搜索账户: {account}")
                    page.wait_for_timeout(1000)
                    try:
                        page.get_by_role("link", name=account).click()
                    except Exception as e:
                        log_local(f"点击链接失败 (项目: {xiangmu}, 账户: {account}): {str(e)}")
                        page.screenshot(path=os.path.join(base_path, f"error_select_{xiangmu}.png"))
                        continue

                # 查询
                page.get_by_role("button", name=" 查询").click()
                page.wait_for_timeout(1000)  # 等待查询结果加载


                page.wait_for_selector('role=checkbox[name="Toggle Selection of All Rows"]"', timeout=10000)
                checkbox = page.get_by_role("checkbox", name="Toggle Selection of All Rows")
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
                    page.get_by_text("对账单导出").click()
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
                    save_as_pdf_hover_path = os.path.join(base_path, "save_as_pdf_hover.jpg")
                    save_button_path = os.path.join(base_path, "save_button.jpg")
                    if not (os.path.exists(target_printer_path) and os.path.exists(
                            save_as_pdf_default_path) and os.path.exists(save_as_pdf_hover_path) and os.path.exists(
                            save_button_path)):
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

@app.route('/')
def index():
    return send_from_directory('data', 'index.html')

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    try:
        log("收到上传 Excel 请求", base_path=r"D:\Data", log_callback=broadcast_log)
        if 'file' not in request.files:
            log("未收到文件上传请求", base_path=r"D:\Data", log_callback=broadcast_log)
            return jsonify({'success': False, 'error': '未收到文件上传请求'})
        file = request.files['file']
        if not file or not file.filename:
            log("文件为空或文件名缺失", base_path=r"D:\Data", log_callback=broadcast_log)
            return jsonify({'success': False, 'error': '请选择有效的 Excel 文件'})
        if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            log(f"文件格式不支持: {file.filename}", base_path=r"D:\Data", log_callback=broadcast_log)
            return jsonify({'success': False, 'error': '请选择有效的 Excel 文件 (.xlsx 或 .xls)'})
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            file.save(tmp.name)
            df = pd.read_excel(tmp.name, header=0)
        os.unlink(tmp.name)
        if df.empty or len(df.columns) < 2:
            log("Excel 文件格式错误，至少需要两列", base_path=r"D:\Data", log_callback=broadcast_log)
            return jsonify({'success': False, 'error': 'Excel 文件格式错误，至少需要两列（项目名称和银行账户）'})
        excel_data = [(str(project).strip(), str(account).strip()) for project, account in
                      zip(df.iloc[:, 0], df.iloc[:, 1]) if str(project).strip() and str(account).strip()]
        log(f"Excel 文件上传成功，数据行数: {len(excel_data)}", base_path=r"D:\Data", log_callback=broadcast_log)
        return jsonify({'success': True, 'data': excel_data})
    except Exception as e:
        log(f"上传 Excel 失败: {str(e)}", base_path=r"D:\Data", log_callback=broadcast_log)
        return jsonify({'success': False, 'error': str(e)})

@app.route('/run_export', methods=['POST'])
def run_export():
    try:
        data = request.json
        bank = data.get('bank')
        start_date = data.get('startDate')
        end_date = data.get('endDate')
        base_path = data.get('basePath')
        excel_data = data.get('excelData')
        if not all([bank, start_date, end_date, base_path, excel_data]):
            log("请填写完整日期、路径和 Excel 数据", base_path=r"D:\Data", log_callback=broadcast_log)
            return jsonify({'success': False, 'error': '请填写完整日期、路径和 Excel 数据'})
        if bank != 'Ningbo Bank':
            log("仅支持宁波银行", base_path=r"D:\Data", log_callback=broadcast_log)
            return jsonify({'success': False, 'error': '仅支持宁波银行'})
        with sync_playwright() as playwright:
            run_ningbo_bank(playwright, base_path, excel_data, start_date, end_date, log_callback=broadcast_log)
        log("导出任务完成", base_path=r"D:\Data", log_callback=broadcast_log)
        return jsonify({'success': True})
    except Exception as e:
        log(f"执行出错: {str(e)}", base_path=r"D:\Data", log_callback=broadcast_log)
        return jsonify({'success': False, 'error': str(e)})

class Api:
    def pick_directory(self):
        try:
            log("调用 pick_directory，尝试打开文件夹选择对话框", base_path=r"D:\Data", log_callback=broadcast_log)
            if not webview.windows:
                log("窗口列表为空，无法打开文件夹对话框", base_path=r"D:\Data", log_callback=broadcast_log)
                return ''
            path = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
            log(f"文件夹选择结果: {path}", base_path=r"D:\Data", log_callback=broadcast_log)
            return path or ''
        except Exception as e:
            log(f"选择文件夹失败: {str(e)}", base_path=r"D:\Data", log_callback=broadcast_log)
            return ''

def fix_window_size(window):
    hwnd = None
    def enum_windows_callback(hwnd, results):
        title = win32gui.GetWindowText(hwnd)
        if "银行流水回单导出" in title:
            results.append(hwnd)
    for attempt in range(10):
        hwnd_list = []
        win32gui.EnumWindows(enum_windows_callback, hwnd_list)
        if hwnd_list:
            hwnd = hwnd_list[0]
            log(f"找到窗口句柄: {hwnd}, 标题: {win32gui.GetWindowText(hwnd)}", base_path=r"D:\Data")
            break
        log(f"尝试 {attempt + 1}/10: 未找到应用窗口，等待 0.2 秒", base_path=r"D:\Data")
        time.sleep(0.2)
    if not hwnd:
        log("错误: 无法找到应用窗口，跳过窗口调整", base_path=r"D:\Data")
        return
    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    style &= ~(win32con.WS_SIZEBOX | win32con.WS_MAXIMIZEBOX)  # 禁用调整大小和最大化
    style |= win32con.WS_MINIMIZEBOX  # 显式启用最小化按钮
    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
    screen_width = ctypes.windll.user32.GetSystemMetrics(0)
    screen_height = ctypes.windll.user32.GetSystemMetrics(1)
    window_width = 600
    window_height = 700
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    win32gui.SetWindowPos(hwnd, None, x, y, window_width, window_height,
                          win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
    window.resize(window_width, window_height)
    window.move(x, y)
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
    # 再次验证窗口样式
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    rect = win32gui.GetWindowRect(hwnd)
    actual_width = rect[2] - rect[0]
    actual_height = rect[3] - rect[1]
    log(f"窗口大小验证: {actual_width}x{actual_height}, 位置: ({x}, {y})", base_path=r"D:\Data")
    log(f"窗口样式验证: WS_SIZEBOX={bool(style & win32con.WS_SIZEBOX)}, WS_MAXIMIZEBOX={bool(style & win32con.WS_MAXIMIZEBOX)}, WS_MINIMIZEBOX={bool(style & win32con.WS_MINIMIZEBOX)}", base_path=r"D:\Data")

if __name__ == "__main__":
    force_check_expiration_local()
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}", base_path=r"D:\Data")
    threading.Thread(target=lambda: app.run(host='127.0.0.1', port=5000), daemon=True).start()
    time.sleep(1)  # Wait for Flask to start
    api = Api()
    window = webview.create_window(
        title='银行流水回单导出',
        url='http://127.0.0.1:5000',
        width=600,
        height=700,
        resizable=False,
        js_api=api
    )
    window.events.loaded += lambda: fix_window_size(window)
    try:
        webview.start()
    except Exception as e:
        log(f"应用启动失败: {e}", base_path=r"D:\Data")
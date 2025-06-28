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
from flet import Page, FilePicker, FilePickerResultEvent

def force_check_expiration_local(expire_date_str="2026-06-01"):
    """使用本地系统时间判断是否过期，过期则退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期为 {expire_date_str}）。")
        sys.exit(0)

def log(msg, base_path=r"D:\Desktop", log_callback=None):
    """记录日志到文件和UI"""
    print(msg)
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")
    if log_callback:
        log_callback(msg)

def get_resource_path(relative_path, subfolder="seek"):
    """获取资源路径，优先检查 seek 子目录"""
    base_path = os.path.abspath(os.path.dirname(__file__))
    full_path = os.path.join(base_path, subfolder, relative_path)
    log(f"检查路径: {full_path}, 存在: {os.path.exists(full_path)}", base_path)
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    full_path = os.path.join(base_path, relative_path)
    log(f"回退路径: {full_path}, 存在: {os.path.exists(full_path)}", base_path)
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    return full_path

def read_bank_config(base_path):
    """读取 config.txt 中的宁波银行配置"""
    config_path = get_resource_path("config.txt")
    log(f"尝试加载配置文件: {config_path}", base_path)
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
    return username, password, login_url, config_path

def find_and_click_image(template_path, offset_x=0, offset_y=0, threshold=0.5, max_attempts=10):
    """使用模板匹配找到图像并点击"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        if not os.path.exists(template_path):
            log(f"模板路径不存在: {template_path}")
            return None
        try:
            template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception as e:
            log(f"模板加载异常: {template_path}, 错误: {e}")
            return None
        if template is None:
            log(f"无法加载模板图像: {template_path}")
            return None
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            x, y = max_loc
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)
            log(f"第{attempt+1}次尝试成功，点击位置: ({x + offset_x}, {y + offset_y})")
            return (x, y)
        time.sleep(1)
    log(f"未找到模板: {template_path}，尝试次数: {max_attempts}")
    return None

def handle_overwrite_dialog():
    """处理文件覆盖确认弹窗"""
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
        log(f"未检测到覆盖确认窗口或点击失败: {e}")

def handle_save_dialog(save_path, pdf_filename):
    """处理另存为窗口，输入路径并保存"""
    full_path = os.path.join(save_path, pdf_filename)
    log(f"尝试捕捉‘另存为’窗口，目标路径: {full_path}")
    try:
        desktop = Desktop(backend="win32")
        dialogs = desktop.windows(title_re="^另存为$")
        if not dialogs:
            raise Exception("未找到标题为 '另存为' 的窗口")
        for i, dlg_wrapper in enumerate(dialogs):
            try:
                handle = dlg_wrapper.handle
                app = Application(backend="win32").connect(handle=handle)
                dlg = app.window(handle=handle)
                dlg.set_focus()
                time.sleep(0.3)
                edit = dlg.child_window(class_name="Edit")
                edit.set_focus()
                edit.set_edit_text(full_path)
                log(f"窗口{i + 1}设置路径成功: {full_path}")
                time.sleep(0.5)
                save_btn = dlg.child_window(class_name="Button", title_re="保存|Save")
                save_btn.click()
                log("点击保存按钮完成")
                time.sleep(3)
                handle_overwrite_dialog()
                return
            except Exception as inner_e:
                log(f"窗口{i + 1}处理失败: {inner_e}")
                continue
        raise Exception("未找到可用的‘另存为’窗口")
    except Exception as e:
        log(f"快速保存失败，错误: {e}")
        raise

def run_ningbo_bank(playwright: Playwright, base_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None):
    """执行宁波银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, base_path, log_callback)
    log_local("启动宁波银行导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(base_path)
        log_local(f"加载配置文件: {config_path}")
    except Exception as e:
        log_local(f" config.txt 文件加载失败: {str(e)}")
        raise
    browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
    if not os.path.exists(browser_path):
        log_local(f"Playwright 浏览器路径不存在: {browser_path}")
        raise FileNotFoundError(f"Playwright 浏览器路径不存在: {browser_path}")
    try:
        log_local("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(90000)
        log_local(f"访问登录页面: {login_url}")
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
            log_local(f"处理项目：{xiangmu}，银行账号：{account}")
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
                    log_local(f"使用银行账号查询：{account}")
                    page.wait_for_timeout(1000)
                    try:
                        page.get_by_role("listitem").filter(has_text=xiangmu).locator("span").nth(2).click()
                    except Exception as e:
                        log_local(f"点击搜索结果失败（项目：{xiangmu}，账号：{account}）：{str(e)}")
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
                    log_local(f"使用银行账号查询：{account}")
                    page.wait_for_timeout(1000)
                    try:
                        page.get_by_role("link", name=account).click()
                    except Exception as e:
                        log_local(f"点击链接失败（项目：{xiangmu}，账号：{account}）：{str(e)}")
                        page.screenshot(path=os.path.join(base_path, f"error_select_{xiangmu}.png"))
                        continue
                page.get_by_role("button", name=" 查询").click()
                page.wait_for_selector('role=checkbox[name="Toggle Selection of All Rows"]', timeout=10000)
                checkbox = page.get_by_role("checkbox", name="Toggle Selection of All Rows")
                if not (checkbox.is_visible() and checkbox.is_enabled()):
                    log_local(f"无回单或流水数据，跳过导出（项目：{xiangmu}，账号：{account}）")
                    page.screenshot(path=os.path.join(base_path, f"error_no_data_{xiangmu}.png"))
                    continue
                try:
                    if not checkbox.is_checked():
                        checkbox.check()
                    log_local("复选框已选中")
                except Exception as e:
                    log_local(f"选中复选框失败（项目：{xiangmu}，账号：{account}）：{str(e)}")
                    page.screenshot(path=os.path.join(base_path, f"error_checkbox_{xiangmu}.png"))
                    continue
                # 导出回单
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
                            log_local(f"银行回单导出完成：{filename}")
                            found = True
                            break
                    if not found:
                        log_local(f"未找到‘凭证导出’菜单项（项目：{xiangmu}）")
                        page.screenshot(path=os.path.join(base_path, f"error_menu_{xiangmu}.png"))
                except Exception as e:
                    log_local(f"导出银行回单失败：{str(e)}")
                # 导出流水
                page.get_by_role("button", name="导出").click()
                with page.expect_download() as download_info:
                    page.get_by_text("对账单导出", exact=True).click()
                download = download_info.value
                filename = f"{xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
                download.save_as(os.path.join(duizhang_path, filename))
                log_local(f"银行流水导出完成：{filename}")
                # 打印对账单为PDF
                try:
                    page.get_by_role("button", name="打印 ").click()
                    page.wait_for_timeout(2000)
                    menus = page.locator("css=[id^='dropdown-menu-']")
                    count = menus.count()
                    found = False
                    for i in range(count):
                        item = menus.nth(i).filter(has_text="对账单打印")
                        if item.is_visible():
                            item.click()
                            log_local(f"点击了对账单打印菜单项")
                            found = True
                            break
                    if not found:
                        log_local(f"未找到‘对账单打印’菜单项（项目：{xiangmu}）")
                        page.screenshot(path=os.path.join(base_path, f"error_print_menu_{xiangmu}.png"))
                        continue
                    log_local("等待 Chrome 打印窗口...")
                    time.sleep(5)
                    target_printer_path = get_resource_path("target_printer.bmp")
                    save_as_pdf_default_path = get_resource_path("save_as_pdf_default.bmp")
                    save_as_pdf_hover_path = get_resource_path("save_as_pdf_hover.bmp")
                    save_button_path = get_resource_path("save_button.bmp")
                    if not (os.path.exists(target_printer_path) and os.path.exists(save_as_pdf_default_path) and
                            os.path.exists(save_as_pdf_hover_path) and os.path.exists(save_button_path)):
                        log_local(f"模板图像缺失或大小为0")
                        raise FileNotFoundError("请准备相关模板图像并放入 seek 文件夹")
                    log_local("定位‘目标打印机’位置...")
                    target_pos = find_and_click_image(target_printer_path)
                    if not target_pos:
                        log_local("未找到‘目标打印机’文字")
                        pyautogui.screenshot(os.path.join(base_path, f"error_target_printer_{xiangmu}.png"))
                        continue
                    x_target, y_target = target_pos
                    x_offset = 250
                    pyautogui.click(x_target + x_offset, y_target)
                    log_local(f"模拟点击偏移位置: ({x_target + x_offset}, {y_target})")
                    time.sleep(2)
                    pdf_clicked = False
                    for attempt in range(3):
                        pyautogui.moveTo(x_target + x_offset, y_target + 20)
                        time.sleep(0.5)
                        if find_and_click_image(save_as_pdf_default_path):
                            log_local("成功点击默认状态的‘另存为 PDF’按钮")
                            pdf_clicked = True
                            time.sleep(2)
                            break
                        elif find_and_click_image(save_as_pdf_hover_path):
                            log_local("成功点击悬停状态的‘另存为 PDF’按钮")
                            pdf_clicked = True
                            time.sleep(2)
                            break
                    if not pdf_clicked:
                        log_local("未找到‘另存为 PDF’按钮")
                        pyautogui.screenshot(os.path.join(base_path, f"error_save_as_pdf_{xiangmu}.png"))
                        continue
                    if find_and_click_image(save_button_path):
                        log_local("成功点击‘保存’按钮")
                        time.sleep(2)
                    else:
                        log_local("未找到‘保存’按钮")
                        pyautogui.screenshot(os.path.join(base_path, f"error_save_button_{xiangmu}.png"))
                        continue
                    pdf_filename = f"{xiangmu}_对账单打印_{kaishiriqi}_{jieshuriqi}.pdf"
                    handle_save_dialog(duizhangdan_path, pdf_filename)
                    pdf_path = os.path.join(duizhangdan_path, pdf_filename)
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                        log_local(f"对账单打印 PDF 完成：{pdf_path}")
                    else:
                        log_local(f"PDF 文件生成失败或为空：{pdf_path}")
                except Exception as e:
                    log_local(f"打印对账单 PDF 失败（项目：{xiangmu}）：{str(e)}")
                    page.screenshot(path=os.path.join(base_path, f"error_print_{xiangmu}.png"))
                previous_xiangmu = xiangmu
            except Exception as e:
                log_local(f"导出失败（项目：{xiangmu}）：{str(e)}")
                continue
        context.close()
        browser.close()
    except Exception as e:
        log_local(f"初始化失败：{str(e)}")
        raise
    finally:
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()

def main(page: Page):
    """Flet桌面应用主函数"""
    page.title = "银行流水回单导出"
    page.window_width = 800
    page.window_height = 600
    page.window_resizable = True

    # UI状态变量
    excel_data = []
    base_path = r"D:\Desktop"
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

    def select_base_path(e: FilePickerResultEvent):
        if e.path:
            nonlocal base_path
            base_path = e.path
            base_path_text.value = f"下载路径: {base_path}"
            update_log(f"下载路径设置为：{base_path}")
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
        if not (kaishi and jieshu and base_path):
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
                    run_ningbo_bank(playwright, base_path, excel_data, kaishi, jieshu, log_callback=update_log)
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

    run_button.on_click = run_export  # 确保绑定

    # 文件选择器
    file_picker = FilePicker(on_result=import_excel)
    dir_picker = FilePicker(on_result=select_base_path)
    page.overlay.extend([file_picker, dir_picker])

    # UI布局
    base_path_text = ft.Text(f"下载路径: {base_path}")
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
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}")
    try:
        ft.app(target=main)
    except Exception as e:
        log(f"应用启动失败: {str(e)}")
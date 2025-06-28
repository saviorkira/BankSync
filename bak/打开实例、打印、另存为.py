import os
import sys
import time
import pyautogui
import cv2
from pywinauto import Desktop
from pywinauto.application import Application
import numpy as np
from datetime import datetime
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from playwright.sync_api import sync_playwright, Playwright
import configparser

# 全局日志函数
def log(msg, base_path=r"D:\Desktop"):
    print(msg)
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")

def force_check_expiration_local(expire_date_str="2026-06-01"):
    """使用本地系统时间判断是否过期，过期则退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期为 {expire_date_str})。")
        sys.exit(0)

def get_resource_path(relative_path, subfolder="seek"):
    """获取资源路径，优先检查 seek 子目录"""
    base_path = os.path.abspath(os.path.dirname(__file__))
    full_path = os.path.join(base_path, subfolder, relative_path)
    log(f"检查路径: {full_path}, 存在: {os.path.exists(full_path)}, 大小: {os.path.getsize(full_path) if os.path.exists(full_path) else 0} bytes")
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    full_path = os.path.join(base_path, relative_path)
    log(f"回退路径: {full_path}, 存在: {os.path.exists(full_path)}, 大小: {os.path.getsize(full_path) if os.path.exists(full_path) else 0} bytes")
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    return full_path

def read_bank_config():
    """读取 config.txt 中的宁波银行配置"""
    config_path = get_resource_path("../config.txt")
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
    """使用模板匹配找到图像并点击，添加偏移，降低阈值以容忍差异"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        # 使用 imdecode 读取模板图像（支持中文路径）
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
            with open(template_path, 'rb') as f:
                log(f"文件内容前10字节: {f.read(10).hex()}")
            return None

        log(f"模板尺寸: {template.shape}")

        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        log(f"匹配度: {max_val}, 位置: {max_loc}")

        if max_val >= threshold:
            x, y = max_loc
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)
            log(f"第{attempt+1}次尝试成功，点击位置: ({x + offset_x}, {y + offset_y})")
            return (x, y)

        time.sleep(1)

    log(f"未找到模板: {template_path}，尝试次数: {max_attempts}")
    return None

def handle_overwrite_dialog():
    """处理文件已存在的覆盖确认弹窗，自动点击‘是’或‘替换’按钮"""
    from pywinauto import Desktop
    try:
        app = Desktop(backend="win32")
        dialog = app.window(title_re=".*文件已存在.*|.*确认保存.*|.*确认另存为.*|.*Confirm Save.*|.*Replace.*")
        dialog.wait("exists ready", timeout=5)
        dialog.set_focus()
        print(f"覆盖确认窗口标题: {dialog.window_text()}")  # 调试输出

        replace_btn = dialog.child_window(title_re="是.*|替换.*|Yes.*|Replace.*", class_name="Button")
        replace_btn.wait("exists enabled visible ready", timeout=3)
        print(f"找到覆盖按钮，标题: {replace_btn.element_info.name}")  # 调试输出
        replace_btn.click()
        print("点击覆盖按钮成功")
        time.sleep(1)
    except Exception as e:
        print(f"未检测到覆盖确认窗口或点击失败: {e}")

def handle_save_dialog(huidan_path, pdf_filename):
    """处理另存为窗口，输入完整路径并保存"""
    import os, time
    full_path = os.path.join(huidan_path, pdf_filename)
    log("尝试捕捉‘另存为’窗口...")

    try:
        desktop = Desktop(backend="win32")
        dialogs = desktop.windows(title_re="^另存为$")

        if not dialogs:
            raise Exception("未找到标题为 '另存为' 的窗口")

        for i, dlg_wrapper in enumerate(dialogs):
            try:
                log(f"尝试处理第{i + 1}个窗口: {dlg_wrapper.window_text()}")
                handle = dlg_wrapper.handle

                # 使用 handle 获取 WindowSpecification 对象
                app = Application(backend="win32").connect(handle=handle)
                dlg = app.window(handle=handle)

                dlg.set_focus()
                time.sleep(0.3)

                edit = dlg.child_window(class_name="Edit")
                edit.set_focus()
                edit.set_edit_text(full_path)
                log(f"第{i + 1}个窗口设置路径成功: {full_path}")

                time.sleep(0.5)
                save_btn = dlg.child_window(class_name="Button", title_re="保存|Save")
                save_btn.click()
                log("点击保存按钮完成")
                time.sleep(3)

                handle_overwrite_dialog()
                return  # 成功后退出
            except Exception as inner_e:
                log(f"窗口{i + 1}处理失败: {inner_e}")
                continue

        raise Exception("未找到可用的‘另存为’窗口，或操作失败")

    except Exception as e:
        log(f"快速保存失败，错误: {e}")
        raise

def run_print_test(playwright: Playwright):
    """测试对账单打印功能，生成 PDF"""
    xiangmu = "重信·北极22026·华睿精选9号集合资金信托计划"
    account = "77970122000179566"
    kaishiriqi = "2025-03-21"
    jieshuriqi = "2025-03-21"
    base_path = r"D:\Desktop"

    log(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
    if not os.path.exists(browser_path):
        log(f"Playwright 浏览器路径不存在: {browser_path}")
        log(f"请将 'playwright-browsers' 文件夹放置在 .exe 所在目录: {os.path.dirname(sys.executable)}")
        log("或运行 'playwright install' 确保浏览器可用")
        raise FileNotFoundError(f"Playwright 浏览器路径不存在: {browser_path}")

    username, password, login_url, config_path = read_bank_config()
    log(f"加载配置文件: {config_path}")

    try:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(90000)  # 增加超时时间至 90 秒
        page.goto(login_url)
        page.get_by_role("textbox", name="用户名").fill(username)
        page.get_by_role("textbox", name="请输入您的密码").fill(password)
        page.wait_for_selector('text=账户管理', timeout=90000)
        page.get_by_role("link", name="账户管理").click()
        page.get_by_role("link", name="账户明细").click()

        log(f"处理项目：{xiangmu}，银行账号：{account}")
        huidan_path = os.path.join(base_path, xiangmu, "银行回单")
        os.makedirs(huidan_path, exist_ok=True)

        # 查询
        page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
        page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").fill(account)
        log(f"使用银行账号查询：{account}")
        page.wait_for_timeout(1000)
        try:
            page.get_by_role("listitem").filter(has_text=xiangmu).locator("span").nth(2).click()
        except Exception as e:
            log(f"点击搜索结果失败（项目：{xiangmu}，账号：{account}）：{str(e)}")
            page.screenshot(path=f"error_select_{xiangmu}.png")
            log(f"已保存选择错误截图：error_select_{xiangmu}.png")
            raise

        page.get_by_text("展开").first.click()
        page.get_by_role("textbox", name="开始日期").fill(kaishiriqi)
        page.get_by_role("textbox", name="开始日期").press("Enter")
        page.get_by_role("textbox", name="结束日期").fill(jieshuriqi)
        page.get_by_role("textbox", name="结束日期").press("Enter")
        page.get_by_role("button", name=" 查询").click()

        # 等待查询结果加载
        page.wait_for_selector('role=checkbox[name="Toggle Selection of All Rows"]', timeout=10000)
        page.wait_for_timeout(3000)

        # 确认并选中复选框
        checkbox = page.get_by_role("checkbox", name="Toggle Selection of All Rows")
        if not (checkbox.is_visible() and checkbox.is_enabled()):
            log(f"无查询结果或复选框不可用（项目：{xiangmu}，账号：{account}）")
            page.screenshot(path=f"error_no_data_{xiangmu}.png")
            log(f"已保存无数据截图：error_no_data_{xiangmu}.png")
            raise Exception("无查询结果或复选框不可用")

        try:
            if not checkbox.is_checked():
                checkbox.check()
                log("复选框已选中")
            else:
                log("复选框已默认选中")
        except Exception as e:
            log(f"选中复选框失败（项目：{xiangmu}，账号：{account}）：{str(e)}")
            page.screenshot(path=f"error_checkbox_{xiangmu}.png")
            log(f"已保存复选框错误截图：error_checkbox_{xiangmu}.png")
            raise

        # 打印对账单
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
                    log(f"点击了对账单打印菜单项")
                    found = True
                    break
            if not found:
                log(f"未找到‘对账单打印’菜单项（项目：{xiangmu}）")
                page.screenshot(path=f"error_print_menu_{xiangmu}.png")
                log(f"已保存打印菜单错误截图：error_print_menu_{xiangmu}.png")
                raise Exception("未找到对账单打印菜单项")

            # 等待 Chrome 打印窗口
            log("等待 Chrome 打印窗口...")
            time.sleep(5)

            # 加载模板文件
            target_printer_path = get_resource_path("target_printer.bmp")
            printer_dropdown_path = get_resource_path("printer_dropdown.bmp")
            save_as_pdf_default_path = get_resource_path("save_as_pdf_default.bmp")
            save_as_pdf_hover_path = get_resource_path("save_as_pdf_hover.bmp")
            save_button_path = get_resource_path("save_button.bmp")

            if not (os.path.exists(target_printer_path) and os.path.exists(printer_dropdown_path) and
                    os.path.exists(save_as_pdf_default_path) and os.path.exists(save_as_pdf_hover_path) and
                    os.path.exists(save_button_path)):
                log(f"模板图像缺失或大小为0: target_printer.bmp, printer_dropdown.bmp, save_as_pdf_default.bmp, save_as_pdf_hover.bmp 或 save_button.bmp")
                raise FileNotFoundError("请准备相关模板图像并放入 seek 文件夹，确保文件大小大于0")

            # 定位“目标打印机”文字
            log("定位‘目标打印机’位置...")
            target_pos = find_and_click_image(target_printer_path)
            if target_pos:
                x_target, y_target = target_pos
                log(f"找到‘目标打印机’文字，位置: ({x_target}, {y_target})")
            else:
                log("未找到‘目标打印机’文字")
                pyautogui.screenshot(f"error_target_printer_{xiangmu}.png")
                raise Exception("未找到打印窗口中的‘目标打印机’文字")

            # 点击“目标打印机”右侧偏移 250 像素模拟下拉按钮
            log("点击‘目标打印机’右侧偏移 250 像素模拟下拉按钮...")
            x_offset = 250
            pyautogui.click(x_target + x_offset, y_target)
            log(f"模拟点击偏移位置: ({x_target + x_offset}, {y_target})")
            time.sleep(2)

            # 尝试匹配并点击“另存为 PDF”
            log("尝试匹配并点击‘另存为 PDF’选项...")
            pdf_clicked = False
            for attempt in range(3):
                pyautogui.moveTo(x_target + x_offset, y_target + 20)
                time.sleep(0.5)
                if find_and_click_image(save_as_pdf_default_path):
                    log("成功点击默认状态的‘另存为 PDF’按钮")
                    pdf_clicked = True
                    time.sleep(2)
                    break
                elif find_and_click_image(save_as_pdf_hover_path):
                    log("成功点击悬停状态的‘另存为 PDF’按钮")
                    pdf_clicked = True
                    time.sleep(2)
                    break
                log(f"第{attempt+1}次尝试未找到‘另存为 PDF’，重试...")
                time.sleep(1)

            if not pdf_clicked:
                log("未找到‘另存为 PDF’按钮（默认或悬停状态），请检查截图")
                pyautogui.screenshot(f"error_save_as_pdf_{xiangmu}.png")
                raise Exception("未找到打印窗口中的‘另存为 PDF’按钮")

            # 点击“保存”按钮
            log("尝试点击‘保存’按钮...")
            if find_and_click_image(save_button_path):
                log("成功点击‘保存’按钮")
                time.sleep(2)
            else:
                log("未找到‘保存’按钮，请检查截图")
                pyautogui.screenshot(f"error_save_button_{xiangmu}.png")
                raise Exception("未找到打印窗口中的‘保存’按钮")

            # 处理另存为窗口，写入文件名并保存
            pdf_filename = f"{xiangmu}_对账单打印_{kaishiriqi}_{jieshuriqi}.pdf"
            handle_save_dialog(huidan_path, pdf_filename)

            pdf_path = os.path.join(huidan_path, pdf_filename)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                log(f"对账单打印 PDF 完成：{pdf_filename}")
            else:
                log(f"PDF 文件生成失败或为空：{pdf_path}")
                raise Exception("PDF 文件生成失败或为空")
        except Exception as e:
            log(f"打印对账单 PDF 失败（项目：{xiangmu}）：{str(e)}")
            if 'page' in locals():
                page.screenshot(path=f"error_print_{xiangmu}.png")
            raise
        finally:
            if 'context' in locals():
                context.close()
            if 'browser' in locals():
                browser.close()
    except Exception as e:
        log(f"测试失败：{str(e)}")
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()
        raise

if __name__ == "__main__":
    force_check_expiration_local("2026-06-01")
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}")
    with sync_playwright() as playwright:
        run_print_test(playwright)
import os
import sys
import configparser
import time
import pyautogui
import cv2
import numpy as np
from PIL import Image
from playwright.sync_api import sync_playwright, Playwright
from datetime import datetime

# 全局日志函数
def log(msg, base_path=r"D:\Desktop"):
    print(msg)
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")

def force_check_expiration_local(expire_date_str="2025-06-30"):
    """使用本地系统时间判断是否过期，过期则退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期为 {expire_date_str})。")
        sys.exit(0)

def get_resource_path(relative_path, subfolder=""):
    """获取资源路径，支持子文件夹，优先检查 .exe 目录，再检查 _MEIPASS 或脚本目录"""
    base_path = ""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    elif hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))

    if subfolder:
        full_path = os.path.join(base_path, subfolder, relative_path)
        log(f"检查路径: {full_path}")
        if os.path.exists(full_path):
            return full_path
    full_path = os.path.join(base_path, relative_path)
    log(f"回退路径: {full_path}")
    if os.path.exists(full_path):
        return full_path
    return full_path  # 若不存在，返回默认路径，便于调试

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

def find_and_click_image(template_path, offset_x=0, offset_y=0, threshold=0.8, max_attempts=10):
    """使用模板匹配找到图像并点击，添加偏移"""
    for _ in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        template = cv2.imread(template_path)
        if template is None:
            raise ValueError(f"无法加载模板图像: {template_path}")
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        loc = cv2.minMaxLoc(result)[3]
        if loc[2] >= threshold:
            x, y = loc[2]
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)  # 偏移点击
            return (x, y)  # 返回匹配位置
        time.sleep(1)
    return None

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
            time.sleep(5)  # 等待打印窗口出现
            # target_printer_path = get_resource_path("target_printer.png", "seek")
            # printer_dropdown_path = get_resource_path("printer_dropdown.png", "seek")
            # save_as_pdf_path = get_resource_path("save_as_pdf.png", "seek")
            # if not (os.path.exists(target_printer_path) and os.path.exists(printer_dropdown_path) and os.path.exists(save_as_pdf_path)):
            #     log(f"模板图像缺失: target_printer.png, printer_dropdown.png 或 save_as_pdf.png 在 seek 文件夹中")
            #     raise FileNotFoundError("请准备 target_printer.png, printer_dropdown.png 和 save_as_pdf.png 模板图像并放入 seek 文件夹")

            # # 识图定位“目标打印机”文字
            # target_pos = find_and_click_image(target_printer_path)
            # if target_pos:
            #     x_target, y_target = target_pos
            #     log(f"找到‘目标打印机’文字，位置: ({x_target}, {y_target})")
            # else:
            #     log("未找到‘目标打印机’文字")
            #     pyautogui.screenshot(f"error_target_printer_{xiangmu}.png")
            #     raise Exception("未找到打印窗口中的‘目标打印机’文字")
            #
            # # 识图并点击“目标打印机”右侧的下拉箭头
            # dropdown_offset_x = 80  # 调整此值以匹配“目标打印机”右侧箭头位置，单位像素
            # if find_and_click_image(printer_dropdown_path, offset_x=dropdown_offset_x, offset_y=0):
            #     log(f"成功点击‘目标打印机’右侧下拉菜单按钮")
            #     time.sleep(2)  # 等待下拉菜单出现
            # else:
            #     log("未找到‘目标打印机’右侧下拉菜单按钮")
            #     pyautogui.screenshot(f"error_printer_dropdown_{xiangmu}.png")
            #     raise Exception("未找到打印窗口中的‘目标打印机’右侧下拉菜单按钮")

            # 暂停，保持打印窗口打开供调试
            log("打印窗口已打开，暂停脚本。请使用 测试Chrome打印窗口及弹出的另存为.py 调试。")
            log("输入 'q' 退出并关闭浏览器，或按 Enter 保持打开继续手动操作...")
            while True:
                user_input = input("输入您的选择: ")
                if user_input.lower() == 'q':
                    break
                log("保持打印窗口打开，随时使用 测试Chrome打印窗口及弹出的另存为.py 调试...")

            # 识图并点击“另存为 PDF”（若需选择）
            if find_and_click_image(save_as_pdf_path):
                log("成功点击‘另存为 PDF’按钮")
                time.sleep(2)  # 等待保存对话框
                pyautogui.write(huidan_path)  # 输入保存路径
                pyautogui.press("enter")  # 确认保存
                log(f"保存 PDF 到: {huidan_path}")
                time.sleep(5)  # 等待保存完成
            else:
                # 若默认已为“另存为 PDF”，直接保存
                log("未找到‘另存为 PDF’按钮，检查默认选项")
                pyautogui.press("enter")  # 尝试直接保存
                time.sleep(5)
                log(f"尝试以默认选项保存 PDF 到: {huidan_path}")

            pdf_filename = f"{xiangmu}_对账单打印_{kaishiriqi}_{jieshuriqi}.pdf"
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
    print("设置 PLAYWRIGHT_BROWSERS_PATH:", os.environ["PLAYWRIGHT_BROWSERS_PATH"])
    with sync_playwright() as playwright:
        run_print_test(playwright)
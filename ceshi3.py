from playwright.sync_api import Playwright
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_overwrite_dialog, handle_save_dialog
import os
import time
import pyautogui

def run_ningbo_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None):
    """执行宁波银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, project_root, log_callback)  # 日志保存到 download_path
    log_local("启动宁波银行导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(project_root, "ningbo_bank")
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
        page.set_default_timeout(120000)
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
                duizhang_path = os.path.join(download_path, xiangmu, "银行流水")
                huidan_path = os.path.join(download_path, xiangmu, "银行回单")
                duizhangdan_path = os.path.join(download_path, xiangmu, "银行对账单")
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
                        page.screenshot(path=os.path.join(download_path, f"error_select_{xiangmu}.png"))
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
                        page.screenshot(path=os.path.join(download_path, f"error_select_{xiangmu}.png"))
                        continue
                page.get_by_role("button", name=" 查询").click()
                page.wait_for_selector('role=checkbox[name="Toggle Selection of All Rows"]', timeout=10000)
                checkbox = page.get_by_role("checkbox", name="Toggle Selection of All Rows")
                if not (checkbox.is_visible() and checkbox.is_enabled()):
                    log_local(f"无回单或流水数据，跳过导出（项目：{xiangmu}，账号：{account}）")
                    page.screenshot(path=os.path.join(download_path, f"error_no_data_{xiangmu}.png"))
                    continue
                try:
                    if not checkbox.is_checked():
                        checkbox.check()
                    log_local("复选框已选中")
                except Exception as e:
                    log_local(f"选中复选框失败（项目：{xiangmu}，账号：{account}）：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_checkbox_{xiangmu}.png"))
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
                        page.screenshot(path=os.path.join(download_path, f"error_menu_{xiangmu}.png"))
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
                    log_local("点击打印按钮，等待对账单打印按钮...")
                    time.sleep(2)
                    duizhangdan_button_path = get_resource_path("ningbo_duizhangdandayin.bmp", project_root)
                    if not os.path.exists(duizhangdan_button_path):
                        log_local(f"模板图像不存在: {duizhangdan_button_path}")
                        raise FileNotFoundError(f"模板图像不存在: {duizhangdan_button_path}")
                    log_local("定位‘对账单打印’按钮...")
                    if find_and_click_image(duizhangdan_button_path, download_path, max_attempts=15):
                        log_local("成功点击‘对账单打印’按钮")
                    else:
                        log_local("未找到‘对账单打印’按钮")
                        pyautogui.screenshot(os.path.join(download_path, f"error_duizhangdan_button_{xiangmu}.png"))
                        continue
                    log_local("等待 Chrome 打印窗口...")
                    time.sleep(2)
                    target_printer_path = get_resource_path("target_printer.bmp", project_root)
                    save_as_pdf_default_path = get_resource_path("save_as_pdf_default.bmp", project_root)
                    save_as_pdf_hover_path = get_resource_path("save_as_pdf_hover.bmp", project_root)
                    save_button_path = get_resource_path("save_button.bmp", project_root)
                    if not (os.path.exists(target_printer_path) and os.path.exists(save_as_pdf_default_path) and
                            os.path.exists(save_as_pdf_hover_path) and os.path.exists(save_button_path)):
                        log_local(f"模板图像缺失或大小为0")
                        raise FileNotFoundError("请准备相关模板图像并放入 seek 文件夹")
                    log_local("定位‘目标打印机’位置...")
                    target_pos = find_and_click_image(target_printer_path, download_path)
                    if not target_pos:
                        log_local("未找到‘目标打印机’文字")
                        pyautogui.screenshot(os.path.join(download_path, f"error_target_printer_{xiangmu}.png"))
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
                        if find_and_click_image(save_as_pdf_default_path, download_path):
                            log_local("成功点击默认状态的‘另存为 PDF’按钮")
                            pdf_clicked = True
                            time.sleep(1)
                            break
                        elif find_and_click_image(save_as_pdf_hover_path, download_path):
                            log_local("成功点击悬停状态的‘另存为 PDF’按钮")
                            pdf_clicked = True
                            time.sleep(1)
                            break
                    if not pdf_clicked:
                        log_local("未找到‘另存为 PDF’按钮")
                        pyautogui.screenshot(os.path.join(download_path, f"error_save_as_pdf_{xiangmu}.png"))
                        continue
                    if find_and_click_image(save_button_path, download_path):
                        log_local("成功点击‘保存’按钮")
                        time.sleep(1)
                    else:
                        log_local("未找到‘保存’按钮")
                        pyautogui.screenshot(os.path.join(download_path, f"error_save_button_{xiangmu}.png"))
                        continue
                    pdf_filename = f"{xiangmu}_对账单打印_{kaishiriqi}_{jieshuriqi}.pdf"
                    handle_save_dialog(duizhangdan_path, pdf_filename, download_path)
                    pdf_path = os.path.join(duizhangdan_path, pdf_filename)
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                        log_local(f"对账单打印 PDF 完成：{pdf_path}")
                    else:
                        log_local(f"PDF 文件生成失败或为空：{pdf_path}")
                except Exception as e:
                    log_local(f"打印对账单 PDF 失败（项目：{xiangmu}）：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_print_{xiangmu}.png"))
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
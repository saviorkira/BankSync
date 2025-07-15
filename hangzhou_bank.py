from playwright.sync_api import Playwright
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_save_dialog
import os
import time
import pyautogui

def run_hangzhou_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None):
    """执行杭州银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, project_root, log_callback)
    log_local("启动杭州银行导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(project_root, "hangzhou_bank")
        log_local(f"加载配置文件: {config_path}")
    except Exception as e:
        log_local(f"config.txt 文件加载失败: {str(e)}")
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
        page.get_by_role("textbox", name="请输入客户号").fill(username)
        page.get_by_role("textbox", name="请输入操作员号").fill("2001")
        # page.get_by_role("textbox", name="请输入您的密码").fill(password)
        # page.get_by_role("button", name="登录").click()
        log_local("等待账户管理页面加载...")
        page.wait_for_selector('text=账户', timeout=30000)
        page.get_by_text("账户", exact=True).click()
        page.get_by_role("menuitem", name="流水查询").click()
        page.get_by_role("combobox", name="开始日期").fill(kaishiriqi)
        page.get_by_role("combobox", name="结束日期").fill(jieshuriqi)
        # page.get_by_role("combobox", name="开始日期").press("Enter")
        # page.get_by_role("combobox", name="结束日期").press("Enter")

        for index, (xiangmu, account) in enumerate(projects_accounts):
            log_local(f"处理项目：{xiangmu}，银行账号：{account}")
            try:
                # 创建文件夹
                duizhang_path = os.path.join(download_path, xiangmu, "银行流水")
                huidan_path = os.path.join(download_path, xiangmu, "银行回单")
                duizhangdan_path = os.path.join(download_path, xiangmu, "银行对账单")
                os.makedirs(duizhang_path, exist_ok=True)
                os.makedirs(huidan_path, exist_ok=True)
                os.makedirs(duizhangdan_path, exist_ok=True)

                # 查询账号
                page.get_by_role("textbox", name="账号/户名").click()
                time.sleep(2)
                page.get_by_role("textbox", name="账号/户名").fill(account)
                page.get_by_role("option", name="-重庆国际信托股份有限公司").click()
                log_local(f"使用银行账号查询：{account}")
                page.get_by_role("button", name="查询").click()
                time.sleep(3)
                # page.wait_for_selector('role=row', state="visible", timeout=10000)  # 等待数据加载
                # # 等待复选框加载
                # checkbox = page.get_by_role("row",
                #                             name="交易时间 交易流水号 收入金额 支出金额 余额 对方 对方开户行 用途 操作").locator(
                #     "span").nth(1)
                # try:
                #     checkbox.wait_for(state="visible", timeout=10000)
                # except Exception:
                #     log_local(f"数据加载超时或无数据（项目：{xiangmu}，账号：{account}）")
                #     page.screenshot(path=os.path.join(download_path, f"error_data_load_{xiangmu}.png"))
                #     continue
                # 等待复选框加载
                checkbox = page.get_by_role("row", name="交易时间 交易流水号 收入金额 支出金额 余额 对方 对方开户行 用途 操作").locator("span").nth(1)
                try:
                    checkbox.wait_for(state="visible", timeout=10000)
                except Exception:
                    log_local(f"数据加载超时或无数据（项目：{xiangmu}，账号：{account}）")
                    page.screenshot(path=os.path.join(download_path, f"error_data_load_{xiangmu}.png"))
                    continue

                # 检查复选框是否可交互
                if not checkbox.is_enabled():
                    log_local(f"无回单或流水数据，跳过导出（项目：{xiangmu}，账号：{account}）")
                    page.screenshot(path=os.path.join(download_path, f"error_no_data_{xiangmu}.png"))
                    continue
                # # 检查复选框状态
                # checkbox = page.get_by_role("row", name="交易时间 交易流水号 收入金额 支出金额 余额 对方 对方开户行 用途 操作").locator("span").nth(1)
                # if not (checkbox.is_visible() and checkbox.is_enabled() and checkbox.is_checked()):
                #     log_local(f"无回单或流水数据，跳过导出（项目：{xiangmu}，账号：{account}）")
                #     page.screenshot(path=os.path.join(download_path, f"error_no_data_{xiangmu}.png"))
                #     continue

                # 导出流水
                try:
                    with page.expect_download() as liushui_download_info:
                        page.get_by_role("button", name="导出Excel").click()
                    download = liushui_download_info.value
                    filename = f"{xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
                    download.save_as(os.path.join(duizhang_path, filename))
                    log_local(f"银行流水导出完成：{filename}")
                except Exception as e:
                    log_local(f"导出银行流水失败（项目：{xiangmu}）：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_export_excel_{xiangmu}.png"))
                    continue  # 失败后跳到下一个项目

                # 导出对账单
                try:
                    with page.expect_download() as duizhangdan_download_info:
                        page.get_by_role("button", name="流水打印", exact=True).click()
                    download = duizhangdan_download_info.value
                    filename = f"{xiangmu}_银行对账单_{kaishiriqi}_{jieshuriqi}.pdf"
                    download.save_as(os.path.join(duizhangdan_path, filename))
                    log_local(f"银行对账单导出完成：{filename}")
                except Exception as e:
                    log_local(f"导出银行对账单失败（项目：{xiangmu}）：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_export_duizhangdan_{xiangmu}.png"))
                    continue  # 失败后跳到下一个项目

                # 导出回单
                try:
                    page.locator("#app").get_by_text("条/页").click()
                    page.wait_for_selector('role=option[name="100条/页"]', state="visible", timeout=5000)
                    page.get_by_role("option", name="100条/页").click()
                    checkbox.click()  # 重新选中复选框
                    log_local("复选框已重新选中")
                    with page.expect_download() as huidan_download_info:
                        page.get_by_role("button", name="回单打印", exact=True).click()
                    download = huidan_download_info.value
                    filename = f"{xiangmu}_银行回单_{kaishiriqi}_{jieshuriqi}.pdf"
                    download.save_as(os.path.join(huidan_path, filename))
                    log_local(f"银行回单导出完成：{filename}")
                except Exception as e:
                    log_local(f"导出银行回单失败（项目：{xiangmu}）：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_export_huidan_{xiangmu}.png"))
                    continue  # 失败后跳到下一个项目

            except Exception as e:
                log_local(f"处理项目失败（项目：{xiangmu}）：{str(e)}")
                page.screenshot(path=os.path.join(download_path, f"error_project_{xiangmu}.png"))
                continue  # 任意步骤失败跳到下一个项目

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
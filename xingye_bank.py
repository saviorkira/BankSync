from playwright.sync_api import Playwright
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_save_dialog, split_date_ranges
import os
import time
import pyautogui
import re

def run_xingye_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None, download_only_liushui=False):
    """执行杭州银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, project_root, log_callback)
    log_local("启动杭州银行导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(project_root, "xingye_bank")
        log_local(f"加载配置文件: {config_path}")
    except Exception as e:
        log_local(f"config.txt 文件加载失败: {str(e)}")
        raise
    browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
    if not os.path.exists(browser_path):
        log_local(f"Playwright 浏览器路径不存在: {browser_path}")
        raise FileNotFoundError(f"Playwright 浏览器路径不存在: {browser_path}")

    # 拆分日期范围
    try:
        date_ranges = split_date_ranges(kaishiriqi, jieshuriqi, max_months=3)
        log_local(f"日期范围已拆分为 {len(date_ranges)} 个子范围: {date_ranges}")
    except Exception as e:
        log_local(f"日期拆分失败: {str(e)}")
        raise

    try:
        log_local("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(120000)
        log_local(f"访问登录页面: {login_url}")
        page.goto(login_url)
        log_local("输入用户名和密码...")
        page.get_by_role("textbox", name="用户名/邮箱/手机号").fill(username)
        page.get_by_role("textbox", name="密码").fill(password)
        # page.get_by_role("textbox", name="请输入您的密码").fill(password)
        log_local("等待账户管理页面加载...")
        page.wait_for_selector('text=账务查询', timeout=30000)


        # 外层循环：遍历日期范围
        for range_index, (start_date, end_date) in enumerate(date_ranges):
            log_local(f"处理日期范围: {start_date} 至 {end_date}")



            for index, (xiangmuid, xiangmu, account) in enumerate(projects_accounts):
                log_local(f"处理产品：{xiangmuid}_{xiangmu}，托管账户：{account}")
                try:
                    # 创建文件夹
                    folder_name = f"{xiangmuid}_{xiangmu}"
                    liushui_path = os.path.join(download_path, folder_name, "银行流水")
                    huidan_path = os.path.join(download_path, folder_name, "银行回单")
                    duizhangdan_path = os.path.join(download_path, folder_name, "银行对账单")
                    os.makedirs(liushui_path, exist_ok=True)
                    if not download_only_liushui:
                        os.makedirs(huidan_path, exist_ok=True)
                        os.makedirs(duizhangdan_path, exist_ok=True)

                    # 查询账号
                    page.locator("a").filter(has_text="账务查询").click()
                    page.locator("a").filter(has_text=re.compile(r"^流水查询$")).click()

                    page.get_by_role("textbox", name="请点击选择按钮 选择产品").click()
                    page.get_by_role("button", name=" 选择").click()
                    page.get_by_role("textbox", name="产品名称").fill(xiangmu)  # 项目
                    page.get_by_role("button", name="全选").click()
                    page.get_by_role("button", name="确定").click()
                    page.get_by_role("textbox", name="交易日期").fill("2025-09-29 - 2025-09-29")  # 日期
                    start_date, end_date


                    # 导出流水
                    try:
                        with page.expect_download() as liushui_download_info:
                            page.get_by_role("button", name="导出流水").click()
                        download = liushui_download_info.value
                        filename = f"{xiangmuid}_{xiangmu}_兴业银行流水_{start_date}_{end_date}.xlsx"
                        download.save_as(os.path.join(liushui_path, filename))
                        log_local(f"银行流水导出完成：{filename}")
                    except Exception as e:
                        log_local(f"导出银行流水失败（产品：{xiangmuid}_{xiangmu}）：{str(e)}")
                        page.screenshot(path=os.path.join(download_path, f"error_export_excel_{xiangmuid}_{xiangmu}.png"))
                        continue

                    # 如果仅下载流水，跳过回单和对账单
                    if download_only_liushui:
                        page.get_by_role("link", name="明细账单查询 关闭").click()
                        continue

                    # 导出回单
                    try:
                        page.get_by_role("button", name="下载回单").click()

                        with page.expect_download() as huidan_download_info:
                            page.get_by_role("link", name="下载").click()
                        download = huidan_download_info.value
                        filename = f"{xiangmuid}_{xiangmu}_银行回单_{start_date}_{end_date}.pdf"
                        download.save_as(os.path.join(huidan_path, filename))
                        page.get_by_role("link", name="关闭", exact=True).click()
                        log_local(f"银行回单导出完成：{filename}")
                    except Exception as e:
                        log_local(f"导出银行回单失败（产品：{xiangmuid}_{xiangmu}）：{str(e)}")
                        page.screenshot(path=os.path.join(download_path, f"error_export_huidan_{xiangmuid}_{xiangmu}.png"))
                        continue

                except Exception as e:
                    log_local(f"处理产品失败（产品：{xiangmuid}_{xiangmu}）：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_project_{xiangmuid}_{xiangmu}.png"))
                    continue

                # 导出对账单
                page.locator("a").filter(has_text="明细账单查询").click()
                page.get_by_role("button", name=" 选择").click()
                page.get_by_role("textbox", name="产品名称").fill(xiangmuid)
                page.get_by_role("button", name="全选").click()
                page.get_by_role("button", name="确定").click()
                page.get_by_role("textbox", name="交易日期").click()


                page.get_by_role("button", name=" 查询").click()

                try:
                    with page.expect_download() as duizhangdan_download_info:
                        page.get_by_role("button", name="导出PDF").click()
                    download = duizhangdan_download_info.value
                    filename = f"{xiangmuid}_{xiangmu}_银行对账单_{start_date}_{end_date}.pdf"
                    download.save_as(os.path.join(duizhangdan_path, filename))
                    page.get_by_role("link", name="明细账单查询 关闭").click()
                    log_local(f"银行对账单导出完成：{filename}")
                except Exception as e:
                    log_local(f"导出银行对账单失败（产品：{xiangmuid}_{xiangmu}）：{str(e)}")
                    page.screenshot(
                        path=os.path.join(download_path, f"error_export_duizhangdan_{xiangmuid}_{xiangmu}.png"))
                    continue



            log_local(f"完成日期范围 {start_date} 至 {end_date} 的下载")

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
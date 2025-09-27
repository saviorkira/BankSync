from playwright.sync_api import Playwright
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_save_dialog, find_image, split_date_ranges
import os
import time
import pyautogui
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re

def generate_months(start_str, end_str):
    """生成从开始月份到有效结束月份的列表（有效结束 = min(结束, 当前月份-1)）"""
    current_date = datetime.now()
    start = datetime.strptime(start_str, '%Y-%m')
    end = datetime.strptime(end_str, '%Y-%m')
    available_end = current_date - relativedelta(months=1)
    effective_end = min(end, available_end)

    months = []
    current_month = start
    while current_month <= effective_end:
        months.append(current_month.strftime('%Y-%m'))
        current_month += relativedelta(months=1)
    return months

def run_pingan_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None):
    """执行平安银行流水导出"""
    def log_local(msg):
        log(msg, project_root, log_callback)
    log_local("启动平安银行流水导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(project_root, "pingan_bank")
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

        page.get_by_role("textbox", name="企业网银/数字财资/企业用户名").fill(username)

        log_local("等待账户管理页面加载...")
        page.wait_for_selector('text=查询中心', timeout=60000)

        # 外层循环：遍历日期范围
        for range_index, (start_date, end_date) in enumerate(date_ranges):
            log_local(f"处理日期范围: {start_date} 至 {end_date}")

            for index, (xiangmuid, xiangmu, account) in enumerate(projects_accounts):
                log_local(f"处理产品：{xiangmuid}_{xiangmu}，托管账户：{account}")
                try:
                    # 创建流水文件夹
                    folder_name = f"{xiangmuid}_{xiangmu}"
                    duizhang_path = os.path.join(download_path, folder_name, "银行流水")
                    os.makedirs(duizhang_path, exist_ok=True)

                    # 流水查询
                    page.get_by_text("首页").first.click()
                    page.get_by_text("查询中心").click()
                    page.get_by_text("账户查询").click()
                    time.sleep(1)
                    page.get_by_text("新版交易明细查询").click()
                    page.get_by_role("tab", name="历史明细查询").click()
                    page.get_by_role("textbox", name="开始日期").fill(start_date)
                    page.get_by_role("textbox", name="结束日期").fill(end_date)
                    page.get_by_role("textbox", name="结束日期").press("Enter")

                    time.sleep(1)
                    page.get_by_role("textbox", name="请选择").first.click()
                    page.get_by_role("textbox").first.fill(account)
                    page.get_by_role("textbox").first.press("Enter")
                    formatted_account = f"{account[:4]} {account[4:8]} {account[8:12]} {account[12:]}"
                    page.get_by_text(formatted_account).click()
                    page.get_by_role("button", name="查 询").click()
                    time.sleep(2)

                    # 检查是否存在 pingan_zanwushuju.bmp
                    zanwushuju_template = get_resource_path("pingan_zanwushuju.bmp", project_root, subfolder="data/cv2")
                    first_attempt = find_image(zanwushuju_template, project_root, threshold=0.8, max_attempts=1)
                    time.sleep(1)
                    second_attempt = find_image(zanwushuju_template, project_root, threshold=0.8, max_attempts=1)
                    if first_attempt is not None and second_attempt is not None:
                        log_local(
                            f"产品 {xiangmuid}_{xiangmu} 无数据（两次检测均成功，位置：{first_attempt}, {second_attempt}），跳过...")
                        continue
                    else:
                        log_local(
                            f"产品 {xiangmuid}_{xiangmu} 图像检测结果：第一次={'成功' if first_attempt else '失败'}, 第二次={'成功' if second_attempt else '失败'}，继续处理...")

                    page.get_by_role("button", name="下 载 ").click()
                    time.sleep(1)

                    # 导出流水
                    try:
                        with page.expect_download() as liushui_download_info:
                            xiazaiexcel_template = get_resource_path("pingan_xiazaiexcelmingxi.bmp", project_root,
                                                                     subfolder="data/cv2")
                            if not find_and_click_image(xiazaiexcel_template, project_root, threshold=0.8, max_attempts=2):
                                log_local(f"未找到下载 Excel 明细按钮，产品：{xiangmuid}_{xiangmu}")
                                page.screenshot(
                                    path=os.path.join(download_path, f"error_xiazaiexcel_{xiangmuid}_{xiangmu}.png"))
                                continue
                        download = liushui_download_info.value
                        filename = f"{xiangmuid}_{xiangmu}_平安银行流水_{start_date}_{end_date}.xlsx"
                        download.save_as(os.path.join(duizhang_path, filename))
                        log_local(f"银行流水导出完成：{filename}")
                    except Exception as e:
                        log_local(f"导出银行流水失败（产品：{xiangmuid}_{xiangmu}）：{str(e)}")
                        page.screenshot(path=os.path.join(download_path, f"error_export_excel_{xiangmuid}_{xiangmu}.png"))
                        continue

                except Exception as e:
                    log_local(f"处理产品 {xiangmuid}_{xiangmu} 失败：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_{xiangmuid}_{xiangmu}.png"))
                    continue

            log_local(f"完成日期范围 {start_date} 至 {end_date} 的流水下载")

        context.close()
        browser.close()
    except Exception as e:
        log_local(f"平安银行流水导出流程异常：{str(e)}")
        raise
    finally:
        browser.close()
        log_local("浏览器已关闭")
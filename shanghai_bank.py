from playwright.sync_api import Playwright
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_save_dialog, find_image
import os
import time
import pyautogui

def run_shanghai_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None):
    """执行上海银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, project_root, log_callback)
    log_local("启动上海银行导出流程...")
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

        page.wait_for_selector('text=查询中心', timeout=30000)

        for index, (xiangmuid, xiangmu, account) in enumerate(projects_accounts):
            log_local(f"处理产品：{xiangmuid}_{xiangmu}，托管账户：{account}")
            try:
                # 创建文件夹
                folder_name = f"{xiangmuid}_{xiangmu}"
                duizhang_path = os.path.join(download_path, folder_name, "银行流水")
                huidan_path = os.path.join(download_path, folder_name, "银行回单")
                duizhangdan_path = os.path.join(download_path, folder_name, "银行对账单")
                os.makedirs(duizhang_path, exist_ok=True)
                os.makedirs(huidan_path, exist_ok=True)
                os.makedirs(duizhangdan_path, exist_ok=True)

                # 查询账号
                page.get_by_text("查询中心").click()
                page.get_by_text("账户查询").click()
                page.get_by_text("新版交易明细查询").click()
                page.get_by_role("tab", name="历史明细查询").click()
                page.get_by_role("textbox", name="开始日期").fill(kaishiriqi)
                page.get_by_role("textbox", name="结束日期").fill(jieshuriqi)
                page.get_by_role("textbox", name="结束日期").press("Enter")


###################
                page.get_by_role("textbox", name="9000 0000 80").fill(account)
                # 格式化 account，添加空格（如 19036817777777 -> 1903 6817 7777 77）
                formatted_account = f"{account[:4]} {account[4:8]} {account[8:12]} {account[12:]}"
                page.get_by_text(formatted_account).click()
                page.get_by_role("button", name="查 询").click()

                # 检查是否存在 pingan_zanwushuju.bmp
                zanwushuju_template = get_resource_path("pingan_zanwushuju.bmp", project_root, subfolder="data/cv2")
                if find_image(zanwushuju_template, project_root, threshold=0.8, max_attempts=5):
                    log_local(f"产品 {xiangmuid}_{xiangmu} 无数据，跳过...")
                    continue  # 跳到下一个项目

                page.get_by_role("button", name="下 载 ").click()

                # 使用 cv2 识别 pingan_xiazaiexcelmingxi.bmp 并点击
                xiazaiexcel_template = get_resource_path("pingan_xiazaiexcelmingxi.bmp", project_root, subfolder="data/cv2")
                if not find_and_click_image(xiazaiexcel_template, project_root, threshold=0.8, max_attempts=5):
                    log_local(f"未找到下载 Excel 明细按钮，产品：{xiangmuid}_{xiangmu}")
                    page.screenshot(path=os.path.join(download_path, f"error_xiazaiexcel_{xiangmuid}_{xiangmu}.png"))
                    continue

                # 导出流水
                try:
                    with page.expect_download() as liushui_download_info:
                        page.get_by_role("button", name="导出Excel").click()
                    download = liushui_download_info.value
                    filename = f"{xiangmuid}_{xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
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

    except Exception as e:
        log_local(f"上海银行导出流程异常：{str(e)}")
        raise
    finally:
        browser.close()
        log_local("浏览器已关闭")
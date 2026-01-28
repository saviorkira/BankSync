from playwright.sync_api import Playwright
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_save_dialog, find_image, find_all_images, split_date_ranges
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

def run_pingan_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None, download_liushui=False, download_huidan=False, download_duizhangdan=False):
    """执行平安银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, project_root, log_callback)
    log_local("启动平安银行导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    log_local(f"下载选项: 仅流水={'是' if download_liushui else '否'}")
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

    # 添加标志变量，记录是否已处理通知提示
    notification_handled = [False]  # 使用列表以支持闭包修改
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
                    # 创建文件夹
                    folder_name = f"{xiangmuid}_{xiangmu}"
                    duizhang_path = os.path.join(download_path, folder_name, "银行流水")
                    huidan_path = os.path.join(download_path, folder_name, "银行回单")
                    duizhangdan_path = os.path.join(download_path, folder_name, "银行对账单")
                    os.makedirs(duizhang_path, exist_ok=True)
                    os.makedirs(huidan_path, exist_ok=True)
                    os.makedirs(duizhangdan_path, exist_ok=True)
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
                    # page.get_by_role("textbox", name="9000 0000 80").fill(account)
                    # 格式化 account，添加空格（如 19036817777777 -> 1903 6817 7777 77）
                    formatted_account = f"{account[:4]} {account[4:8]} {account[8:12]} {account[12:]}"
                    try:
                        # 等待元素出现，最多等 5 秒
                        page.wait_for_selector(f"text={formatted_account}", timeout=4000)
                        page.get_by_text(formatted_account).click()
                    except TimeoutError:
                        print(f"没有找到: {formatted_account}，跳过")
                        continue
                    # page.get_by_text(formatted_account).click()
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
                        page.get_by_text("首页").first.click()
                        page.get_by_text("查询中心").first.click()
                        page.get_by_text("电子账单").click()
                        continue
                    else:
                        log_local(
                            f"产品 {xiangmuid}_{xiangmu} 图像检测结果：第一次={'成功' if first_attempt else '失败'}, 第二次={'成功' if second_attempt else '失败'}，继续处理...")


                    # if find_image(zanwushuju_template, project_root, threshold=0.8, max_attempts=2):
                    #     log_local(f"产品 {xiangmuid}_{xiangmu} 无数据，跳过...")
                    #     continue  # 跳到下一个项目

                    page.get_by_role("button", name="下 载 ").click()
                    time.sleep(1)

                    # 导出流水
                    if download_liushui:
                        try:
                            with page.expect_download() as liushui_download_info:
                                # 使用 cv2 识别 pingan_xiazaiexcelmingxi.bmp 并点击
                                xiazaiexcel_template = get_resource_path("pingan_xiazaiexcelmingxi.bmp", project_root,
                                                                         subfolder="data/cv2")
                                if not find_and_click_image(xiazaiexcel_template, project_root, threshold=0.8, max_attempts=1):
                                    log_local(f"未找到下载 Excel 明细按钮，产品：{xiangmuid}_{xiangmu}")
                                    page.screenshot(
                                        path=os.path.join(download_path, f"error_xiazaiexcel_{xiangmuid}_{xiangmu}.png"))
                                    continue
                                # page.get_by_role("button", name="导出Excel").click()
                            download = liushui_download_info.value
                            filename = f"{xiangmuid}_{xiangmu}_平安银行流水_{start_date}_{end_date}.xlsx"
                            download.save_as(os.path.join(duizhang_path, filename))
                            log_local(f"银行流水导出完成：{filename}")
                        except Exception as e:
                            log_local(f"导出银行流水失败（产品：{xiangmuid}_{xiangmu}）：{str(e)}")
                            page.screenshot(path=os.path.join(download_path, f"error_export_excel_{xiangmuid}_{xiangmu}.png"))
                            continue
                    else:
                        log_local("未勾选流水下载，跳过银行流水导出")


                    # 导出回单
                    if download_huidan:
                        page.get_by_text("首页").first.click()
                        page.get_by_text("查询中心").first.click()
                        page.get_by_text("电子账单").click()
                        # page.get_by_text("电子回单(新)").first.click()
                        # 账号选择
                        page.locator("text=/请输入账号[0-9]{14}/").click()
                        time.sleep(2)
                        page.get_by_role("combobox").filter(has=page.locator("text=/请输入账号[0-9]{14}/")).get_by_role("textbox").fill(account)
                        time.sleep(2)
                        formatted_account = f"{account[4:8]} {account[8:12]} {account[12:]}"
                        page.get_by_text(formatted_account).click()
                        page.get_by_text("~").first.click()
                        page.get_by_role("textbox", name="开始日期").nth(1).fill(start_date)
                        page.get_by_text("~").first.click()
                        page.get_by_role("textbox", name="结束日期").nth(1).fill(end_date)
                        # 模拟按两次回车
                        pyautogui.click(x=900, y=200)
                        page.get_by_role("button", name="查 询").click()
                        time.sleep(2)

                        # 检查无数据
                        zanwushuju_template = get_resource_path("pingan_zanwushuju.bmp", project_root, subfolder="data/cv2")
                        if find_image(zanwushuju_template, project_root, threshold=0.8, max_attempts=2):
                            log_local(f"产品 {xiangmuid}_{xiangmu} 回单无数据，跳过...")
                            continue

                        page.get_by_role("combobox").filter(has_text="条/页").click()
                        page.get_by_role("option", name="50 条/页").click()
                        time.sleep(1)
                        page.get_by_role("row", name="交易类型 付款方账号/户名 收款方账号/户名 交易币种 交易金额 交易日期 交易用途 操作").get_by_label("").check()
                        time.sleep(2)
                        page.get_by_role("button", name="导出").click()

                        with page.expect_download() as download_info:
                            page.get_by_role("menuitem", name="一页(A4)三张电子回单").click()
                        download = download_info.value
                        filename = f"{xiangmuid}_{xiangmu}_平安银行回单_{start_date}_{end_date}.pdf"
                        download.save_as(os.path.join(huidan_path, filename))
                        log_local(f"银行回单导出完成：{filename}")
                        time.sleep(2)
                    else:
                        log_local("未勾选回单下载，跳过银行回单导出")

                    # 导出对账单
                    if download_duizhangdan:
                        # page.get_by_text("首页").first.click()
                        # page.get_by_text("查询中心").click()
                        # page.get_by_text("电子账单").click()
                        page.get_by_text("电子月结单下载(新)").click()
                        time.sleep(1)
                        page.get_by_placeholder("请输入账号").click()
                        page.get_by_role("textbox", name=re.compile(r"[0-9]{4}\s[0-9]{4}\s[0-9]{2}")).wait_for(state="visible", timeout=30000)
                        time.sleep(1)
                        page.get_by_role("textbox", name=re.compile(r"[0-9]{4}\s[0-9]{4}\s[0-9]{2}")).fill(account)
                        time.sleep(2)
                        page.get_by_role("textbox", name=re.compile(r"[0-9]{4}\s[0-9]{4}\s[0-9]{2}")).press("Enter")
                        time.sleep(1)
                        formatted_account = f"{account[4:8]} {account[8:12]} {account[12:]}"
                        page.get_by_text(formatted_account).click()

                        def format_date(date_str):
                            return date_str[:7]  # 截取 "2025-xx"

                        formatted_start_date = format_date(start_date)
                        formatted_end_date = format_date(end_date)
                        months = generate_months(formatted_start_date, formatted_end_date)
                        if not months:
                            log_local(f"产品 {xiangmuid}_{xiangmu} 无可用对账单月份，跳过...")
                            continue

                        page.get_by_role("textbox", name="开始月份").fill(formatted_start_date)
                        page.get_by_role("textbox", name="结束月份").fill(formatted_end_date)
                        page.get_by_role("button", name="查 询").click()
                        time.sleep(2)

                        # 使用 CV2 找到所有 "下载PDF" 按钮
                        pdf_template = get_resource_path("pingan_xiazaipdf.bmp", project_root, subfolder="data/cv2")
                        points, w, h = find_all_images(pdf_template, project_root, threshold=0.8, max_attempts=3)

                        if len(points) != len(months):
                            log_local(f"警告: 找到 {len(points)} 个按钮，但预期 {len(months)} 个月份，可能数据不全")

                        # 动态映射月份：按钮数量不足时，从末尾开始
                        start_index = max(0, len(months) - len(points))  # 计算起始月份索引
                        for i, pt in enumerate(points):
                            if i >= len(months):
                                log_local(f"警告: 按钮数量多于月份数，跳过多余按钮: {i+1}")
                                break
                            month = months[start_index + i]  # 从后几月开始映射
                            try:
                                with page.expect_download() as download_info:
                                    pyautogui.click(pt[0] + w // 2, pt[1] + h // 2)
                                    time.sleep(1)
                                download = download_info.value
                                filename = f"{xiangmuid}_{xiangmu}_平安银行对账单_{month}.pdf"
                                download.save_as(os.path.join(duizhangdan_path, filename))
                                log_local(f"银行对账单导出完成：{filename}")
                                pyautogui.moveTo(500, 500)
                                pyautogui.click()
                                time.sleep(2)
                                # 从第三个项目开始检查禁用通知提示
                                # if index >= 1 and not notification_handled[0]:
                                #     chrome_notification_path = get_resource_path("chrome_jinyongtongzhi.bmp", project_root)
                                #     if find_and_click_image(chrome_notification_path, project_root, max_attempts=3):
                                #         log_local("检测到并点击‘禁用通知’提示")
                                #         pyautogui.moveTo(495, 495)
                                #         pyautogui.click()
                                #         time.sleep(1)
                                #         notification_handled[0] = True
                                #     else:
                                #         log_local("未检测到‘禁用通知’提示")
                            except Exception as e:
                                log_local(f"导出对账单 {month} 失败：{str(e)}")
                                continue
                    else:
                        log_local("未勾选对账单下载，跳过银行对账单导出")

                except Exception as e:
                    log_local(f"处理产品 {xiangmuid}_{xiangmu} 失败：{str(e)}")
                    page.screenshot(path=os.path.join(download_path, f"error_{xiangmuid}_{xiangmu}.png"))
                    continue


            log_local(f"完成日期范围 {start_date} 至 {end_date} 的下载")

        context.close()
        browser.close()
    except Exception as e:
        log_local(f"平安银行导出流程异常：{str(e)}")
        raise
    finally:
        browser.close()
        log_local("浏览器已关闭")
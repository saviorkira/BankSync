import os
import re
import subprocess
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright

# 参数设置
xiangmu = "重信·北极22026·华睿精选9号集合资金信托计划"
kaishiriqi = "2025-03-01"
jieshuriqi = "2025-03-31"

# 日志记录函数
def log_error(message: str):
    log_dir = "D:\\Desktop\\导出日志"
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "导出错误日志.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()} - {message}\n")

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.e-custody.com/#/login")

    # 登录
    page.get_by_role("textbox", name="用户名").fill("18323580933")
    page.get_by_role("textbox", name="请输入您的密码").fill("2780zjj?")

    # 等待验证码输入和手动操作
    input("请手动完成验证码输入并登录后按回车继续...")

    # 页面导航
    page.wait_for_selector('text=账户管理', timeout=30000)
    page.get_by_role("link", name="账户管理").click()
    page.get_by_role("link", name="账户明细").click()

    # 填写项目信息
    page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
    page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").fill(xiangmu)
    page.get_by_role("listitem").filter(has_text=xiangmu).locator("span").nth(2).click()

    # 设置日期
    page.get_by_text("展开").first.click()
    page.get_by_role("textbox", name="开始日期").fill(kaishiriqi)
    page.get_by_role("textbox", name="开始日期").press("Enter")
    page.get_by_role("textbox", name="结束日期").fill(jieshuriqi)
    page.get_by_role("textbox", name="结束日期").press("Enter")

    # 执行查询
    page.get_by_role("button", name=" 查询").click()

    # 导出对账单
    page.get_by_role("button", name="导出 ").click()
    with page.expect_download() as download_info:
        page.get_by_text("对账单导出", exact=True).click()
    download = download_info.value

    # 保存对账单
    safe_xiangmu = re.sub(r'[\\/:*?<>|"\n]+', "_", xiangmu)
    folder_path = fr"D:\\Desktop\\{safe_xiangmu}\\银行流水"
    os.makedirs(folder_path, exist_ok=True)
    filename = f"{safe_xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
    save_path = os.path.join(folder_path, filename)

    try:
        download.save_as(save_path)
        print(f"对账单已保存到：{save_path}")
    except Exception as e:
        log_error(f"对账单保存失败：{e}")

    # 导出银行回单
    page.get_by_role("checkbox", name="Toggle Selection of All Rows").check()
    page.get_by_role("button", name="导出 ").click()
    with page.expect_download() as download1_info:
        page.get_by_text("对账单导出", exact=True).press("ArrowDown")
        page.get_by_text("英文对账单导出").press("ArrowDown")
    download1 = download1_info.value

    receipt_folder = fr"D:\\Desktop\\{safe_xiangmu}\\银行回单"
    os.makedirs(receipt_folder, exist_ok=True)
    receipt_filename = f"{safe_xiangmu}_银行回单_{kaishiriqi}_{jieshuriqi}.pdf"
    receipt_path = os.path.join(receipt_folder, receipt_filename)

    try:
        download1.save_as(receipt_path)
        print(f"银行回单已保存到：{receipt_path}")
    except Exception as e:
        log_error(f"银行回单保存失败：{e}")

    # 打开文件夹位置
    try:
        subprocess.Popen(f'explorer "{os.path.dirname(folder_path)}"')
    except Exception as e:
        log_error(f"打开文件夹失败：{e}")

    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright)

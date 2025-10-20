import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://tgt.cib.com.cn/custody/")


    page.get_by_role("textbox", name="用户名/邮箱/手机号").fill("cqxt")
    page.get_by_role("textbox", name="密码").fill("333222aa")



    page.locator("a").filter(has_text="账务查询").click()
    page.locator("a").filter(has_text=re.compile(r"^流水查询$")).click()


    page.get_by_role("textbox", name="请点击选择按钮 选择产品").click()
    page.get_by_role("button", name=" 选择").click()
    page.get_by_role("textbox", name="产品名称").fill("秀") #项目
    page.get_by_role("textbox", name="产品名称").click()

    page.get_by_role("textbox", name="交易日期").fill("2025-09-29 - 2025-09-29") #日期

    page.get_by_role("row", name="付款账户名 产品名称 交易日期 交易时间 交易金额 账户余额 摘要说明 支付备注 付款方行名 付款账号 收款方行名 收款账户名 收款账号 流水号").locator("span").nth(1).click()
    

    with page.expect_download() as download1_info:
        page.get_by_role("button", name="导出流水").click()
    download1 = download1_info.value




    # page.get_by_role("textbox", name="请选择模板类型").click()
    # page.get_by_text("一页两条回单").click()
    page.get_by_role("button", name="下载回单").click()
    with page.expect_download() as download_info:
        page.get_by_role("link", name="下载").click()
    download = download_info.value
    page.get_by_role("link", name="关闭", exact=True).click()






    page.locator("a").filter(has_text="明细账单查询").click()
    page.get_by_role("link", name="流水查询 关闭").click()
    page.get_by_role("button", name=" 选择").click()
    page.get_by_role("textbox", name="产品名称").click()
    page.get_by_role("textbox", name="产品名称").fill("渝信")
    page.get_by_role("button", name="全选").click()
    page.get_by_role("button", name="确定").click()
    page.get_by_role("textbox", name="交易日期").click()
    page.get_by_role("cell", name="1", exact=True).first.click()
    page.get_by_role("cell", name="22").first.click()
    page.get_by_role("button", name=" 查询").click()
    with page.expect_download() as download3_info:
        page.get_by_role("button", name="导出PDF").click()
    download3 = download3_info.value








    page.get_by_role("link", name="银企对账 关闭").click()
    page.get_by_role("link", name="明细账单查询 关闭").click()


    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

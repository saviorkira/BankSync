import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://auth.orangebank.com.cn/cimp-ccs-pc/#/p/ebank-login")
    page.get_by_role("textbox", name="企业网银/数字财资/企业用户名").click()
    page.get_by_role("textbox", name="企业网银/数字财资/企业用户名").fill("89875338")
    page.get_by_text("请输入登录密码").first.click()
    page.get_by_role("textbox", name="企业网银/数字财资/企业用户名").click()
    page.get_by_role("textbox", name="企业网银/数字财资/企业用户名").fill("LINJIA")
    page.get_by_text("请输入登录密码").first.click()
    page.get_by_role("button", name="登录").click()
    page.get_by_text("查询中心").click()
    page.get_by_text("查询中心").click()
    page.get_by_text("账户查询").click()
    page.get_by_text("新版交易明细查询").click()


    page.locator("div").filter(has_text=re.compile(r"^账号$")).get_by_placeholder("请选择").click()
    page.get_by_role("textbox", name="9000 0000 80").fill("19036817777777")
    page.get_by_text("6817 7777 77").click()

    page.get_by_role("tab", name="历史明细查询").click()
    page.locator("div").filter(has_text=re.compile(r"^账号$")).get_by_placeholder("请选择").click()
    page.get_by_role("textbox", name="6817 7777 77").fill("19036817777777")
    page.get_by_text("6817 7777 77").click()
    page.get_by_role("textbox", name="开始日期").click()
    page.get_by_text("1", exact=True).first.click()
    page.get_by_text("30").nth(1).click()
    page.get_by_role("textbox", name="结束日期").click()
    page.get_by_text("30").nth(1).click()
    page.get_by_text("1", exact=True).first.click()
    page.get_by_role("button", name="查 询").click()
    page.get_by_role("button", name="下 载 ").click()
    with page.expect_download() as download_info:
        page.get_by_text("下载PDF").click()
    download = download_info.value
    page.locator("a").filter(has_text="展开").click()
    page.locator("a").filter(has_text="收起").click()
    page.locator(".elp-table__fixed-body-wrapper > .elp-table__body > tbody > tr:nth-child(5) > .elp-table_3_column_40 > .cell > .f > .elp-dropdown > .elp-dropdown-link").click()
    page.get_by_role("button", name="打 印").click()
    page.get_by_role("button", name="Close").click()
    page.get_by_text("查询中心").click()
    page.get_by_text("账户查询").click()
    page.get_by_text("新版交易明细查询").click()
    page.get_by_role("tab", name="历史明细查询").click()
    page.get_by_role("textbox", name="请选择").first.click()
    page.get_by_role("textbox", name="9000 0000 80").fill("19036817777777")
    page.get_by_text("6817 7777 77").click()
    page.get_by_role("textbox", name="开始日期").click()
    page.get_by_text("1", exact=True).first.click()
    page.get_by_text("30").nth(1).click()
    page.get_by_role("textbox", name="结束日期").click()
    page.get_by_text("1", exact=True).first.click()
    page.get_by_text("30").nth(1).click()
    page.get_by_role("button", name="查 询").click()
    page.get_by_role("button", name="下 载 ").click()
    with page.expect_download() as download1_info:
        page.locator("#dropdown-menu-5177").get_by_text("下载EXCEL").click()
    download1 = download1_info.value
    page.locator("label").filter(has_text="查询中心").click()
    page.get_by_text("电子账单").click()
    page.get_by_role("combobox").filter(has_text="请输入账号").click()
    page.get_by_role("combobox").filter(has_text="请输入账号").get_by_role("textbox").fill("19036817777777")
    page.get_by_text("6817 7777 77").click()
    page.get_by_role("textbox", name="开始日期").click()
    page.get_by_title("年6月1日").locator("div").click()
    page.get_by_text("30").nth(1).click()
    page.get_by_role("textbox", name="结束日期").click()
    page.get_by_title("年6月1日").locator("div").click()
    page.get_by_text("30").nth(1).click()
    page.get_by_role("button", name="查 询").click()
    page.get_by_role("row", name="交易类型 付款方账号/户名 收款方账号/户名 交易币种 交易金额 交易日期 交易用途 操作").get_by_label("").check()
    page.get_by_role("button", name="导出").click()
    with page.expect_download() as download2_info:
        page.get_by_role("menuitem", name="一页(A4)三张电子回单").click()
    download2 = download2_info.value

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

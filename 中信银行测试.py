import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://ib.citicbank.com/html/#/index")
    page.locator("#new_header").get_by_text("登录").click()
    page.get_by_role("textbox", name="手机号").click()
    page.get_by_role("textbox", name="手机号").fill("13983426184")
    page.locator("#PwdIdBoxUkeyChrome_login #noUkeyPwd_str_login").click()
    page.locator("input[type=\"password\"]").click()
    page.locator("input[type=\"password\"]").fill("********")
    page.get_by_role("button", name="登录").click()
    page.get_by_text("会员中心").click()
    page.get_by_role("link", name="托管业务 ").click()
    page.get_by_role("link", name="托管账户 ").click()
    page.get_by_role("link", name="托管账户明细查询").click()

#选择项目
    page.locator("#inputPro").click()
    page.locator("#inputPro").fill("简州空港")
    page.locator("#inputPro").press("Enter")
    page.locator("#ulPro").click()

#选择日期
    page.locator("input[name=\"startDate\"]").click()
    page.get_by_role("cell", name="六月").click()
    page.get_by_role("cell", name="2025").click()

    page.locator("input[name=\"startDate\"]").click()
    page.get_by_role("cell", name="5", exact=True).first.click()
    page.get_by_text("查询", exact=True).click()

    page.locator("input[name=\"startDate\"]").click()

    page.get_by_role("cell", name="六月").click()
    page.get_by_role("cell", name="2025", exact=True).click()

    page.get_by_text("2025", exact=True).nth(1).click()
    page.get_by_text("六月").nth(3).click()
    page.get_by_role("cell", name="24", exact=True).click()

    page.locator("input[name=\"endDate\"]").click()
    page.locator("input[name=\"endDate\"]").click()
    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

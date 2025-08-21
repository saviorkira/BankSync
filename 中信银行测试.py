import os
from playwright.sync_api import sync_playwright

def log_local(msg):
    print("[调试]", msg)

if __name__ == "__main__":
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
        os.path.abspath(os.path.dirname(__file__)), "playwright-browsers"
    )
    print("设置 PLAYWRIGHT_BROWSERS_PATH:", os.environ["PLAYWRIGHT_BROWSERS_PATH"])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

        # 伪装成 IE11 的 UA
        ua_ie11 = "Mozilla/5.0 (Windows NT 10.0; Trident/7.0; rv:11.0) like Gecko"

        context = browser.new_context(
            viewport=None,
            accept_downloads=True,
            user_agent=ua_ie11,
        )

        page = context.new_page()
        log_local(f"当前 User-Agent: {ua_ie11}")

        # 民生银行测试
        page.goto("https://ent.cmbc.com.cn/trust-bank/#/login?_k=ur38w6")

        # 暂停手动操作
        page.pause()

        context.close()
        browser.close()

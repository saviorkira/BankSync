import os
import sys
import code
from playwright.sync_api import sync_playwright

# 假设日志函数如下（你可以替换成 log()）
def log_local(msg):
    print("[交互调试]", msg)

if __name__ == "__main__":
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
        os.path.abspath(os.path.dirname(__file__)), "playwright-browsers"
    )
    print("设置 PLAYWRIGHT_BROWSERS_PATH:", os.environ["PLAYWRIGHT_BROWSERS_PATH"])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(viewport=None)
        page = context.new_page()

        page.goto("https://www.e-custody.com/#/login")  # 你可以换成你的银行地址
        # page.get_by_role("textbox", name="用户名").fill("18323580933")
        # page.get_by_role("textbox", name="请输入您的密码").fill("2780zjj?")
        # page.wait_for_selector('text=账户管理', timeout=90000)
        # page.get_by_role("link", name="账户管理").click()
        # page.get_by_role("link", name="账户明细").click()


        page.pause()  # ✅ 打开 Inspector 并暂停

        # 🔧 打开交互式控制台，手动输入指令操作 page
        # print("\n你现在可以使用 Python 控制 page，如：")
        # print("page.get_by_role('button', name='打印 ').click()")
        # print("menus = page.locator(\"css=[id^='dropdown-menu-']\")")
        # print("menus.count() 或 menus.nth(0).inner_text() 等")

        # 启动交互式 shell
        code.interact(local=locals())

        context.close()
        browser.close()

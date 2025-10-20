import os
import sys
import code
from playwright.sync_api import sync_playwright
import time
import pyautogui
import pygetwindow as gw
import keyboard
import re

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
        context = browser.new_context(
            viewport=None,
            accept_downloads=True,
            # downloads_path="D:/Downloads"  # 设置下载路径
        )
        page = context.new_page()
        # page.goto("https://www.e-custody.com/#/")
        page.goto("https://ent.cmbc.com.cn/trust-bank/#/login?_k=ur38w6")  # 你可以换成你的银行地址 民生
        # page.goto("https://tgt.cib.com.cn/custody/")  # 你可以换成你的银行地址 兴业
        # page.goto("https://auth.orangebank.com.cn/cimp-ccs-pc/#/p/ebank-login")  # 你可以换成你的银行地址 平安
        # page.goto("https://ib.citicbank.com/html/#/index")  # 你可以换成你的银行地址 中信
        # page.goto("https://custody.hzbank.com.cn/#/login")  # 你可以换成你的银行地址 中信
        # page.get_by_role("textbox", name="请输入客户号").fill("800067725")
        # page.get_by_role("textbox", name="请输入操作员号").fill("2001")
        # page.get_by_role("textbox", name="请输入登录密码").click()


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

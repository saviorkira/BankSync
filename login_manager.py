import os
import time
import threading
import flet as ft
from playwright.sync_api import sync_playwright
from utils import read_bank_config

# 网站登录处理函数映射
SITE_HANDLERS = {
    "ningbo_bank": lambda playwright, project_root, update_log: login_ningbo_bank(playwright, project_root, update_log),
    "hangzhou_bank": lambda playwright, project_root, update_log: login_hangzhou_bank(playwright, project_root, update_log),
    "pingan_bank": lambda playwright, project_root, update_log: login_hangzhou_bank(playwright, project_root, update_log),
    "shanghai_bank": lambda playwright, project_root, update_log: login_hangzhou_bank(playwright, project_root, update_log),
    "zheshang_bank": lambda playwright, project_root, update_log: login_hangzhou_bank(playwright, project_root, update_log),
    "zhongxin_bank": lambda playwright, project_root, update_log: login_hangzhou_bank(playwright, project_root, update_log),

    # 未来添加其他网站，例如：
    # "huaxia_bank": lambda playwright, project_root, update_log: login_huaxia_bank(playwright, project_root, update_log),
    # "other_site": lambda playwright, project_root, update_log: login_other_site(playwright, project_root, update_log),
}

def load_site_icons(project_root: str, update_log, login_callback):
    """加载网站图标，返回图标控件列表"""
    login_dir = os.path.join(project_root, "data", "login")
    update_log(f"扫描图标目录: {login_dir}")
    icons = []
    if os.path.exists(login_dir):
        for file in os.listdir(login_dir):
            if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                icon_path = os.path.join(login_dir, file)
                site_name = os.path.splitext(file)[0]
                update_log(f"找到图标: {icon_path}")
                icon = ft.Image(
                    src=icon_path,
                    width=100,
                    height=100,
                    fit=ft.ImageFit.CONTAIN,
                    tooltip=f"登录 {site_name.replace('_', ' ')}",
                )
                icon_container = ft.Container(
                    content=icon,
                    on_click=lambda e, name=site_name: login_callback(name),
                    padding=5,
                    border_radius=8,
                    ink=True,
                )
                icons.append(icon_container)
    else:
        update_log(f"图标目录不存在: {login_dir}")
    if not icons:
        update_log("未找到任何网站图标")
    return icons

def login_ningbo_bank(playwright, project_root, update_log):
    """宁波银行登录逻辑"""
    try:
        update_log("启动宁波银行登录流程...")
        update_log(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
        try:
            username, password, login_url, config_path = read_bank_config(project_root, "ningbo_bank")
            update_log(f"加载配置文件: {config_path}")
        except Exception as e:
            update_log(f"config.json 文件加载失败: {str(e)}")
            return
        browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
        if not os.path.exists(browser_path):
            update_log(f"Playwright 浏览器路径不存在: {browser_path}")
            return
        update_log("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(120000)
        update_log(f"访问登录页面: {login_url}")
        page.goto(login_url)
        update_log("输入用户名和密码...")
        page.get_by_role("textbox", name="用户名").fill(username)
        page.get_by_role("textbox", name="请输入您的密码").fill(password)
        update_log("用户名和密码已输入，请手动执行后续操作...")
        update_log("浏览器窗口将保持打开状态，手动关闭浏览器以继续程序...")
        while True:
            time.sleep(1)
            if not browser.contexts:
                update_log("检测到浏览器已关闭，结束登录流程")
                break
        context.close()
        browser.close()
        update_log("宁波银行登录流程完成")
    except Exception as e:
        update_log(f"宁波银行登录失败：{str(e)}")
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()


def login_hangzhou_bank(playwright, project_root, update_log):
    """杭州银行登录逻辑"""
    try:
        update_log("启动杭州银行登录流程...")
        update_log(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
        try:
            username, password, login_url, config_path = read_bank_config(project_root, "hangzhou_bank")
            update_log(f"加载配置文件: {config_path}")
        except Exception as e:
            update_log(f"config.json 文件加载失败: {str(e)}")
            return
        browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
        if not os.path.exists(browser_path):
            update_log(f"Playwright 浏览器路径不存在: {browser_path}")
            return  # 修正缩进，确保与 if 块同级
        update_log("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(120000)
        update_log(f"访问登录页面: {login_url}")
        page.goto(login_url)
        update_log("输入用户名和密码...")
        page.get_by_role("textbox", name="请输入客户号").fill(username)
        page.get_by_role("textbox", name="请输入操作员号").fill("2001")
        update_log("用户名和操作员号已输入，请手动执行后续操作...")
        update_log("浏览器窗口将保持打开状态，手动关闭浏览器以继续程序...")
        while True:
            time.sleep(1)
            if not browser.contexts:
                update_log("检测到浏览器已关闭，结束登录流程")
                break
        context.close()
        browser.close()
        update_log("杭州银行登录流程完成")
    except Exception as e:
        update_log(f"杭州银行登录失败：{str(e)}")
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()





def login_site(site_name: str, project_root: str, update_log, last_click_time: list):
    """处理网站登录逻辑"""
    current_time = time.time()
    if current_time - last_click_time[0] < 1:
        update_log(f"点击过于频繁，请稍后再试")
        return
    last_click_time[0] = current_time

    if site_name not in SITE_HANDLERS:
        update_log(f"暂不支持 {site_name} 的登录")
        return

    def login_worker():
        try:
            playwright = sync_playwright().start()  # 手动启动 playwright
            try:
                SITE_HANDLERS[site_name](playwright, project_root, update_log)
            finally:
                playwright.stop()  # 手动停止 playwright
        except Exception as ex:
            update_log(f"登录 {site_name} 失败：{str(ex)}")

    threading.Thread(target=login_worker, daemon=True).start()
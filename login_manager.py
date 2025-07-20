import os
import time
import threading
import flet as ft
from playwright.sync_api import sync_playwright
from utils import read_bank_config, log  # 导入 log 函数以备回退
import json
import sys

# 网站登录处理函数映射
SITE_HANDLERS = {
    "ningbo_bank": lambda playwright, project_root, update_log: login_ningbo_bank(playwright, project_root, update_log),
    "hangzhou_bank": lambda playwright, project_root, update_log: login_hangzhou_bank(playwright, project_root, update_log),
    "TM": lambda playwright, project_root, update_log: login_generic_site(playwright, project_root, update_log, "TM"),
    "TA": lambda playwright, project_root, update_log: login_generic_site(playwright, project_root, update_log, "TA"),
    "AM": lambda playwright, project_root, update_log: login_generic_site(playwright, project_root, update_log, "AM"),
    "kuaiji": lambda playwright, project_root, update_log: login_generic_site(playwright, project_root, update_log, "kuaiji"),
    "jianguan": lambda playwright, project_root, update_log: login_generic_site(playwright, project_root, update_log, "jianguan"),
}

def load_site_icons(project_root: str, update_log, login_callback):
    """加载网站图标，返回图标控件列表"""
    if getattr(sys, 'frozen', False):  # 打包环境
        project_root = os.path.dirname(sys.executable)
        update_log(f"检测到打包环境，project_root 设置为: {project_root}")
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

def get_cookies_path(project_root: str, site_name: str):
    """返回 cookies 文件的存储路径，适配开发和打包环境"""
    if getattr(sys, 'frozen', False):  # 打包环境
        project_root = os.path.dirname(sys.executable)
    cookies_path = os.path.join(project_root, "data", "cookies", f"{site_name}_cookies.json")
    return cookies_path

def login_ningbo_bank(playwright, project_root, update_log):
    """宁波银行登录逻辑，支持 cookies"""
    def safe_log(msg):
        """安全日志记录函数"""
        if callable(update_log):
            update_log(msg)
        else:
            log(f"日志记录失败 (ningbo_bank): {msg}", project_root)

    try:
        safe_log("启动宁波银行登录流程...")
        safe_log(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")

        # 加载配置文件
        try:
            username, password, login_url, config_path = read_bank_config(project_root, "ningbo_bank")
            safe_log(f"加载配置文件: {config_path}")
        except Exception as e:
            safe_log(f"config.json 文件加载失败: {str(e)}")
            return

        # 检查 Playwright 浏览器路径
        browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
        if not os.path.exists(browser_path):
            safe_log(f"Playwright 浏览器路径不存在: {browser_path}")
            return

        # 尝试加载现有 cookies
        cookies_path = get_cookies_path(project_root, "ningbo_bank")
        context = None
        if os.path.exists(cookies_path):
            safe_log(f"找到已保存的 cookies: {cookies_path}")
            try:
                with open(cookies_path, 'r', encoding='utf-8') as f:
                    storage_state = json.load(f)
                context = playwright.chromium.new_context(storage_state=storage_state, viewport=None)
                page = context.new_page()
                page.set_default_timeout(120000)
                safe_log(f"访问登录页面以验证 cookies: {login_url}")
                page.goto(login_url)

                # 检查是否已登录
                if page.query_selector("text=退出登录"):
                    safe_log("Cookies 有效，已自动登录")
                    page.wait_for_timeout(5000)
                    context.close()
                    return
                else:
                    safe_log("Cookies 无效，重新登录")
                    context.close()
            except Exception as e:
                safe_log(f"加载 cookies 失败: {str(e)}")
                if context:
                    context.close()

        # 执行登录流程
        safe_log("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(120000)
        safe_log(f"访问登录页面: {login_url}")
        page.goto(login_url)
        safe_log("输入用户名和密码...")
        page.get_by_role("textbox", name="用户名").fill(username)
        page.get_by_role("textbox", name="请输入您的密码").fill(password)
        safe_log("用户名和密码已输入，请手动执行后续操作...")
        safe_log("浏览器窗口将保持打开状态，手动关闭浏览器以继续程序...")

        # 等待浏览器关闭并保存 cookies
        while True:
            time.sleep(1)
            if not browser.contexts:
                safe_log("检测到浏览器已关闭，尝试保存 cookies")
                try:
                    storage_state = context.storage_state()
                    cookies_dir = os.path.dirname(cookies_path)
                    os.makedirs(cookies_dir, exist_ok=True)
                    safe_log(f"确保 cookies 目录存在: {cookies_dir}")
                    if not os.access(cookies_dir, os.W_OK):
                        safe_log(f"错误：无写入权限: {cookies_dir}")
                        return
                    with open(cookies_path, 'w', encoding='utf-8') as f:
                        json.dump(storage_state, f, indent=2)
                    safe_log(f"Cookies 已保存至: {cookies_path}")
                    if os.path.exists(cookies_path):
                        safe_log(f"确认 Cookies 文件存在: {cookies_path}, 大小: {os.path.getsize(cookies_path)} 字节")
                    else:
                        safe_log(f"错误：Cookies 文件未生成: {cookies_path}")
                except Exception as e:
                    safe_log(f"保存 cookies 失败: {str(e)}")
                break
        context.close()
        browser.close()
        safe_log("宁波银行登录流程完成")
    except Exception as e:
        safe_log(f"宁波银行登录失败：{str(e)}")
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()

def login_hangzhou_bank(playwright, project_root, update_log):
    """杭州银行登录逻辑，支持 cookies"""
    def safe_log(msg):
        """安全日志记录函数"""
        if callable(update_log):
            update_log(msg)
        else:
            log(f"日志记录失败 (hangzhou_bank): {msg}", project_root)

    try:
        safe_log("启动杭州银行登录流程...")
        safe_log(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")

        # 加载配置文件
        try:
            username, password, login_url, config_path = read_bank_config(project_root, "hangzhou_bank")
            safe_log(f"加载配置文件: {config_path}")
        except Exception as e:
            safe_log(f"config.json 文件加载失败: {str(e)}")
            return

        # 检查 Playwright 浏览器路径
        browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
        if not os.path.exists(browser_path):
            safe_log(f"Playwright 浏览器路径不存在: {browser_path}")
            return

        # 尝试加载现有 cookies
        cookies_path = get_cookies_path(project_root, "hangzhou_bank")
        context = None
        if os.path.exists(cookies_path):
            safe_log(f"找到已保存的 cookies: {cookies_path}")
            try:
                with open(cookies_path, 'r', encoding='utf-8') as f:
                    storage_state = json.load(f)
                context = playwright.chromium.new_context(storage_state=storage_state, viewport=None)
                page = context.new_page()
                page.set_default_timeout(120000)
                safe_log(f"访问登录页面以验证 cookies: {login_url}")
                page.goto(login_url)

                # 检查是否已登录
                if page.query_selector("text=退出登录"):
                    safe_log("Cookies 有效，已自动登录")
                    page.wait_for_timeout(5000)
                    context.close()
                    return
                else:
                    safe_log("Cookies 无效，重新登录")
                    context.close()
            except Exception as e:
                safe_log(f"加载 cookies 失败: {str(e)}")
                if context:
                    context.close()

        # 执行登录流程
        safe_log("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(120000)
        safe_log(f"访问登录页面: {login_url}")
        page.goto(login_url)
        safe_log("输入用户名和密码...")
        page.get_by_role("textbox", name="请输入客户号").fill(username)
        page.get_by_role("textbox", name="请输入操作员号").fill("2001")
        safe_log("用户名和操作员号已输入，请手动执行后续操作...")
        safe_log("浏览器窗口将保持打开状态，手动关闭浏览器以继续程序...")

        # 等待浏览器关闭并保存 cookies
        while True:
            time.sleep(1)
            if not browser.contexts:
                safe_log("检测到浏览器已关闭，尝试保存 cookies")
                try:
                    storage_state = context.storage_state()
                    cookies_dir = os.path.dirname(cookies_path)
                    os.makedirs(cookies_dir, exist_ok=True)
                    safe_log(f"确保 cookies 目录存在: {cookies_dir}")
                    if not os.access(cookies_dir, os.W_OK):
                        safe_log(f"错误：无写入权限: {cookies_dir}")
                        return
                    with open(cookies_path, 'w', encoding='utf-8') as f:
                        json.dump(storage_state, f, indent=2)
                    safe_log(f"Cookies 已保存至: {cookies_path}")
                    if os.path.exists(cookies_path):
                        safe_log(f"确认 Cookies 文件存在: {cookies_path}, 大小: {os.path.getsize(cookies_path)} 字节")
                    else:
                        safe_log(f"错误：Cookies 文件未生成: {cookies_path}")
                except Exception as e:
                    safe_log(f"保存 cookies 失败: {str(e)}")
                break
        context.close()
        browser.close()
        safe_log("杭州银行登录流程完成")
    except Exception as e:
        safe_log(f"杭州银行登录失败：{str(e)}")
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()

def login_generic_site(playwright, project_root, update_log, site_name):
    """通用网站登录逻辑，支持 cookies"""
    def safe_log(msg):
        """安全日志记录函数"""
        if callable(update_log):
            update_log(msg)
        else:
            log(f"日志记录失败 ({site_name}): {msg}", project_root)

    try:
        safe_log(f"启动 {site_name} 登录流程...")
        safe_log(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")

        # 加载配置文件
        try:
            username, password, login_url, config_path = read_bank_config(project_root, site_name)
            safe_log(f"加载配置文件: {config_path}")
        except Exception as e:
            safe_log(f"config.json 文件加载失败: {str(e)}")
            return

        # 检查 Playwright 浏览器路径
        browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
        if not os.path.exists(browser_path):
            safe_log(f"Playwright 浏览器路径不存在: {browser_path}")
            return

        # 尝试加载现有 cookies
        cookies_path = get_cookies_path(project_root, site_name)
        context = None
        if os.path.exists(cookies_path):
            safe_log(f"找到已保存的 cookies: {cookies_path}")
            try:
                with open(cookies_path, 'r', encoding='utf-8') as f:
                    storage_state = json.load(f)
                context = playwright.chromium.new_context(storage_state=storage_state, viewport=None)
                page = context.new_page()
                page.set_default_timeout(120000)
                safe_log(f"访问登录页面以验证 cookies: {login_url}")
                page.goto(login_url)
                safe_log(f"页面加载完成，当前 URL: {page.url}")
                if "login" not in page.url.lower() and "dashboard" in page.url.lower():
                    safe_log("警告：可能访问了登录后页面，请检查 login_url")

                # 检查是否已登录
                if page.query_selector("text=退出|Logout|登出") or "dashboard" in page.url:
                    safe_log("Cookies 有效，已自动登录")
                    page.wait_for_timeout(5000)
                    context.close()
                    return
                else:
                    safe_log("Cookies 无效，重新登录")
                    context.close()
            except Exception as e:
                safe_log(f"加载 cookies 失败: {str(e)}")
                if context:
                    context.close()

        # 执行登录流程
        safe_log("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(120000)
        safe_log(f"访问登录页面: {login_url}")
        try:
            page.goto(login_url)
            safe_log(f"页面加载完成，当前 URL: {page.url}")
            if "login" not in page.url.lower() and "dashboard" in page.url.lower():
                safe_log("警告：可能访问了登录后页面，请检查 login_url")
        except Exception as e:
            safe_log(f"页面加载失败: {str(e)}")
            return
        safe_log("请手动输入用户名、密码和可能的验证码...")
        safe_log("浏览器窗口将保持打开状态，手动关闭浏览器以继续程序...")

        # 等待浏览器关闭并保存 cookies
        while True:
            time.sleep(1)
            if not browser.contexts:
                safe_log("检测到浏览器已关闭，尝试保存 cookies")
                try:
                    storage_state = context.storage_state()
                    cookies_dir = os.path.dirname(cookies_path)
                    os.makedirs(cookies_dir, exist_ok=True)
                    safe_log(f"确保 cookies 目录存在: {cookies_dir}")
                    if not os.access(cookies_dir, os.W_OK):
                        safe_log(f"错误：无写入权限: {cookies_dir}")
                        return
                    with open(cookies_path, 'w', encoding='utf-8') as f:
                        json.dump(storage_state, f, indent=2)
                    safe_log(f"Cookies 已保存至: {cookies_path}")
                    if os.path.exists(cookies_path):
                        safe_log(f"确认 Cookies 文件存在: {cookies_path}, 大小: {os.path.getsize(cookies_path)} 字节")
                    else:
                        safe_log(f"错误：Cookies 文件未生成: {cookies_path}")
                except Exception as e:
                    safe_log(f"保存 cookies 失败: {str(e)}")
                break
        context.close()
        browser.close()
        safe_log(f"{site_name} 登录流程完成")
    except Exception as e:
        safe_log(f"{site_name} 登录失败：{str(e)}")
        if 'context' in locals():
            context.close()
        if 'browser' in locals():
            browser.close()

def login_site(site_name: str, project_root: str, update_log, last_click_time: list):
    """处理网站登录逻辑"""
    if not callable(update_log):
        log(f"错误: update_log 不是可调用的函数: {update_log}", project_root)
        return

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
                update_log(f"调用 {site_name} 的登录处理函数")
                SITE_HANDLERS[site_name](playwright, project_root, update_log)
            finally:
                playwright.stop()  # 手动停止 playwright
        except Exception as ex:
            update_log(f"登录 {site_name} 失败：{str(ex)}")

    threading.Thread(target=login_worker, daemon=True).start()
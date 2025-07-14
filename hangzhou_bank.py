from playwright.sync_api import Playwright
from utils import log, read_bank_config, get_resource_path, find_and_click_image, handle_overwrite_dialog, handle_save_dialog
import os
import time
import pyautogui

def run_hangzhou_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi, log_callback=None):
    """执行宁波银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, download_path, log_callback)  # 日志保存到 download_path
    log_local("启动杭州银行导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(project_root)
        log_local(f"加载配置文件: {config_path}")
    except Exception as e:
        log_local(f" config.txt 文件加载失败: {str(e)}")
        raise
    browser_path = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
    if not os.path.exists(browser_path):
        log_local(f"Playwright 浏览器路径不存在: {browser_path}")
        raise FileNotFoundError(f"Playwright 浏览器路径不存在: {browser_path}")
    try:
        log_local("启动浏览器...")
        browser = playwright.chromium.launch(headless=False, timeout=30000)
        context = browser.new_context(viewport=None)
        page = context.new_page()
        page.set_default_timeout(120000)
        log_local(f"访问登录页面: {login_url}")
        page.goto(login_url)
        log_local("输入用户名和密码...")
        page.get_by_role("textbox", name="用户名").fill(username)
        page.get_by_role("textbox", name="请输入操作员号").fill("2001")
        # page.get_by_role("textbox", name="请输入您的密码").fill(password)
        log_local("等待账户管理页面加载...")
        page.wait_for_selector('text=账户', timeout=60000)
        page.get_by_text("账户", exact=True).click()
        page.get_by_role("menuitem", name="流水查询").click()
        page.get_by_role("combobox", name="开始日期").fill("kaishiriqi")
        page.get_by_role("combobox", name="结束日期").fill("jieshuriqi")
    except Exception as e:
        log_local(f"测试")

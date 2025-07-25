from playwright.sync_api import Playwright, Page
import os
import time
from utils import log, read_bank_config

def run_zhongxin_bank(playwright: Playwright, project_root, download_path, projects_accounts, kaishiriqi, jieshuriqi,
                      log_callback=None):
    """执行中信银行流水、回单导出及对账单打印"""
    def log_local(msg):
        log(msg, project_root, log_callback)

    log_local("启动中信银行导出流程...")
    log_local(f"Playwright内核路径: {os.environ.get('PLAYWRIGHT_BROWSERS_PATH')}")
    try:
        username, password, login_url, config_path = read_bank_config(project_root, "zhongxin_bank")
        log_local(f"加载配置文件: {config_path}")
    except Exception as e:
        log_local(f"config.json 文件加载失败: {str(e)}")
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

        page.locator("#new_header").get_by_text("登录").click()

        log_local("输入用户名和密码...")
        page.get_by_role("textbox", name="手机号").click()
        page.get_by_role("textbox", name="手机号").fill(username)
        page.locator("#PwdIdBoxUkeyChrome_login #noUkeyPwd_str_login").click()
        page.locator("input[type=\"password\"]").click()
        page.locator("input[type=\"password\"]").fill(password)
        page.locator("input[type=\"password\"]").press("Enter")

        log_local("等待账户管理页面加载...")
        page.wait_for_selector('text=会员中心', timeout=30000)
        page.get_by_text("会员中心").click()
        page.get_by_role("link", name="托管业务 ").click()
        page.get_by_role("link", name="托管账户 ").click()
        page.get_by_role("link", name="托管账户明细查询").click()
        page.wait_for_load_state("networkidle", timeout=30000)
        log_local("已进入托管账户明细查询页面")

        # 设置日期范围并查询
        log_local("设置开始日期和结束日期...")
        try:
            page.locator('input[name="startDate"]').evaluate(
                f'element => {{ element.value = "{kaishiriqi}"; element.dispatchEvent(new Event("input", {{ bubbles: true }})); element.dispatchEvent(new Event("change", {{ bubbles: true }})); }}'
            )
            log_local(f"已设置开始日期: {kaishiriqi}")
            page.locator('input[name="endDate"]').evaluate(
                f'element => {{ element.value = "{jieshuriqi}"; element.dispatchEvent(new Event("input", {{ bubbles: true }})); element.dispatchEvent(new Event("change", {{ bubbles: true }})); }}'
            )
            log_local(f"已设置结束日期: {jieshuriqi}")
        except Exception as e:
            log_local(f"设置日期失败: {str(e)}")
            raise

        # 验证日期输入
        try:
            start_value = page.locator('input[name="startDate"]').input_value()
            if start_value != kaishiriqi:
                log_local(f"开始日期验证失败: 期望 {kaishiriqi}, 实际 {start_value}")
                raise Exception("开始日期验证失败")
            log_local(f"开始日期验证成功: {start_value}")
            end_value = page.locator('input[name="endDate"]').input_value()
            if end_value != jieshuriqi:
                log_local(f"结束日期验证失败: 期望 {jieshuriqi}, 实际 {end_value}")
                raise Exception("结束日期验证失败")
            log_local(f"结束日期验证成功: {end_value}")
        except Exception as e:
            log_local(f"日期验证失败: {str(e)}")
            raise

        # 点击查询按钮
        try:
            page.get_by_text("查询", exact=True).click()
            page.wait_for_load_state("networkidle", timeout=30000)
            log_local("查询按钮点击成功")
        except Exception as e:
            log_local(f"查询按钮点击失败: {str(e)}")
            raise

        log_local("日期设置和查询完成，暂停以便手动检查")
        input("按 Enter 继续或 Ctrl+C 退出...")

    except Exception as e:
        log_local(f"中信银行导出流程异常: {str(e)}")
    finally:
        browser.close()
        log_local("浏览器已关闭")
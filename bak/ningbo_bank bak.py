
import os
import sys
import configparser
from playwright.sync_api import Playwright, sync_playwright

def read_bank_config():
    config_path = os.path.join(os.path.dirname(__file__), "config.txt")
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")

    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')

    if "ningbo_bank" not in config:
        raise KeyError("配置文件中缺少 [ningbo_bank] 配置项")

    bank_conf = config["ningbo_bank"]

    username = bank_conf.get("username")
    password = bank_conf.get("password")
    login_url = bank_conf.get("login_url")

    if not all([username, password, login_url]):
        raise ValueError("ningbo_bank 配置不完整，请检查 config.txt")

    return username, password, login_url

def run(playwright: Playwright, base_path, xiangmu, kaishiriqi, jieshuriqi):
    username, password, login_url = read_bank_config()
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")

    try:
        duizhang_path = os.path.join(base_path, xiangmu, "银行流水")
        huidan_path = os.path.join(base_path, xiangmu, "银行回单")
        os.makedirs(duizhang_path, exist_ok=True)
        os.makedirs(huidan_path, exist_ok=True)
        os.makedirs(os.path.dirname(log_path), exist_ok=True)

        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(login_url)

        # 登录
        page.get_by_role("textbox", name="用户名").fill(username)
        page.get_by_role("textbox", name="请输入您的密码").fill(password)

        # 等待人工输入验证码 & 登录后页面加载
        page.wait_for_selector('text=账户管理', timeout=60000)

        # 导出对账单逻辑（同你现有）
        page.get_by_role("link", name="账户管理").click()
        page.get_by_role("link", name="账户明细").click()

        page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
        page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").fill(xiangmu)
        page.get_by_role("listitem").filter(has_text=xiangmu).locator("span").nth(2).click()

        page.get_by_text("展开").first.click()
        page.get_by_role("textbox", name="开始日期").fill(kaishiriqi)
        page.get_by_role("textbox", name="开始日期").press("Enter")
        page.get_by_role("textbox", name="结束日期").fill(jieshuriqi)
        page.get_by_role("textbox", name="结束日期").press("Enter")

        page.get_by_role("button", name=" 查询").click()
        page.get_by_role("checkbox", name="Toggle Selection of All Rows").check()
        page.get_by_role("button", name="导出 ").click()

        with page.expect_download() as download_info:
            page.get_by_text("对账单导出", exact=True).click()
        download = download_info.value
        filename = f"{xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
        download.save_as(os.path.join(duizhang_path, filename))

        page.get_by_role("checkbox", name="Toggle Selection of All Rows").uncheck()
        page.get_by_role("checkbox", name="Toggle Selection of All Rows").check()

        page.get_by_role("button", name="导出 ").click()

        found = False
        for i in range(1, 10000):
            menu_selector = f"#dropdown-menu-{i:04d}"
            try:
                menu_item = page.locator(f"{menu_selector}").get_by_text("凭证导出", exact=True)
                if menu_item.is_visible():
                    with page.expect_download() as download_info:
                        menu_item.click()
                    download = download_info.value
                    filename = f"{xiangmu}_银行回单_{kaishiriqi}_{jieshuriqi}.pdf"
                    download.save_as(os.path.join(huidan_path, filename))
                    found = True
                    break
            except:
                continue

        if not found:
            raise Exception("未找到‘凭证导出’菜单项，请确认菜单是否成功弹出。")

        os.startfile(os.path.join(base_path, xiangmu))

        context.close()
        browser.close()

    except Exception as e:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"导出失败：{xiangmu} {kaishiriqi}~{jieshuriqi} 错误：{str(e)}\n")

def main():
    if len(sys.argv) != 5:
        print("用法: python ningbo_bank.py <项目名称> <开始日期> <结束日期> <基础路径>")
        sys.exit(1)

    xiangmu = sys.argv[1]
    kaishiriqi = sys.argv[2]
    jieshuriqi = sys.argv[3]
    base_path = sys.argv[4]

    with sync_playwright() as playwright:
        run(playwright, base_path, xiangmu, kaishiriqi, jieshuriqi)

if __name__ == "__main__":
    main()

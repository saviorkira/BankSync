import os
from playwright.sync_api import Playwright, sync_playwright

# === 参数区 ===
base_path = r"D:\Desktop"
xiangmu = "重信·北极22026·华睿精选9号集合资金信托计划"
kaishiriqi = "2025-03-01"
jieshuriqi = "2025-03-31"

# 构建导出路径
duizhang_path = os.path.join(base_path, xiangmu, "银行流水")
huidan_path = os.path.join(base_path, xiangmu, "银行回单")
log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")

os.makedirs(duizhang_path, exist_ok=True)
os.makedirs(huidan_path, exist_ok=True)
os.makedirs(os.path.dirname(log_path), exist_ok=True)


def run(playwright: Playwright) -> None:
    try:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto("https://www.e-custody.com/#/login")

        # 登录（人工输入验证码）
        page.get_by_role("textbox", name="用户名").fill("18323580933")
        page.get_by_role("textbox", name="请输入您的密码").fill("2780zjj?")

        # 等待人工输入验证码 & 登录后页面加载
        page.wait_for_selector('text=账户管理', timeout=60000)

        # 导出对账单
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






        # 点击“导出”按钮弹出菜单
        page.get_by_role("button", name="导出 ").click()

        # 遍历 menu ID，尝试点击“凭证导出”
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
                continue  # 如果该ID不存在或不可见，跳过继续

        if not found:
            raise Exception("未找到‘凭证导出’菜单项，请确认菜单是否成功弹出。")




        # 打开文件夹
        os.startfile(os.path.join(base_path, xiangmu))

        context.close()
        browser.close()

    except Exception as e:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"导出失败：{xiangmu} {kaishiriqi}~{jieshuriqi} 错误：{str(e)}\n")


with sync_playwright() as playwright:
    run(playwright)

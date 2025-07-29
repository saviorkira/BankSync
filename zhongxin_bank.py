from playwright.sync_api import Playwright, Page
import os
import time
from utils import log, read_bank_config, find_and_click_image, find_image, get_resource_path

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


        for index, (xiangmuid, xiangmu, account) in enumerate(projects_accounts):
            log_local(f"处理产品：{xiangmuid}_{xiangmu}，托管账户：{account}")
            try:
                # 创建文件夹
                folder_name = f"{xiangmuid}_{xiangmu}"
                duizhang_path = os.path.join(download_path, folder_name, "银行流水")
                huidan_path = os.path.join(download_path, folder_name, "银行回单")
                duizhangdan_path = os.path.join(download_path, folder_name, "银行对账单")
                os.makedirs(duizhang_path, exist_ok=True)
                os.makedirs(huidan_path, exist_ok=True)
                os.makedirs(duizhangdan_path, exist_ok=True)

                page.get_by_role("link", name="托管业务 ").click()
                page.get_by_role("link", name="托管账户 ").click()
                page.get_by_role("link", name="托管账户明细查询").click()
                page.wait_for_load_state("networkidle", timeout=30000)
                log_local("已进入托管账户明细查询页面")

                # 设置查询项目
                log_local(f"设置查询项目: {xiangmu}")
                time.sleep(2)
                page.locator("#inputPro").click()
                # 提取“·”后6个字符
                project_name = xiangmu.split('·')[1][:6] if '·' in xiangmu else xiangmu
                page.locator("#inputPro").fill(project_name)
                time.sleep(1)
                page.locator("#inputPro").press("Enter")
                time.sleep(1)
                # page.wait_for_selector("#ulPro", timeout=10000)
                page.locator("#ulPro").click()
                time.sleep(2)
                log_local(f"已选择项目: {project_name}")

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

                # 点击查询按钮
                log_local("执行查询...")
                page.get_by_text("查询", exact=True).click()
                page.wait_for_load_state("networkidle", timeout=30000)
                log_local("查询完成，等待页面加载...")
                time.sleep(1)

                # 检查是否有流水
                wuliushui_template = get_resource_path("zhongxin_wuliushui.bmp", project_root)
                position = find_image(
                    template_path=wuliushui_template,
                    base_path=project_root,
                    threshold=0.8,
                    max_attempts=5
                )
                if position:
                    log_local(f"检测到无流水图片: {wuliushui_template}，跳过当前产品")
                    continue  # 跳到下一个产品

                # 导出流水
                log_local("开始导出流水...")
                page.get_by_text("导出").click()
                page.get_by_role("button", name="确认").click()
                time.sleep(1)
                page.get_by_role("link", name="下载中心 ").click()
                page.get_by_role("link", name="异步下载").click()
                log_local("进入下载中心，检查文件处理状态...")

                # 检查文件处理状态
                wenjianchulizhong_template = get_resource_path("zhongxin_wenjianchulizhong.bmp", project_root)
                while True:
                    position = find_image(
                        template_path=wenjianchulizhong_template,
                        base_path=project_root,
                        threshold=0.8,
                        max_attempts=5
                    )
                    if position:
                        log_local("检测到文件处理中，点击查询按钮...")
                        page.get_by_role("button", name="查询").click()
                        time.sleep(2)  # 等待2秒
                    else:
                        log_local("文件处理完成或未检测到处理中状态，继续下载...")
                        break

                # 使用识图点击下载按钮并捕获下载
                template_path = get_resource_path("zhongxin_xiazai.bmp", project_root)
                with page.expect_download() as download_info:
                    log_local("等待文件下载...")
                    position = find_and_click_image(
                        template_path=template_path,
                        base_path=project_root,
                        offset_y=8,  # 向下偏移4像素
                        threshold=0.8,  # 提高阈值以确保准确性
                        max_attempts=10
                    )
                    if not position:
                        log_local(f"识图失败，未找到下载按钮: {template_path}")
                        raise Exception("识图失败，未找到下载按钮")
                    time.sleep(2)  # 确保点击后下载触发
                download = download_info.value
                filename = f"{xiangmuid}_{xiangmu}_银行流水_{kaishiriqi}_{jieshuriqi}.xlsx"
                download.save_as(os.path.join(duizhang_path, filename))
                log_local(f"银行流水导出完成：{filename}")
                time.sleep(0.5)

            except Exception as e:
                log_local(f"处理产品 {xiangmuid}_{xiangmu} 失败: {str(e)}")
                continue  # 继续处理下一个产品

    except Exception as e:
        log_local(f"中信银行导出流程异常: {str(e)}")
        raise
    finally:
        browser.close()
        log_local("浏览器已关闭")
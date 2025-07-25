page.get_by_role("textbox", name="请输入用户名").click()
page.get_by_role("textbox", name="请输入用户名").fill("80064231302")
page.get_by_role("textbox", name="请输入登录密码").click()
page.get_by_role("textbox", name="请输入登录密码").press("Unidentified")
page.get_by_role("button", name="登录").click()

page.get_by_text("托管业务").click()
page.get_by_role("listitem", name="账户明细查询").click()
# page.get_by_role("textbox", name="开始日期").click()
page.get_by_role("textbox", name="开始日期").fill("2025-06-01")
page.get_by_role("textbox", name="开始日期").press("Enter")
# page.get_by_role("textbox", name="结束日期").click()
page.get_by_role("textbox", name="结束日期").fill("2025-07-23")
page.get_by_role("textbox", name="结束日期").press("Enter")


#导出银行流水
page.get_by_role("textbox", name="请选择产品").first.dblclick()
page.locator(".appmain-content").click()
page.get_by_role("textbox", name="请选择产品").first.click()
page.get_by_role("textbox", name="请选择产品").first.fill("江宁经开")
page.get_by_role("textbox", name="请选择产品").first.press("Enter")
page.get_by_role("textbox", name="请选择产品").first.press("ArrowDown")
page.get_by_role("textbox", name="请选择产品").first.press("Enter")
page.get_by_role("button", name="查询").click()
with page.expect_download() as download_info:
    page.get_by_text("导出", exact=True).click()
download = download_info.value


#导出对账单



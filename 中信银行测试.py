page.get_by_role("link", name="托管业务 ").click()
    page.get_by_role("link", name="托管账户 ").click()
    page.get_by_role("link", name="托管账户回单查询").click()
    page.locator("#inputPro").click()
    page.locator("#inputPro").fill("简州空港")
    page.locator("#inputPro").press("Enter")
    page.locator("#ulPro").click()


    # page.get_by_role("listitem").filter(has_text="账户信息： 8111201012500701674").locator("span").click()
    # page.get_by_role("checkbox", name="8111201012500701674").check()
    # page.get_by_role("listitem").filter(has_text="账户信息： 8111201012500701674").locator("span").click()

page.get_by_role("listitem").filter(has_text=f"账户信息： {account}").locator("span").click()
page.get_by_role("checkbox", name=account).check()
page.get_by_role("listitem").filter(has_text=f"账户信息： {account}").locator("span").click()



    page.locator("input[name=\"hisDay\"]").check()

page.locator('input[name="startDate"]').evaluate(
    f'element => {{ element.value = "{kaishiriqi}"; element.dispatchEvent(new Event("input", {{ bubbles: true }})); element.dispatchEvent(new Event("change", {{ bubbles: true }})); }}'
)

page.locator('input[name="endDate"]').evaluate(
    f'element => {{ element.value = "{jieshuriqi}"; element.dispatchEvent(new Event("input", {{ bubbles: true }})); element.dispatchEvent(new Event("change", {{ bubbles: true }})); }}'
)

    page.get_by_role("button", name="批量下载").click()
    page.get_by_role("button", name="确认").click()
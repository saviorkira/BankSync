page.get_by_role("button", name="打印 ").click()

menus.nth(0).locator("text=对账单导出", exact=True).click()

menus.nth(0).locator("text=对账单导出").click()
menus.nth(0).locator("li").filter(has_text="对账单导出").first.click()


page.get_by_text("对账单导出", exact=True)

menus.nth(0).get_by_text("对账单导出", exact=True).click()

menus = page.locator("css=[id^='dropdown-menu-']")
count = menus.count()
found = False
for i in range(count):
    item = menus.nth(i).filter(has_text="对账单导出")
    if item.is_visible():
        item.click()
        log_local(f"点击了对账单打印菜单项")
        found = True
        break
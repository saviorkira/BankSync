import os
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox
from playwright.sync_api import Playwright, sync_playwright

# 这里是你的导出目录和日期设置，可以改成界面输入（演示写死）
base_path = r"D:\Desktop"
kaishiriqi = "2025-03-01"
jieshuriqi = "2025-03-31"
log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
os.makedirs(os.path.dirname(log_path), exist_ok=True)

def run(playwright: Playwright, xiangmu: str, kaishi: str, jieshu: str):
    duizhang_path = os.path.join(base_path, xiangmu, "银行流水")
    huidan_path = os.path.join(base_path, xiangmu, "银行回单")
    os.makedirs(duizhang_path, exist_ok=True)
    os.makedirs(huidan_path, exist_ok=True)

    try:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto("https://www.e-custody.com/#/login")

        # 登录（人工输入验证码）
        page.get_by_role("textbox", name="用户名").fill("18323580933")
        page.get_by_role("textbox", name="请输入您的密码").fill("2780zjj?")
        page.wait_for_selector('text=账户管理', timeout=60000)

        # 导出对账单
        page.get_by_role("link", name="账户管理").click()
        page.get_by_role("link", name="账户明细").click()

        page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").click()
        page.get_by_role("textbox", name="输入项目名称或项目对应账号关键字进行查询").fill(xiangmu)
        page.get_by_role("listitem").filter(has_text=xiangmu).locator("span").nth(2).click()

        page.get_by_text("展开").first.click()
        page.get_by_role("textbox", name="开始日期").fill(kaishi)
        page.get_by_role("textbox", name="开始日期").press("Enter")
        page.get_by_role("textbox", name="结束日期").fill(jieshu)
        page.get_by_role("textbox", name="结束日期").press("Enter")

        page.get_by_role("button", name=" 查询").click()
        page.get_by_role("checkbox", name="Toggle Selection of All Rows").check()
        page.get_by_role("button", name="导出 ").click()

        with page.expect_download() as download_info:
            page.get_by_text("对账单导出", exact=True).click()
        download = download_info.value
        filename = f"{xiangmu}_银行流水_{kaishi}_{jieshu}.xlsx"
        download.save_as(os.path.join(duizhang_path, filename))

        page.get_by_role("checkbox", name="Toggle Selection of All Rows").uncheck()
        page.get_by_role("checkbox", name="Toggle Selection of All Rows").check()

        # 点击“导出”按钮弹出菜单
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
                    filename = f"{xiangmu}_银行回单_{kaishi}_{jieshu}.pdf"
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
            f.write(f"导出失败：{xiangmu} {kaishi}~{jieshu} 错误：{str(e)}\n")
        raise e

def start_export(xiangmu_list, kaishi, jieshu, log_text, btn_start):
    def task():
        btn_start.config(state="disabled")
        with sync_playwright() as playwright:
            for xiangmu in xiangmu_list:
                log_text.insert(tk.END, f"开始导出项目: {xiangmu}\n")
                log_text.see(tk.END)
                try:
                    run(playwright, xiangmu, kaishi, jieshu)
                    log_text.insert(tk.END, f"项目 {xiangmu} 导出完成\n")
                except Exception as e:
                    log_text.insert(tk.END, f"项目 {xiangmu} 导出失败: {e}\n")
                log_text.see(tk.END)
        btn_start.config(state="normal")
        messagebox.showinfo("完成", "所有项目导出完成！")

    threading.Thread(target=task, daemon=True).start()

def main():
    root = tk.Tk()
    root.title("银行流水与回单导出工具")

    tk.Label(root, text="请输入项目名称（多项目换行）：").pack(anchor="w", padx=10, pady=5)
    txt_projects = scrolledtext.ScrolledText(root, width=50, height=10)
    txt_projects.pack(padx=10, pady=5)

    tk.Label(root, text="开始日期 (YYYY-MM-DD)：").pack(anchor="w", padx=10)
    entry_start = tk.Entry(root)
    entry_start.pack(padx=10, pady=5)
    entry_start.insert(0, kaishiriqi)

    tk.Label(root, text="结束日期 (YYYY-MM-DD)：").pack(anchor="w", padx=10)
    entry_end = tk.Entry(root)
    entry_end.pack(padx=10, pady=5)
    entry_end.insert(0, jieshuriqi)

    btn_start = tk.Button(root, text="开始导出",
                          command=lambda: start_export(
                              [p.strip() for p in txt_projects.get("1.0", tk.END).splitlines() if p.strip()],
                              entry_start.get().strip(),
                              entry_end.get().strip(),
                              log_text,
                              btn_start))
    btn_start.pack(pady=10)

    tk.Label(root, text="日志：").pack(anchor="w", padx=10)
    log_text = scrolledtext.ScrolledText(root, width=60, height=15, state='normal')
    log_text.pack(padx=10, pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()

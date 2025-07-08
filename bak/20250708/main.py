import os
import sys
import json
import pandas as pd
from playwright.sync_api import sync_playwright
from datetime import datetime
from ningbo_bank import run_ningbo_bank
from utils import log, read_bank_config, get_resource_path

os.environ["PYTHONIOENCODING"] = "utf-8"

def force_check_expiration_local(project_root, expire_date_str="2026-06-01"):
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期为 {expire_date_str}）。", project_root)
        sys.exit(1)

def main():
    project_root = os.path.abspath(os.path.dirname(__file__))
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(project_root, "playwright-browsers")
    log(f"设置 PLAYWRIGHT_BROWSERS_PATH: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}", project_root)

    if len(sys.argv) != 5:
        log(f"参数错误: 需提供 download_path, excel_path, start_date, end_date", project_root)
        sys.exit(1)

    download_path, excel_path, start_date, end_date = sys.argv[1:5]

    try:
        df = pd.read_excel(excel_path, header=0)
        if df.empty or len(df.columns) < 2:
            log("错误: Excel 文件格式错误，至少需要两列（项目名称和银行账号）", download_path)
            sys.exit(1)
        excel_data = [(str(project).strip(), str(account).strip()) for project, account in zip(df.iloc[:, 0], df.iloc[:, 1]) if str(project).strip() and str(account).strip()]
        log(f"成功导入 Excel 文件：{excel_path}", download_path)

        with sync_playwright() as playwright:
            run_ningbo_bank(playwright, project_root, download_path, excel_data, start_date, end_date, log_callback=lambda msg: print(json.dumps({"log": msg})))
        print(json.dumps({"status": "success"}))
    except Exception as e:
        print(json.dumps({"status": "error", "message": str(e)}))
        sys.exit(1)

if __name__ == "__main__":
    force_check_expiration_local(os.path.abspath(os.path.dirname(__file__)))
    main()
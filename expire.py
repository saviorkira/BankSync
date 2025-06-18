import sys
from datetime import datetime

def force_check_expiration_local(expire_date_str="2025-06-30"):
    """使用本地系统时间判断是否过期，过期则直接退出，不弹窗"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()

    if current_date > expire_date:
        print(f"程序已过期（截止日期为 {expire_date_str}），自动退出。")
        sys.exit(0)


if __name__ == "__main__":
    force_check_expiration_local("2025-06-01")  # 改成今天前一天测试弹窗退出

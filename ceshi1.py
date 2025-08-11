import sys
import ntplib
from datetime import datetime


def check_expiration_with_ntp(
        ntp_servers=["ntp.ntsc.ac.cn", "cn.pool.ntp.org", "time.edu.cn", "ntp.aliyun.com"],
        expire_date_str="2026-06-01"
):
    ntp_time = None
    for server in ntp_servers:
        try:
            client = ntplib.NTPClient()
            response = client.request(server, timeout=2)  # 设置 2 秒超时
            ntp_time = datetime.fromtimestamp(response.tx_time)
            print(f"Time retrieved from {server}: {ntp_time}")
            break
        except Exception as e:
            print(f"Failed to connect to {server}: {e}")

    if ntp_time is None:
        print("All NTP servers are unreachable, falling back to local time.")
        ntp_time = datetime.now()  # 回退到本地时间

    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    if ntp_time > expire_date:
        print(f"Program has expired! Current time: {ntp_time}")
        sys.exit(1)
    else:
        print(f"Program is valid. Current time: {ntp_time}")


if __name__ == "__main__":
    check_expiration_with_ntp()
import threading
import time
import webview
from app import app

# 启动 Flask 后端线程
def start_server():
    app.run(debug=False, use_reloader=False)

# JS 可以调用的窗口控制 API
class Api:
    def close(self):
        webview.windows[0].destroy()

    def minimize(self):
        webview.windows[0].minimize()

    def maximize(self):
        webview.windows[0].toggle_fullscreen()

if __name__ == '__main__':
    # 启动 Flask 服务
    threading.Thread(target=start_server, daemon=True).start()
    time.sleep(1)

    # 创建窗口，并注册 JS 可调用 API
    window = webview.create_window(
        title="BankSync",
        url="http://127.0.0.1:5000",
        width=900,
        height=600,
        resizable=True,
        frameless=True,
        js_api=Api()  # ✅ 正确注册方式：放在这里！
    )

    # 启动桌面窗口
    webview.start()

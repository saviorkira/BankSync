import ctypes
import flet as ft

def main(page: ft.Page):
    page.title = "自定义任务栏信息"
    page.window_icon = "icons/app.ico"

    # 修改任务栏 AppID（避免显示 flet）
    hwnd = ctypes.windll.user32.FindWindowW(None, page.title)
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MyCustom.App")

    page.add(ft.Text("任务栏右键菜单的描述已经被覆盖"))

ft.app(target=main)

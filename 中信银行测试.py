import flet as ft
import time

def main(page: ft.Page):
    page.title = "Flet 启动动画示例"
    page.window_width = 400
    page.window_height = 400

    # 启动时的 Lottie 动画
    loading = ft.Lottie(
        src="https://assets9.lottiefiles.com/packages/lf20_touohxv0.json",
        width=200,
        height=200,
        repeat=True,
        animate=True,
    )

    # 显示启动页
    page.add(
        ft.Column(
            [
                loading,
                ft.Text("启动中...", size=20, weight="bold")
            ],
            alignment="center",
            horizontal_alignment="center",
            expand=True
        )
    )
    page.update()

    # 模拟加载 3 秒
    time.sleep(3)

    # 清除启动页，进入主界面
    page.clean()
    page.add(ft.Text("欢迎进入主界面！", size=25, color="green"))

ft.app(target=main)

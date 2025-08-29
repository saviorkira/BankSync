import flet as ft

def main(page: ft.Page):
    page.title = "Flet + Lottie Demo"
    page.window_width = 400
    page.window_height = 400

    # 加载 Lottie 动画
    # 这里可以用本地文件路径，也可以用网络 URL
    anim = ft.Lottie(
        src="https://assets9.lottiefiles.com/packages/lf20_touohxv0.json",
        width=200,
        height=200,
        repeat=True,       # 循环播放
        reverse=False,     # 是否反向播放
        animate=True,      # 自动播放
    )

    def toggle_play(e):
        anim.animate = not anim.animate
        anim.update()

    page.add(
        anim,
        ft.ElevatedButton("Play / Pause", on_click=toggle_play)
    )

ft.app(target=main)

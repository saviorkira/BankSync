import flet as ft

def main(page: ft.Page):
    page.title = "自定义窗口大小"
    page.window.bgcolor = ft.Colors.TRANSPARENT
    page.window.title_bar_hidden = True
    page.window.frameless = True
    # 设置窗口位置和大小
    page.window.left = 100
    page.window.top = 100
    page.window.width = 400
    page.window.height = 600

    page.add(
        ft.WindowDragArea(
            content=ft.Container(
                content=ft.ElevatedButton("可拖动的窗口"),
                bgcolor=ft.Colors.BLUE_200,
                padding=10,
            ),
            width=800,
            height=600
        )
    )

ft.app(target=main)
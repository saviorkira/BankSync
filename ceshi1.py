import flet as ft

def main(page: ft.Page):
    container = ft.AnimatedSwitcher(
        width=100,
        height=100,
        # bgcolor=ft.Colors.RED,
        duration=500,
        animate_opacity=300,
    )

    def animate(e):
        container.width = 200
        container.height = 200
        # container.bgcolor = ft.Colors.BLUE
        container.opacity = 0.5
        page.update()

    page.add(
        container,
        ft.ElevatedButton("Animate", on_click=animate)
    )

ft.app(target=main)

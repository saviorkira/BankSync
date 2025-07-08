import flet as ft

def main(page: ft.Page):
    page.window.title_bar_hidden = True
    page.window.title_bar_buttons_hidden = True
    page.padding = 0  # 移除页面默认边距

    page.add(
        ft.Row(
            [
                ft.WindowDragArea(
                    ft.Container(
                        ft.Text(
                            "Drag this area to move, maximize and restore application window.",
                            size=12,
                        ),
                        bgcolor=ft.Colors.AMBER_300,
                        padding=ft.padding.only(left=10, top=5, right=5, bottom=5),  # 仅左侧保留少量padding
                        margin=0,  # Container移除边距
                        height=30,
                    ),
                    expand=True,
                ),
                ft.IconButton(
                    ft.Icons.CLOSE,
                    on_click=lambda _: page.window.close(),
                    icon_size=16,
                    padding=5,
                    width=30,
                    height=30,
                ),
            ],
            height=30,
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=0,  # 移除子控件间距
        )
    )

ft.app(main)
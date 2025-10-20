import os
import flet as ft
from datetime import datetime

def create_version_ui(project_root, update_log):
    """创建版本页面 UI，读取并显示 README.md 内容"""
    readme_path = os.path.join(project_root, "README.md")
    content = "未找到 README.md 文件，请检查项目根目录。"
    try:
        with open(readme_path, 'r', encoding='utf-8') as f:
            content = f.read()
        update_log(f"成功读取 README.md: {readme_path}")
    except UnicodeDecodeError:
        try:
            with open(readme_path, 'r', encoding='gbk') as f:
                content = f.read()
            update_log(f"成功读取 README.md（GBK 编码）: {readme_path}")
        except Exception as e:
            content = f"读取 README.md 失败（GBK 编码）：{str(e)}"
            update_log(content)
    except FileNotFoundError:
        content = "未找到 README.md 文件。"
        update_log(content)
    except Exception as e:
        content = f"读取 README.md 失败：{str(e)}"
        update_log(content)

    return ft.Container(
        content=ft.Column(
            [
                # ft.Text("版本信息", font_family="sansr", size=16, weight=ft.FontWeight.BOLD),
                ft.Markdown(
                    content,
                    selectable=True,
                    extension_set=ft.MarkdownExtensionSet.GITHUB_FLAVORED,
                    code_theme="monokai",
                    # code_style=ft.TextStyle(font_family="sansr", size=12),
                    # text_style=ft.TextStyle(font_family="sansr", size=14),
                    expand=True,
                ),
            ],
            spacing=10,
            scroll=ft.ScrollMode.AUTO,
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
        ),
        padding=5,
        # border_radius=8,
        bgcolor=ft.Colors.WHITE,
        shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
        width=300,  # 初始宽度，toggle_window_size 会动态更新
        expand=True,
    )
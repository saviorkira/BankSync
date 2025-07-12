import os
import json
import flet as ft
from flet import Column, Row, TextField, Checkbox, IconButton, FloatingActionButton, Tabs, Tab, OutlinedButton
from utils import log

class Task(ft.Column):
    def __init__(self, task_name, task_status_change, task_delete, completed=False):
        super().__init__()
        self.completed = completed
        self.task_name = task_name
        self.task_status_change = task_status_change
        self.task_delete = task_delete
        self.display_task = ft.Checkbox(
            value=self.completed,
            label=self.task_name,
            on_change=self.status_changed,
            label_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        )
        self.edit_name = ft.TextField(
            expand=1,
            text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        )

        self.display_view = ft.Row(
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                self.display_task,
                ft.Row(
                    spacing=0,
                    controls=[
                        ft.IconButton(
                            icon=ft.Icons.CREATE_OUTLINED,
                            tooltip="编辑待办",
                            on_click=self.edit_clicked,
                        ),
                        ft.IconButton(
                            ft.Icons.DELETE_OUTLINE,
                            tooltip="删除待办",
                            on_click=self.delete_clicked,
                        ),
                    ],
                ),
            ],
        )

        self.edit_view = ft.Row(
            visible=False,
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                self.edit_name,
                ft.IconButton(
                    icon=ft.Icons.DONE_OUTLINE_OUTLINED,
                    icon_color=ft.Colors.GREEN,
                    tooltip="更新待办",
                    on_click=self.save_clicked,
                ),
            ],
        )
        self.controls = [self.display_view, self.edit_view]

    def edit_clicked(self, e):
        self.edit_name.value = self.display_task.label
        self.display_view.visible = False
        self.edit_view.visible = True
        self.safe_update()

    def save_clicked(self, e):
        self.display_task.label = self.edit_name.value
        self.display_view.visible = True
        self.edit_view.visible = False
        self.safe_update()

    def status_changed(self, e):
        self.completed = self.display_task.value
        self.task_status_change(self)

    def delete_clicked(self, e):
        self.task_delete(self)

    def safe_update(self):
        """仅在控件已添加到页面时调用 update()"""
        if hasattr(self, 'page') and self.page is not None:
            self.update()
            log(f"Task 更新: {self.task_name}, 已添加到页面", os.path.abspath(os.path.dirname(__file__)))
        else:
            log(f"Task 未更新: {self.task_name}, 未添加到页面", os.path.abspath(os.path.dirname(__file__)))

class TodoApp(ft.Column):
    def __init__(self, project_root):
        super().__init__()
        self.project_root = project_root
        self.new_task = ft.TextField(
            hint_text="需要做什么？",
            on_submit=self.add_clicked,
            expand=True,
            text_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
            label_style=ft.TextStyle(font_family="FZLanTingHei", size=14),
        )
        self.tasks = ft.Column()
        self.filter = ft.Tabs(
            scrollable=False,
            selected_index=0,
            on_change=self.tabs_changed,
            tabs=[ft.Tab(text="全部"), ft.Tab(text="未完成"), ft.Tab(text="已完成")],
        )
        self.items_left = ft.Text("0 项未完成", style=ft.TextStyle(font_family="FZLanTingHei", size=14))
        self.controls = [
            ft.Row(
                controls=[
                    self.new_task,
                    ft.FloatingActionButton(
                        icon=ft.Icons.ADD,
                        on_click=self.add_clicked,
                        bgcolor=ft.Colors.BLUE_700,
                        tooltip="添加待办",
                    ),
                ],
            ),
            ft.Column(
                spacing=25,
                controls=[
                    self.filter,
                    self.tasks,
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            self.items_left,
                            ft.OutlinedButton(
                                text="清除已完成",
                                on_click=self.clear_clicked,
                                style=ft.ButtonStyle(
                                    text_style=ft.TextStyle(font_family="FZLanTingHei", size=14)
                                ),
                            ),
                        ],
                    ),
                ],
            ),
        ]
        self.load_tasks()
        log(f"TodoApp 初始化完成，加载任务数: {len(self.tasks.controls)}", self.project_root)

    def load_tasks(self):
        config_path = os.path.join(self.project_root, "config.json")
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for task in data.get("tasks", []):
                        self.tasks.controls.append(
                            Task(
                                task_name=task["name"],
                                task_status_change=self.task_status_change,
                                task_delete=self.task_delete,
                                completed=task["completed"]
                            )
                        )
                log(f"成功加载 config.json，任务数: {len(self.tasks.controls)}", self.project_root)
        except Exception as e:
            log(f"加载待办任务失败：{str(e)}", self.project_root)

    def save_tasks(self):
        config_path = os.path.join(self.project_root, "config.json")
        try:
            # 读取现有配置文件，保留其他部分（如 ningbo_bank）
            config = {}
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            # 更新 tasks 部分
            tasks = [
                {"name": task.display_task.label, "completed": task.completed}
                for task in self.tasks.controls
            ]
            config["tasks"] = tasks
            # 写入更新后的配置
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            log(f"成功保存任务到 config.json，任务数: {len(tasks)}", self.project_root)
        except Exception as e:
            log(f"保存待办任务失败：{str(e)}", self.project_root)

    def add_clicked(self, e):
        if self.new_task.value:
            task = Task(self.new_task.value, self.task_status_change, self.task_delete)
            self.tasks.controls.append(task)
            self.new_task.value = ""
            self.new_task.focus()
            self.save_tasks()
            self.safe_update()
            log(f"添加任务: {task.task_name}", self.project_root)

    def task_status_change(self, task):
        self.save_tasks()
        self.safe_update()
        log(f"任务状态变更: {task.task_name}, 完成状态: {task.completed}", self.project_root)

    def task_delete(self, task):
        self.tasks.controls.remove(task)
        self.save_tasks()
        self.safe_update()
        log(f"删除任务: {task.task_name}", self.project_root)

    def tabs_changed(self, e):
        self.safe_update()
        log(f"切换待办标签: {self.filter.tabs[self.filter.selected_index].text}", self.project_root)

    def clear_clicked(self, e):
        for task in self.tasks.controls[:]:
            if task.completed:
                self.task_delete(task)
        self.save_tasks()
        log("清除已完成任务", self.project_root)

    def before_update(self):
        status = self.filter.tabs[self.filter.selected_index].text
        count = 0
        for task in self.tasks.controls:
            task.visible = (
                status == "全部"
                or (status == "未完成" and not task.completed)
                or (status == "已完成" and task.completed)
            )
            if not task.completed:
                count += 1
        self.items_left.value = f"{count} 项未完成"

    def safe_update(self):
        """仅在控件已添加到页面时调用 update()"""
        if hasattr(self, 'page') and self.page is not None:
            self.update()
            log("TodoApp 更新: 已添加到页面", self.project_root)
        else:
            log("TodoApp 未更新: 未添加到页面", self.project_root)
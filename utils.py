import os
import time
import json  # 替换 configparser
import numpy as np
import cv2
import pyautogui
from pywinauto import Desktop, Application
from datetime import datetime


def log(message, project_root, log_callback=None):
    """记录日志到文件和回调函数"""
    print(message)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_message = f"{timestamp}: {message}\n"
    log_dir = os.path.join(project_root, "data", "log")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "导出日志.txt")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(log_message)
    if log_callback:
        log_callback(message)

def read_bank_config(project_root, site_name):
    """读取指定银行的配置文件"""
    config_path = get_resource_path("config.json", project_root, subfolder="")  # 调整为根目录
    log(f"尝试加载配置文件: {config_path} for {site_name}", project_root)
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        if site_name not in config:
            raise KeyError(f"配置文件中缺少 '{site_name}' 配置项")
        bank_conf = config[site_name]
        username = bank_conf.get("username")
        password = bank_conf.get("password")
        login_url = bank_conf.get("login_url")
        if not all([username, password, login_url]):
            raise ValueError(f"{site_name} 配置不完整，请检查 config.json")
        return username, password, login_url, config_path
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON 配置文件解析失败: {str(e)}")

def get_resource_path(relative_path, project_root, subfolder="data/cv2"):
    """获取资源文件的绝对路径，始终从项目根目录加载"""
    if relative_path == "config.json":  # config.json 在项目根目录
        full_path = os.path.join(project_root, relative_path)
    else:
        full_path = os.path.join(project_root, subfolder, relative_path)
    log(f"检查路径: {full_path}, 存在: {os.path.exists(full_path)}", project_root)
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    raise FileNotFoundError(f"资源文件不存在: {full_path}")

def find_and_click_image(template_path, base_path, offset_x=0, offset_y=0, threshold=0.5, max_attempts=10):
    """使用模板匹配找到图像并点击"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        if not os.path.exists(template_path):
            log(f"模板路径不存在: {template_path}", base_path)
            return None
        try:
            template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception as e:
            log(f"模板加载异常: {template_path}, 错误: {e}", base_path)
            return None
        if template is None:
            log(f"无法加载模板图像: {template_path}", base_path)
            return None
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            x, y = max_loc
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)
            log(f"第{attempt+1}次尝试成功，点击位置: ({x + offset_x}, {y + offset_y})", base_path)
            return (x, y)
        time.sleep(1)
    log(f"未找到模板: {template_path}，尝试次数: {max_attempts}", base_path)
    return None

def handle_overwrite_dialog(base_path):
    """处理文件覆盖对话框"""
    try:
        app = Desktop(backend="win32")
        dialog = app.window(title_re=".*文件已存在.*|.*确认保存.*|.*确认另存为.*|.*Confirm Save.*|.*Replace.*")
        dialog.wait("exists ready", timeout=5)
        dialog.set_focus()
        replace_btn = dialog.child_window(title_re="是.*|替换.*|Yes.*|Replace.*", class_name="Button")
        replace_btn.wait("exists enabled visible ready", timeout=3)
        replace_btn.click()
        time.sleep(1)
    except Exception as e:
        log(f"未检测到覆盖确认窗口或点击失败: {e}", base_path)

def handle_save_dialog(save_path, pdf_filename, base_path):
    """处理保存对话框"""
    full_path = os.path.join(save_path, pdf_filename)
    log(f"尝试捕捉‘另存为’窗口，目标路径: {full_path}", base_path)
    try:
        desktop = Desktop(backend="win32")
        dialogs = desktop.windows(title_re="^另存为$")
        if not dialogs:
            raise Exception("未找到标题为 '另存为' 的窗口")
        for i, dlg_wrapper in enumerate(dialogs):
            try:
                handle = dlg_wrapper.handle
                app = Application(backend="win32").connect(handle=handle)
                dlg = app.window(handle=handle)
                dlg.set_focus()
                time.sleep(0.3)
                edit = dlg.child_window(class_name="Edit")
                edit.set_focus()
                edit.set_edit_text(full_path)
                log(f"窗口{i + 1}设置路径成功: {full_path}", base_path)
                time.sleep(0.5)
                save_btn = dlg.child_window(class_name="Button", title_re="保存|Save")
                save_btn.click()
                log(f"点击保存按钮完成", base_path)
                time.sleep(3)
                handle_overwrite_dialog(base_path)
                return
            except Exception as inner_e:
                log(f"窗口{i + 1}处理失败: {inner_e}", base_path)
                continue
        raise Exception("未找到可用的‘另存为’窗口")
    except Exception as e:
        log(f"快速保存失败，错误: {e}", base_path)
        raise
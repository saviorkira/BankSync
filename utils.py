import os
import sys
import configparser
import numpy as np
import cv2
import time
from datetime import datetime
from pywinauto import Desktop
from pywinauto.application import Application
import pyautogui

def force_check_expiration_local(expire_date_str="2026-06-01"):
    """使用本地系统时间判断是否过期，过期则退出"""
    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    current_date = datetime.now()
    if current_date > expire_date:
        log(f"程序已过期（截止日期为 {expire_date_str}）。")
        sys.exit(0)

def log(msg, base_path=r"D:\Desktop", log_callback=None):
    """记录日志到文件和UI"""
    print(msg)
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")
    if log_callback:
        log_callback(msg)

def get_resource_path(relative_path, subfolder="seek"):
    """获取资源路径，优先检查 seek 子目录"""
    base_path = os.path.abspath(os.path.dirname(__file__))
    full_path = os.path.join(base_path, subfolder, relative_path)
    log(f"检查路径: {full_path}, 存在: {os.path.exists(full_path)}", base_path)
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    full_path = os.path.join(base_path, relative_path)
    log(f"回退路径: {full_path}, 存在: {os.path.exists(full_path)}", base_path)
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    return full_path

def read_bank_config(base_path):
    """读取 config.txt 中的宁波银行配置"""
    config_path = get_resource_path("config.txt")
    log(f"尝试加载配置文件: {config_path}", base_path)
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')
    if "ningbo_bank" not in config:
        raise KeyError("配置文件中缺少 [ningbo_bank] 配置项")
    bank_conf = config["ningbo_bank"]
    username = bank_conf.get("username")
    password = bank_conf.get("password")
    login_url = bank_conf.get("login_url")
    if not all([username, password, login_url]):
        raise ValueError("ningbo_bank 配置不完整，请检查 config.txt")
    return username, password, login_url, config_path

def find_and_click_image(template_path, offset_x=0, offset_y=0, threshold=0.5, max_attempts=10):
    """使用模板匹配找到图像并点击"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        if not os.path.exists(template_path):
            log(f"模板路径不存在: {template_path}")
            return None
        try:
            template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception as e:
            log(f"模板加载异常: {template_path}, 错误: {e}")
            return None
        if template is None:
            log(f"无法加载模板图像: {template_path}")
            return None
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            x, y = max_loc
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)
            log(f"第{attempt+1}次尝试成功，点击位置: ({x + offset_x}, {y + offset_y})")
            return (x, y)
        time.sleep(1)
    log(f"未找到模板: {template_path}，尝试次数: {max_attempts}")
    return None

def handle_overwrite_dialog():
    """处理文件覆盖确认弹窗"""
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
        log(f"未检测到覆盖确认窗口或点击失败: {e}")

def handle_save_dialog(save_path, pdf_filename):
    """处理另存为窗口，输入路径并保存"""
    full_path = os.path.join(save_path, pdf_filename)
    log(f"尝试捕捉‘另存为’窗口，目标路径: {full_path}")
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
                log(f"窗口{i + 1}设置路径成功: {full_path}")
                time.sleep(0.5)
                save_btn = dlg.child_window(class_name="Button", title_re="保存|Save")
                save_btn.click()
                log("点击保存按钮完成")
                time.sleep(3)
                handle_overwrite_dialog()
                return
            except Exception as inner_e:
                log(f"窗口{i + 1}处理失败: {inner_e}")
                continue
        raise Exception("未找到可用的‘另存为’窗口")
    except Exception as e:
        log(f"快速保存失败，错误: {e}")
        raise
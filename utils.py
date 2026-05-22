import os
import time
import json
import numpy as np
import cv2
import pyautogui
from pywinauto import Desktop, Application
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import ntplib

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
    """读取指定网站的配置文件"""
    config_path = get_resource_path("config.json", project_root, subfolder="")
    log(f"尝试加载配置文件: {config_path} for {site_name}", project_root)
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        if site_name not in config:
            raise KeyError(f"配置文件中缺少 '{site_name}' 配置项")
        site_conf = config[site_name]
        username = site_conf.get("username", "")  # 允许 username 为空
        password = site_conf.get("password", "")  # 允许 password 为空
        login_url = site_conf.get("login_url")
        if not login_url:
            raise ValueError(f"{site_name} 配置不完整，缺少 login_url")
        return username, password, login_url, config_path
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON 配置文件解析失败: {str(e)}")

def read_email_config(project_root, site_name="email_263"):
    """读取邮箱配置文件"""
    config_path = get_resource_path("config.json", project_root, subfolder="")
    log(f"尝试加载邮箱配置文件: {config_path} for {site_name}", project_root)
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        if site_name not in config:
            raise KeyError(f"配置文件中缺少 '{site_name}' 配置项")
        site_conf = config[site_name]
        username = site_conf.get("username", "")
        password = site_conf.get("password", "")
        imap_server = site_conf.get("imap_server", "imap.263.net")
        port = site_conf.get("port", 993)
        if not (username and password and imap_server and port):
            raise ValueError(f"{site_name} 配置不完整，缺少 username, password, imap_server 或 port")
        return username, password, imap_server, port
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

def find_image(template_path, project_root, threshold=0.5, max_attempts=10):
    """仅查找图像，不执行点击操作"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        if not os.path.exists(template_path):
            log(f"模板路径不存在: {template_path}", project_root)
            return None
        try:
            template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception as e:
            log(f"模板加载异常: {template_path}, 错误: {e}", project_root)
            return None
        if template is None:
            log(f"无法加载模板图像: {template_path}", project_root)
            return None
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            x, y = max_loc
            log(f"第{attempt+1}次尝试成功，找到图像位置: ({x}, {y})", project_root)
            return (x, y)
        time.sleep(1)
    log(f"未找到模板: {template_path}，尝试次数: {max_attempts}", project_root)
    return None

def find_and_click_image(template_path, project_root, offset_x=0, offset_y=0, threshold=0.5, max_attempts=10):
    """使用模板匹配找到图像并点击"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        if not os.path.exists(template_path):
            log(f"模板路径不存在: {template_path}", project_root)
            return None
        try:
            template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception as e:
            log(f"模板加载异常: {template_path}, 错误: {e}", project_root)
            return None
        if template is None:
            log(f"无法加载模板图像: {template_path}", project_root)
            return None
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= threshold:
            x, y = max_loc
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)
            log(f"第{attempt+1}次尝试成功，点击位置: ({x + offset_x}, {y + offset_y})", project_root)
            return (x, y)
        time.sleep(1)
    log(f"未找到模板: {template_path}，尝试次数: {max_attempts}", project_root)
    return None

def find_all_images(template_path, project_root, threshold=0.8, max_attempts=5):
    """找到所有匹配图像的位置，按 y 坐标从上到下排序"""
    template = cv2.imdecode(np.fromfile(template_path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if template is None:
        log(f"无法加载模板图像: {template_path}", project_root)
        return [], 0, 0

    w, h = template.shape[1], template.shape[0]
    locations = []

    for _ in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        yloc, xloc = np.where(result >= threshold)

        points = []
        for (x, y) in zip(xloc, yloc):
            points.append((x, y))

        # 过滤重复点（距离 < 10 像素视为同一）
        unique_points = []
        for pt in points:
            if not unique_points or all(abs(pt[0] - up[0]) > 10 or abs(pt[1] - up[1]) > 10 for up in unique_points):
                unique_points.append(pt)

        if unique_points:
            # 按 y 排序（从上到下）
            unique_points.sort(key=lambda p: p[1])
            log(f"找到 {len(unique_points)} 个匹配位置: {template_path}", project_root)
            return unique_points, w, h

        time.sleep(1)

    log(f"未找到任何匹配模板: {template_path}，尝试次数: {max_attempts}", project_root)
    return [], 0, 0

def handle_overwrite_dialog(project_root):
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
        log(f"未检测到覆盖确认窗口或点击失败: {e}", project_root)

def handle_save_dialog(save_path, pdf_filename, project_root):
    """处理保存对话框"""
    full_path = os.path.join(save_path, pdf_filename)
    log(f"尝试捕捉‘另存为’窗口，目标路径: {full_path}", project_root)
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
                log(f"窗口{i + 1}设置路径成功: {full_path}", project_root)
                time.sleep(0.5)
                save_btn = dlg.child_window(class_name="Button", title_re="保存|Save")
                save_btn.click()
                log(f"点击保存按钮完成", project_root)
                time.sleep(3)
                handle_overwrite_dialog(project_root)
                return
            except Exception as inner_e:
                log(f"窗口{i + 1}处理失败: {inner_e}", project_root)
                continue
        raise Exception("未找到可用的‘另存为’窗口")
    except Exception as e:
        log(f"快速保存失败，错误: {e}", project_root)
        raise

def check_expiration_with_ntp(project_root, ntp_servers=["ntp.ntsc.ac.cn", "cn.pool.ntp.org", "time.edu.cn", "ntp.aliyun.com"], expire_date_str="2026-12-31"):
    """检查程序是否过期，使用 NTP 服务器获取时间"""
    ntp_time = None
    for server in ntp_servers:
        try:
            client = ntplib.NTPClient()
            response = client.request(server, timeout=2)
            ntp_time = datetime.fromtimestamp(response.tx_time)
            log(f"从 {server} 获取时间成功: {ntp_time}", project_root)
            break
        except Exception as e:
            log(f"无法连接到 {server}: {str(e)}", project_root)

    if ntp_time is None:
        log("所有 NTP 服务器均不可用，回退到本地时间", project_root)
        ntp_time = datetime.now()

    expire_date = datetime.strptime(expire_date_str, "%Y-%m-%d")
    if ntp_time > expire_date:
        # log(f"程序已过期！当前时间: {ntp_time}，过期时间: {expire_date}", project_root)
        sys.exit(1)
    # else:
        # log(f"程序未过期，当前时间: {ntp_time}，过期时间: {expire_date}", project_root)

def split_date_ranges(start_date_str, end_date_str, max_months=3):
    """
    将日期范围拆分为多个不超过 max_months 个月的子范围。

    参数:
        start_date_str (str): 开始日期，格式为 'YYYY-MM-DD'
        end_date_str (str): 结束日期，格式为 'YYYY-MM-DD'
        max_months (int): 最大月份跨度，默认为3个月

    返回:
        List[Tuple[str, str]]: 包含多个 (start_date, end_date) 的列表，日期格式为 'YYYY-MM-DD'
    """
    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
    except ValueError as e:
        raise ValueError(f"日期格式错误，应为 YYYY-MM-DD: {str(e)}")

    if start_date > end_date:
        raise ValueError("开始日期不能晚于结束日期")

    date_ranges = []
    current_start = start_date

    while current_start <= end_date:
        # 计算当前范围的结束日期：开始日期 + max_months 月 - 1 天
        current_end = min(
            current_start + relativedelta(months=max_months) - timedelta(days=1),
            end_date
        )
        # 添加当前范围
        date_ranges.append((
            current_start.strftime("%Y-%m-%d"),
            current_end.strftime("%Y-%m-%d")
        ))
        # 更新下一次循环的开始日期
        current_start = current_end + timedelta(days=1)

    return date_ranges
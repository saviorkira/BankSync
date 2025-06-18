import os
import time
import pyautogui
import cv2
import numpy as np
from datetime import datetime
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

# 全局日志函数
def log(msg, base_path=r"D:\Desktop"):
    print(msg)
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")

def get_resource_path(relative_path, subfolder="seek"):
    """获取资源路径，优先检查 seek 子目录"""
    base_path = os.path.abspath(os.path.dirname(__file__))
    full_path = os.path.join(base_path, subfolder, relative_path)
    log(f"检查路径: {full_path}, 存在: {os.path.exists(full_path)}, 大小: {os.path.getsize(full_path) if os.path.exists(full_path) else 0} bytes")
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    full_path = os.path.join(base_path, relative_path)
    log(f"回退路径: {full_path}, 存在: {os.path.exists(full_path)}, 大小: {os.path.getsize(full_path) if os.path.exists(full_path) else 0} bytes")
    if os.path.exists(full_path) and os.path.getsize(full_path) > 0:
        return full_path
    return full_path

def find_and_click_image(template_path, offset_x=0, offset_y=0, threshold=0.5, max_attempts=10):
    """使用模板匹配找到图像并点击，添加偏移，降低阈值以容忍差异"""
    for attempt in range(max_attempts):
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        # 使用 imdecode 读取模板图像（支持中文路径）
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
            with open(template_path, 'rb') as f:
                log(f"文件内容前10字节: {f.read(10).hex()}")
            return None

        log(f"模板尺寸: {template.shape}")

        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        log(f"匹配度: {max_val}, 位置: {max_loc}")

        if max_val >= threshold:
            x, y = max_loc
            pyautogui.click(x + offset_x + template.shape[1] // 2, y + offset_y + template.shape[0] // 2)
            log(f"第{attempt+1}次尝试成功，点击位置: ({x + offset_x}, {y + offset_y})")
            return (x, y)

        time.sleep(1)

    log(f"未找到模板: {template_path}，尝试次数: {max_attempts}")
    return None

def handle_overwrite_dialog():
    """处理文件已存在的覆盖确认弹窗，自动点击‘是’或‘替换’按钮"""
    from pywinauto import Desktop
    try:
        app = Desktop(backend="win32")
        dialog = app.window(title_re=".*文件已存在.*|.*确认保存.*|.*确认另存为.*|.*Confirm Save.*|.*Replace.*")
        dialog.wait("exists ready", timeout=5)
        dialog.set_focus()
        print(f"覆盖确认窗口标题: {dialog.window_text()}")  # 调试输出

        replace_btn = dialog.child_window(title_re="是.*|替换.*|Yes.*|Replace.*", class_name="Button")
        replace_btn.wait("exists enabled visible ready", timeout=3)
        print(f"找到覆盖按钮，标题: {replace_btn.element_info.name}")  # 调试输出
        replace_btn.click()
        print("点击覆盖按钮成功")
        time.sleep(1)
    except Exception as e:
        print(f"未检测到覆盖确认窗口或点击失败: {e}")


def handle_save_dialog(huidan_path, pdf_filename):
    """处理另存为窗口，输入完整路径并保存"""
    full_path = os.path.join(huidan_path, pdf_filename)
    log("尝试捕捉‘另存为’窗口...")

    try:
        app = Desktop(backend="win32")
        dialog = app.window(title_re=".*另存为.*|.*保存.*|.*Save As.*")
        dialog.wait("exists ready", timeout=10)
        dialog.set_focus()
        log(f"成功捕获另存为窗口: {dialog.window_text()}")

        edit = dialog.child_window(class_name="Edit")
        edit.set_focus()
        edit.set_edit_text(full_path)
        log(f"设置完整路径成功: {full_path}")

        time.sleep(0.5)
        save_btn = dialog.child_window(class_name="Button", title_re="保存|Save")
        save_btn.click()
        log("点击保存按钮完成")
        time.sleep(3)

        # 自动处理覆盖弹窗
        handle_overwrite_dialog()

    except Exception as e:
        log(f"快速保存失败，错误: {e}")
        raise

def debug_print_window():
    xiangmu = "重信·北极22026·华睿精选9号集合资金信托计划"
    base_path = r"D:\Desktop"
    huidan_path = os.path.join(base_path, xiangmu, "银行回单")
    os.makedirs(huidan_path, exist_ok=True)

    # 这里你要准备好以下模板图像放在 seek 文件夹中
    target_printer_path = get_resource_path("target_printer.bmp")
    printer_dropdown_path = get_resource_path("printer_dropdown.bmp")
    save_as_pdf_default_path = get_resource_path("save_as_pdf_default.bmp")
    save_as_pdf_hover_path = get_resource_path("save_as_pdf_hover.bmp")
    save_button_path = get_resource_path("save_button.bmp")

    if not (os.path.exists(target_printer_path) and os.path.exists(printer_dropdown_path) and
            os.path.exists(save_as_pdf_default_path) and os.path.exists(save_as_pdf_hover_path) and
            os.path.exists(save_button_path)):
        log(f"模板图像缺失或大小为0: target_printer.bmp, printer_dropdown.bmp, save_as_pdf_default.bmp, save_as_pdf_hover.bmp 或 save_button.bmp")
        raise FileNotFoundError("请准备相关模板图像并放入 seek 文件夹，确保文件大小大于0")

    # 定位“目标打印机”文字
    log("定位‘目标打印机’位置...")
    target_pos = find_and_click_image(target_printer_path)
    if target_pos:
        x_target, y_target = target_pos
        log(f"找到‘目标打印机’文字，位置: ({x_target}, {y_target})")
    else:
        log("未找到‘目标打印机’文字")
        pyautogui.screenshot(f"error_target_printer_{xiangmu}.png")
        raise Exception("未找到打印窗口中的‘目标打印机’文字")

    # 点击“目标打印机”右侧偏移 250 像素
    log("点击‘目标打印机’右侧偏移 250 像素模拟下拉按钮...")
    x_offset = 250
    pyautogui.click(x_target + x_offset, y_target)
    log(f"模拟点击偏移位置: ({x_target + x_offset}, {y_target})")
    time.sleep(2)

    # 尝试匹配并点击“另存为 PDF”
    log("尝试匹配并点击‘另存为 PDF’选项...")
    pdf_clicked = False
    for attempt in range(3):
        pyautogui.moveTo(x_target + x_offset, y_target + 20)
        time.sleep(0.5)
        if find_and_click_image(save_as_pdf_default_path):
            log("成功点击默认状态的‘另存为 PDF’按钮")
            pdf_clicked = True
            time.sleep(2)
            break
        elif find_and_click_image(save_as_pdf_hover_path):
            log("成功点击悬停状态的‘另存为 PDF’按钮")
            pdf_clicked = True
            time.sleep(2)
            break
        log(f"第{attempt+1}次尝试未找到‘另存为 PDF’，重试...")
        time.sleep(1)

    if not pdf_clicked:
        log("未找到‘另存为 PDF’按钮（默认或悬停状态），请检查截图")
        pyautogui.screenshot(f"error_save_as_pdf_{xiangmu}.png")
        raise Exception("未找到打印窗口中的‘另存为 PDF’按钮")

    # 点击“保存”按钮
    log("尝试点击‘保存’按钮...")
    if find_and_click_image(save_button_path):
        log("成功点击‘保存’按钮")
        time.sleep(2)
    else:
        log("未找到‘保存’按钮，请检查截图")
        pyautogui.screenshot(f"error_save_button_{xiangmu}.png")
        raise Exception("未找到打印窗口中的‘保存’按钮")

    # 处理另存为窗口，写入文件名并保存
    pdf_filename = f"{xiangmu}_对账单打印_{kaishiriqi}_{jieshuriqi}.pdf"
    handle_save_dialog(huidan_path, pdf_filename)

    pdf_path = os.path.join(huidan_path, pdf_filename)
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
        log(f"对账单打印 PDF 完成：{pdf_filename}")
    else:
        log(f"PDF 文件生成失败或为空：{pdf_path}")
        raise Exception("PDF 文件生成失败或为空")

if __name__ == "__main__":
    kaishiriqi = "2025-03-21"
    jieshuriqi = "2025-03-21"
    log("开始调试打印窗口...")
    try:
        debug_print_window()
    except Exception as e:
        log(f"调试失败：{str(e)}")

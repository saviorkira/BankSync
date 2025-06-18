import os
import time
import pyautogui
import cv2
import numpy as np
from datetime import datetime

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

def find_and_click_image(template_path, offset_x=0, offset_y=0, threshold=0.6, max_attempts=10):
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

def debug_print_window():
    xiangmu = "重信·北极22026·华睿精选9号集合资金信托计划"
    base_path = r"D:\Desktop"
    huidan_path = os.path.join(base_path, xiangmu, "银行回单")
    os.makedirs(huidan_path, exist_ok=True)

    # 加载模板文件 (更新为 .bmp)
    target_printer_path = get_resource_path("target_printer.bmp")
    printer_dropdown_path = get_resource_path("printer_dropdown.bmp")
    save_as_pdf_default_path = get_resource_path("save_as_pdf_default.bmp")
    save_as_pdf_expanded_path = get_resource_path("save_as_pdf_expanded.bmp")
    if not (os.path.exists(target_printer_path) and os.path.exists(printer_dropdown_path) and 
            os.path.exists(save_as_pdf_default_path) and os.path.exists(save_as_pdf_expanded_path)):
        log(f"模板图像缺失或大小为0: target_printer.bmp, printer_dropdown.bmp, save_as_pdf_default.bmp 或 save_as_pdf_expanded.bmp 在 seek 文件夹中")
        raise FileNotFoundError("请准备 target_printer.bmp, printer_dropdown.bmp, save_as_pdf_default.bmp 和 save_as_pdf_expanded.bmp 模板图像并放入 seek 文件夹，确保文件大小大于0")

    # 定位“目标打印机”文字
    log("开始定位‘目标打印机’文字...")
    target_pos = find_and_click_image(target_printer_path)
    if target_pos:
        x_target, y_target = target_pos
        log(f"找到‘目标打印机’文字，位置: ({x_target}, {y_target})")
    else:
        log("未找到‘目标打印机’文字")
        pyautogui.screenshot(f"error_target_printer_{xiangmu}.png")
        raise Exception("未找到打印窗口中的‘目标打印机’文字")

    # 点击“目标打印机”右侧下拉箭头
    dropdown_offset_x = 35  # 调整此值以匹配箭头位置
    log("开始点击‘目标打印机’右侧下拉箭头...")
    if find_and_click_image(printer_dropdown_path, offset_x=dropdown_offset_x, offset_y=0):
        log(f"成功点击‘目标打印机’右侧下拉菜单按钮")
        time.sleep(2)  # 等待下拉菜单出现
    else:
        log("未找到‘目标打印机’右侧下拉菜单按钮")
        pyautogui.screenshot(f"error_printer_dropdown_{xiangmu}.png")
        raise Exception("未找到打印窗口中的‘目标打印机’右侧下拉菜单按钮")

    # 尝试匹配默认状态的“另存为 PDF”
    log("尝试匹配默认状态的‘另存为 PDF’...")
    if find_and_click_image(save_as_pdf_default_path):
        log("成功点击默认状态的‘另存为 PDF’按钮")
        time.sleep(2)  # 等待保存对话框
        pyautogui.write(huidan_path)
        pyautogui.press("enter")
        log(f"保存 PDF 到: {huidan_path}")
        time.sleep(5)  # 等待保存完成
    else:
        # 若默认状态未匹配，匹配展开状态
        log("默认状态未匹配，尝试展开后匹配‘另存为 PDF’...")
        if find_and_click_image(save_as_pdf_expanded_path):
            log("成功点击展开状态的‘另存为 PDF’按钮")
            time.sleep(2)  # 等待保存对话框
            pyautogui.write(huidan_path)
            pyautogui.press("enter")
            log(f"保存 PDF 到: {huidan_path}")
            time.sleep(5)  # 等待保存完成
        else:
            log("未找到‘另存为 PDF’按钮（默认或展开状态），请检查截图")
            pyautogui.screenshot(f"error_save_as_pdf_{xiangmu}.png")
            raise Exception("未找到打印窗口中的‘另存为 PDF’按钮，请确认 save_as_pdf_default.bmp 和 save_as_pdf_expanded.bmp")

    pdf_filename = f"{xiangmu}_对账单打印_{kaishiriqi}_{jieshuriqi}.pdf"
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
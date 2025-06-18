import os
import time
from datetime import datetime
from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys

def log(msg, base_path=r"D:\Desktop"):
    print(msg)
    log_path = os.path.join(base_path, "导出日志", "导出错误日志.txt")
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")

def wait_for_print_dialog(timeout=10):
    log("等待打印窗口弹出...")
    for _ in range(timeout):
        try:
            dlg = Desktop(backend="uia").window(title_re=".*打印.*", control_type="Window")
            if dlg.exists():
                log("已检测到打印窗口")
                return dlg
        except Exception:
            pass
        time.sleep(1)
    raise TimeoutError("未在规定时间内找到打印窗口")

def select_printer_and_save(dlg, printer_name="另存为 PDF", save_path=None):
    dlg.set_focus()
    time.sleep(1)

    try:
        log(f"尝试选择打印机: {printer_name}")
        # 打开打印机下拉菜单（有些窗口不需要）
        printer_dropdown = dlg.child_window(control_type="ComboBox")
        printer_dropdown.select(printer_name)
    except Exception as e:
        log(f"选择打印机失败，尝试使用键盘操作: {e}")
        send_keys("{TAB 2}{DOWN 5}")  # 模拟选择“另存为 PDF”

    time.sleep(1)

    log("点击打印按钮")
    try:
        print_btn = dlg.child_window(title="打印", control_type="Button")
        print_btn.click_input()
    except:
        log("找不到打印按钮，尝试模拟回车")
        send_keys("{ENTER}")

    # 等待另存为对话框
    log("等待另存为窗口...")
    for _ in range(10):
        try:
            save_dlg = Desktop(backend="uia").window(title_re=".*另存为.*", control_type="Window")
            if save_dlg.exists():
                log("另存为窗口已出现")
                break
        except:
            pass
        time.sleep(1)
    else:
        raise Exception("另存为窗口未出现")

    save_dlg.set_focus()
    time.sleep(1)

    if save_path:
        log(f"输入保存路径: {save_path}")
        send_keys(save_path, with_spaces=True)
        time.sleep(1)
        send_keys("{ENTER}")
        time.sleep(2)

def debug_print_window():
    xiangmu = "重信·北极22026·华睿精选9号集合资金信托计划"
    base_path = r"D:\Desktop"
    huidan_path = os.path.join(base_path, xiangmu, "银行回单")
    os.makedirs(huidan_path, exist_ok=True)

    pdf_filename = f"{xiangmu}_对账单打印_2025-03-21_2025-03-21.pdf"
    full_pdf_path = os.path.join(huidan_path, pdf_filename)

    dlg = wait_for_print_dialog()
    select_printer_and_save(dlg, save_path=full_pdf_path)

    if os.path.exists(full_pdf_path) and os.path.getsize(full_pdf_path) > 0:
        log(f"对账单 PDF 完成: {pdf_filename}")
    else:
        log(f"PDF 文件生成失败或为空: {full_pdf_path}")
        raise Exception("PDF 文件生成失败或为空")

if __name__ == "__main__":
    log("开始调试打印窗口（pywinauto 版本）...")
    try:
        debug_print_window()
    except Exception as e:
        log(f"调试失败：{str(e)}")
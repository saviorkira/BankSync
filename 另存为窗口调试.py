from pywinauto import Desktop
import time

def debug_save_dialog_structure():
    print("正在查找窗口：另存为 / 保存 / Save As ...")

    try:
        app = Desktop(backend="win32")
        all_dialogs = app.windows(title_re="^另存为$")
        print(f"\n共找到匹配窗口数：{len(all_dialogs)}\n")

        if not all_dialogs:
            print("未找到匹配窗口，请确认另存为窗口已弹出。")
            return

        for i, dlg in enumerate(all_dialogs):
            try:
                # 尝试获取底层wrapper对象
                try:
                    wrapper = dlg.wrapper_object()
                except Exception:
                    wrapper = dlg

                print(f"\n【窗口{i + 1}】标题: {wrapper.window_text()}")
                wrapper.set_focus()
                time.sleep(0.5)
                print("控件结构如下:\n")
                wrapper.print_control_identifiers(depth=4)
            except Exception as e:
                print(f"窗口{i + 1} 处理失败: {e}")
    except Exception as e:
        print(f"整体操作失败: {e}")

if __name__ == "__main__":
    debug_save_dialog_structure()

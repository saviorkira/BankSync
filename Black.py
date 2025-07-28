import time
import threading
import keyboard
import win32gui
import win32con
import signal

screen_off_flag = True

def turn_off_screen():
    try:
        while screen_off_flag:
            print("尝试关闭屏幕...")
            win32gui.PostMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, 2)
            time.sleep(1.5)
    except Exception as e:
        print(f"关闭屏幕时出错: {e}")

def listen_combo():
    global screen_off_flag
    try:
        print("等待 Ctrl+Alt+L 组合键...")
        keyboard.wait("ctrl+alt+l")
        screen_off_flag = False
        print("组合键触发，停止强制黑屏")
        win32gui.PostMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, -1)  # 唤醒屏幕
    except Exception as e:
        print(f"监听组合键时出错: {e}")

# 捕获终止信号以确保亮屏
def handle_exit(signum, frame):
    global screen_off_flag
    screen_off_flag = False
    print("程序被终止，唤醒屏幕")
    win32gui.PostMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, -1)
    exit(0)

signal.signal(signal.SIGTERM, handle_exit)
signal.signal(signal.SIGINT, handle_exit)

thread1 = threading.Thread(target=turn_off_screen, daemon=True)
thread2 = threading.Thread(target=listen_combo, daemon=True)

thread1.start()
thread2.start()

try:
    while thread1.is_alive() or thread2.is_alive():
        time.sleep(0.1)
except KeyboardInterrupt:
    handle_exit(None, None)
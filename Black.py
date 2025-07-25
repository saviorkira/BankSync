import time, threading
import keyboard
import ctypes

screen_off_flag = True

def turn_off_screen():
    while screen_off_flag:
        ctypes.windll.user32.SendMessageW(-1, 0x0112, 0xF170, 2)
        time.sleep(1.5)  # 每1.5秒强制熄屏一次

def listen_combo():
    global screen_off_flag
    keyboard.wait("ctrl+alt+l")
    screen_off_flag = False
    print("组合键触发，停止强制黑屏")

# 启动两个线程
threading.Thread(target=turn_off_screen).start()
threading.Thread(target=listen_combo).start()

import os
import shutil
import subprocess
import sys


def clean_build_dirs():
    project_root = os.path.abspath(os.path.dirname(__file__))
    for directory in [os.path.join(project_root, "build"), os.path.join(project_root, "dist", "BankSync-debug")]:
        if os.path.exists(directory):
            shutil.rmtree(directory, ignore_errors=True)


def copy_required_files():
    project_root = os.path.abspath(os.path.dirname(__file__))
    dist_dir = os.path.join(project_root, "dist", "BankSync-debug")
    required_items = ["playwright-browsers", "data", "config.json"]
    os.makedirs(dist_dir, exist_ok=True)
    for item in required_items:
        src = os.path.join(project_root, item)
        dst = os.path.join(dist_dir, item)
        if os.path.exists(src):
            if os.path.isdir(src):
                shutil.copytree(src, dst, dirs_exist_ok=True)
            else:
                shutil.copy2(src, dst)
            print(f"已复制: {item}")
        else:
            print(f"警告: {item} 不存在")


def run_pyinstaller():
    project_root = os.path.abspath(os.path.dirname(__file__))
    icon_path = os.path.join(project_root, "data", "S.ico")

    command = [
        "pyinstaller",
        "--noconfirm",
        "--onedir",
        "--console",
        "--noupx",
        f"--icon={icon_path}",
        "--name=BankSync-debug",
        "--hidden-import", "openai",
        "--hidden-import", "pandas",
        "--hidden-import", "flet",
        "--hidden-import", "playwright",
        "--hidden-import", "pyautogui",
        "--hidden-import", "pywinauto",
        "--hidden-import", "numpy",
        "--hidden-import", "cv2",
        "gui.py"
    ]

    print("执行调试打包（控制台模式）:", " ".join(command))
    try:
        subprocess.run(command, check=True)
        print("调试打包完成！")
        copy_required_files()
        print("调试版本位于: dist/BankSync-debug/")
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    print("开始打包 BankSync 调试版本（带控制台窗口）...")
    clean_build_dirs()
    run_pyinstaller()

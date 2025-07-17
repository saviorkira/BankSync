import os
import shutil
import subprocess
import sys


def clean_build_dirs():
    """清空 build 和 dist 目录"""
    project_root = os.path.abspath(os.path.dirname(__file__))
    build_dir = os.path.join(project_root, "build")
    dist_dir = os.path.join(project_root, "dist")

    for directory in [build_dir, dist_dir]:
        if os.path.exists(directory):
            print(f"删除目录: {directory}")
            shutil.rmtree(directory, ignore_errors=True)


def copy_required_files():
    """将 playwright-browsers、data 和 config.json 复制到 dist/ 目录"""
    project_root = os.path.abspath(os.path.dirname(__file__))
    dist_dir = os.path.join(project_root, "dist")
    required_items = [
        "playwright-browsers",
        "data",
        "config.json"
    ]

    for item in required_items:
        src_path = os.path.join(project_root, item)
        dst_path = os.path.join(dist_dir, item)
        if os.path.exists(src_path):
            if os.path.isdir(src_path):
                print(f"复制目录: {src_path} -> {dst_path}")
                shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
            else:
                print(f"复制文件: {src_path} -> {dst_path}")
                shutil.copy2(src_path, dst_path)
        else:
            print(f"警告: {src_path} 不存在，无法复制")


def run_pyinstaller():
    """运行 PyInstaller 进行单文件打包"""
    project_root = os.path.abspath(os.path.dirname(__file__))
    icon_path = os.path.join(project_root, "data", "S.ico")

    # PyInstaller 命令
    command = [
        "pyinstaller",
        "--noconfirm",  # 覆盖已有输出目录
        "--onefile",  # 单文件模式
        "--windowed",  # 无控制台窗口
        f"--icon={icon_path}",  # 设置图标
        "--name=BankSync",  # 设置可执行文件名
        "--hidden-import", "openai",
        "--hidden-import", "pandas",
        "--hidden-import", "flet",
        "--hidden-import", "playwright",
        "--hidden-import", "pyautogui",
        "--hidden-import", "pywinauto",
        "--hidden-import", "numpy",
        "--hidden-import", "cv2",
        "gui.py"  # 入口文件
    ]

    print("执行 PyInstaller 打包命令:", " ".join(command))
    try:
        subprocess.run(command, check=True)
        print("打包完成！可执行文件位于: dist/BankSync.exe")
        print("复制必要的文件和文件夹到 dist/ 目录...")
        copy_required_files()
        print("所有文件已复制到 dist/ 目录！")
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        sys.exit(1)


def main():
    print("开始打包 BankSync 项目...")
    print("1. 清空 build 和 dist 目录")
    clean_build_dirs()
    print("2. 执行 PyInstaller 单文件打包")
    run_pyinstaller()


if __name__ == "__main__":
    main()
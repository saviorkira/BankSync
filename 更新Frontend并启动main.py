import subprocess
import os
import sys
import importlib.util

def check_python_dependencies():
    """检查 Python 依赖是否安装"""
    required = [
        'pandas', 'openpyxl', 'playwright', 'pyautogui',
        'opencv-python', 'pywinauto', 'pywebview'
    ]
    missing = []
    for module in required:
        if importlib.util.find_spec(module) is None:
            missing.append(module)
    if missing:
        print(f"❌ 缺少以下 Python 依赖：{missing}")
        print("请运行以下命令安装：")
        print("pip install " + " ".join(missing))
        sys.exit(1)
    print("✅ 所有 Python 依赖已安装")

def run_command(command, cwd=None, step_name=""):
    """执行命令并处理错误，使用 utf-8 编码"""
    print(f"\n📦 正在执行：{step_name} -> {command}")
    try:
        result = subprocess.run(
            command, cwd=cwd, shell=True, check=True,
            text=True, capture_output=True, encoding='utf-8', errors='replace'
        )
        print(f"✅ {step_name} 成功完成")
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)
    except subprocess.CalledProcessError as e:
        print(f"❌ {step_name} 失败，退出码：{e.returncode}")
        print(f"错误输出：{e.stderr}")
        sys.exit(1)
    except UnicodeDecodeError as e:
        print(f"❌ {step_name} 编码错误：{e}")
        print("尝试以 ignore 模式重试...")
        result = subprocess.run(
            command, cwd=cwd, shell=True, check=True,
            text=True, capture_output=True, encoding='utf-8', errors='ignore'
        )
        print(f"✅ {step_name} 成功完成（忽略编码错误）")
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)

def main():
    project_root = os.getcwd()
    frontend_path = os.path.join(project_root, "frontend")
    main_py = os.path.join(project_root, "main.py")

    # Step 0: 检查 Python 依赖
    check_python_dependencies()

    # Step 1: 检查 Playwright 浏览器
    playwright_browsers = os.path.join(project_root, "playwright-browsers")
    if not os.path.isdir(playwright_browsers):
        print("❌ 未找到 playwright-browsers 目录，尝试安装 Chromium...")
        # run_command("playwright install chromium", step_name="安装 Playwright Chromium")

    # Step 2: 构建前端
    if not os.path.isdir(frontend_path):
        print("❌ 未找到 frontend 目录，请检查路径")
        sys.exit(1)
    # run_command("npm install", cwd=frontend_path, step_name="安装前端依赖")
    run_command("npm run build", cwd=frontend_path, step_name="构建前端项目")

    # Step 3: 启动 Python 主程序
    if not os.path.isfile(main_py):
        print("❌ 未找到 main.py，请检查路径")
        sys.exit(1)
    run_command("python main.py", cwd=project_root, step_name="启动主程序")

if __name__ == "__main__":
    main()
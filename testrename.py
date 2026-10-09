import os
import shutil


def restructure(root_dir):
    """
    批量整理银行下载目录：
    1. 把每个项目文件夹下子文件夹（银行流水/银行回单/银行对账单）的文件提到项目文件夹根目录
    2. 删除空的子文件夹
    3. 项目文件夹去掉第一个下划线（如 23064_重庆信托 → 23064重庆信托）
    """
    if not os.path.isdir(root_dir):
        print(f"目录不存在: {root_dir}")
        return

    for project_name in os.listdir(root_dir):
        project_path = os.path.join(root_dir, project_name)
        if not os.path.isdir(project_path):
            continue

        print(f"处理: {project_name}")

        # 把子文件夹的文件提到项目文件夹根目录
        for sub_name in os.listdir(project_path):
            sub_path = os.path.join(project_path, sub_name)
            if not os.path.isdir(sub_path):
                continue
            for file_name in os.listdir(sub_path):
                file_path = os.path.join(sub_path, file_name)
                if not os.path.isfile(file_path):
                    continue
                target = os.path.join(project_path, file_name)
                if os.path.exists(target):
                    print(f"  文件已存在，跳过: {file_name}")
                else:
                    shutil.move(file_path, target)
                    print(f"  移动: {sub_name}/{file_name} -> {file_name}")
            # 删除空子文件夹
            if not os.listdir(sub_path):
                os.rmdir(sub_path)
                print(f"  删除空文件夹: {sub_name}")

        # 项目文件夹改名：去掉第一个 _
        if '_' in project_name:
            new_name = project_name.replace('_', '', 1)
            new_path = os.path.join(root_dir, new_name)
            if os.path.exists(new_path):
                print(f"  目标文件夹已存在，跳过改名: {new_name}")
            else:
                os.rename(project_path, new_path)
                print(f"  改名: {project_name} -> {new_name}")

    print("全部完成")


if __name__ == '__main__':
    root = input("请输入根目录路径（如 D:\\Desktop\\平安银行）: ").strip()
    if not root:
        root = r"D:\Desktop\平安银行"
    restructure(root)

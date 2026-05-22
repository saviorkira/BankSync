import imaplib
import email
import os
import ssl
import re
from datetime import datetime, timedelta
from utils import read_email_config, log
import flet as ft
import threading
from imaplib import IMAP4

def create_email_ui(page: ft.Page, project_root, is_running, update_log):
    """创建邮件下载相关的 UI 组件和事件处理，参考 TodoApp 的自适应布局"""
    # 文件类型 Checkbox，'ALL' 第一个，默认选中
    email_file_types = {
        'ALL': ft.Checkbox(label="ALL", value=True, tooltip="下载所有文件类型"),
        '.xls': ft.Checkbox(label=".xls", value=False, tooltip="下载 .xls 文件"),
        '.xlsx': ft.Checkbox(label=".xlsx", value=False, tooltip="下载 .xlsx 文件"),
        '.zip': ft.Checkbox(label=".zip", value=False, tooltip="下载 .zip 文件"),
        '.rar': ft.Checkbox(label=".rar", value=False, tooltip="下载 .rar 文件"),
        '.png': ft.Checkbox(label=".png", value=False, tooltip="下载 .png 文件"),
    }

    # Checkbox 变更事件：ALL 选中时取消其他，其他选中时取消 ALL
    def checkbox_changed(e):
        if e.control == email_file_types['ALL']:
            if e.control.value:
                for ext in email_file_types:
                    if ext != 'ALL':
                        email_file_types[ext].value = False
                        email_file_types[ext].update()
        else:
            if e.control.value:
                email_file_types['ALL'].value = False
                email_file_types['ALL'].update()

    for cb in email_file_types.values():
        cb.on_change = checkbox_changed

    email_start_date = ft.TextField(
        label="开始日期",
        value=(datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d"),
        border_radius=8,
        expand=True,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        tooltip="邮件附件下载的起始日期",
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
    )

    email_end_date = ft.TextField(
        label="结束日期",
        value=datetime.now().strftime("%Y-%m-%d"),
        border_radius=8,
        expand=True,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text="格式: YYYY-MM-DD",
        tooltip="邮件附件下载的结束日期",
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
    )

    email_subject_filter = ft.TextField(
        label="邮件标题筛选，可以为空",
        value="",
        border_radius=8,
        expand=True,
        filled=True,
        bgcolor=ft.Colors.WHITE,
        hint_text=";(；)分隔项目,*为代位符",
        tooltip="例：华享*20251017；华睿*20251021",
        text_style=ft.TextStyle(font_family="sansr", size=14),
        label_style=ft.TextStyle(font_family="sansr", size=14),
    )

    email_download_button = ft.ElevatedButton(
        text="开始下载",
        icon=ft.Icons.PLAY_CIRCLE,
        disabled=False,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.padding.symmetric(horizontal=10),
            text_style=ft.TextStyle(font_family="sansr", size=14, overflow=ft.TextOverflow.ELLIPSIS),
        ),
        tooltip="开始下载邮件附件",
        expand=True,
    )

    email_download_path_button = ft.ElevatedButton(
        text="选择保存目录",
        icon=ft.Icons.FOLDER_OPEN,
        style=ft.ButtonStyle(
            color=ft.Colors.WHITE,
            bgcolor=ft.Colors.BLUE_700,
            shape=ft.RoundedRectangleBorder(radius=8),
            padding=ft.padding.symmetric(horizontal=10),
            text_style=ft.TextStyle(font_family="sansr", size=14, overflow=ft.TextOverflow.ELLIPSIS),
        ),
        tooltip="选择保存邮件附件的目录",
        expand=True,
    )

    email_download_path = [r"D:\Desktop"]

    def select_email_download_path(e: ft.FilePickerResultEvent):
        if e.path and os.path.exists(e.path):
            email_download_path[0] = e.path
            email_download_path_button.text = f"保存到: {email_download_path[0]}"
            email_download_path_button.style.padding = ft.padding.symmetric(horizontal=10)
            update_log(f"邮件附件保存路径设置为：{email_download_path[0]}")
            page.update()

    email_dir_picker = ft.FilePicker(on_result=select_email_download_path)
    page.overlay.append(email_dir_picker)
    page.update()

    email_download_path_button.on_click = lambda _: email_dir_picker.get_directory_path()

    def run_email_downloader(e):
        nonlocal is_running
        if is_running[0]:
            update_log("提示: 下载进程正在运行，请等待")
            return
        start = email_start_date.value.strip()
        end = email_end_date.value.strip()
        if not (start and end and email_download_path[0]):
            update_log("提示: 请填写完整日期和下载路径")
            return
        try:
            start_date = datetime.strptime(start, "%Y-%m-%d")
            end_date = datetime.strptime(end, "%Y-%m-%d")
            if start_date > end_date:
                update_log("错误: 开始日期不能晚于结束日期")
                return
        except ValueError:
            update_log("错误: 日期格式不正确，应为 YYYY-MM-DD")
            return

        if email_file_types['ALL'].value:
            file_types = []
        else:
            file_types = [ext for ext, cb in email_file_types.items() if cb.value and ext != 'ALL']

        if not file_types and not email_file_types['ALL'].value:
            update_log("错误: 请至少选择一种文件类型或 ALL")
            return

        # 支持 ; 和 ； 作为分隔符
        subject_filters = [f.strip() for f in re.split(r'[;；]', email_subject_filter.value.strip()) if f.strip()]
        if not subject_filters:
            subject_filters = []  # 如果空，代表下载所有邮件
            update_log("无标题筛选，下载所有匹配日期的附件")
        else:
            update_log(f"识别到所有标题筛选关键词: {subject_filters}")

        is_running[0] = True
        email_download_button.disabled = True
        email_download_button.text = "下载中..."
        email_download_button.icon = ft.Icons.HOURGLASS_TOP
        email_download_button.update()
        update_log("正在启动全局高速下载进程...")

        def worker():
            try:
                # 💡 优化点：不再循环调用后端，而是把关键词列表 subject_filters 整体传过去
                success = download_attachments(
                    project_root,
                    start,
                    end,
                    file_types,
                    subject_filters,
                    email_download_path[0],
                    update_log
                )
                if success:
                    update_log("邮件附件高速下载任务全部完成！")
                else:
                    update_log("邮件附件下载终止，请检查日志。")
            except Exception as ex:
                update_log(f"邮件附件下载出错：{str(ex)}")
            finally:
                is_running[0] = False
                email_download_button.disabled = False
                email_download_button.text = "开始下载"
                email_download_button.icon = ft.Icons.PLAY_CIRCLE
                email_download_button.update()

        threading.Thread(target=worker, daemon=True).start()

    email_download_button.on_click = run_email_downloader

    return ft.Column(
        [
            ft.Container(
                content=ft.Column(
                    [
                        email_subject_filter,
                        ft.Row(
                            [email_start_date, email_end_date],
                            spacing=10,
                            expand=True,
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        ),
                        ft.Row(
                            [email_file_types['ALL'], email_file_types['.xls'], email_file_types['.xlsx'],
                             email_file_types['.zip'], email_file_types['.rar'], email_file_types['.png']],
                            wrap=False,
                            scroll=ft.ScrollMode.AUTO,
                            spacing=10,
                            run_spacing=5,
                            alignment=ft.MainAxisAlignment.START,
                            expand=True,
                        ),
                        email_download_path_button,
                        email_download_button,
                    ],
                    spacing=10,
                    scroll=ft.ScrollMode.AUTO,
                    alignment=ft.MainAxisAlignment.START,
                    expand=True,
                    horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                ),
                padding=5,
                # border_radius=8,
                bgcolor=ft.Colors.WHITE,
                shadow=ft.BoxShadow(blur_radius=5, color=ft.Colors.GREY_400),
                width=page.window.width - 70,
                height=page.window.height - 70,
                alignment=ft.alignment.top_left,
                expand=True,
            ),
        ],
        spacing=10,
        # scroll=ft.ScrollMode.AUTO,
        alignment=ft.MainAxisAlignment.START,
    )


def download_attachments(project_root, start_date, end_date, file_types, subject_filters, download_folder,
                         log_callback):
    try:
        username, password, imap_server, port = read_email_config(project_root, "email_263")
    except Exception as e:
        log_callback(f"加载邮箱配置失败：{e}")
        return False

    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    log_file = os.path.join(download_folder, 'download_log.txt')

    def log_to_file(message):
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now()}: {message}\n")
        log_callback(message)

    def sanitize_filename(filename):
        return re.sub(r'[<>:"/\\|?*]', '_', filename)

    def decode_filename(filename):
        if not filename:
            return None
        try:
            decoded = email.header.decode_header(filename)[0][0]
            if isinstance(decoded, bytes):
                for encoding in ('utf-8', 'gb2312', 'gbk'):
                    try:
                        return decoded.decode(encoding)
                    except UnicodeDecodeError:
                        continue
                return decoded.decode('utf-8', errors='replace')
            return decoded
        except Exception as e:
            return filename

    folder_translation = {
        'INBOX': '收件箱',
        '&XfJT0ZAB-': '已发送',
        '&XfJSIJZk-': '已删除',
        '&g0l6P3ux-': '草稿箱',
        '&XfJfUmhj-': '已发送邮件箱',
        '&U05biY1Ee6E-': '垃圾邮件'
    }

    def get_friendly_folder_name(raw_name):
        return folder_translation.get(raw_name, raw_name)

    try:
        start_dt = datetime.strptime(start_date.strip(), '%Y-%m-%d')
        end_dt = datetime.strptime(end_date.strip(), '%Y-%m-%d')
        imap_start_str = start_dt.strftime('%d-%b-%Y')
        imap_end_plus_one_str = (end_dt + timedelta(days=1)).strftime('%d-%b-%Y')
    except ValueError as e:
        log_to_file(f"日期格式错误：{e}")
        return False

    log_to_file(f"开始扫描全邮箱，时间范围: {start_dt.strftime('%Y-%m-%d')} 至 {end_dt.strftime('%Y-%m-%d')}")

    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE

    try:
        mail = imaplib.IMAP4_SSL(imap_server, port, ssl_context=context)
        mail.login(username, password)
        status, folder_list = mail.list()

        folders = []
        for f in folder_list:
            f_str = f.decode('utf-8', errors='ignore')
            match = re.search(r'"([^"]+)"\s*$', f_str)
            if match:
                folders.append(match.group(1))
            else:
                parts = f_str.split(' "/" ')
                if len(parts) > 1:
                    folders.append(parts[-1].strip())

        friendly_folders = [get_friendly_folder_name(f) for f in folders]
        log_to_file(f"识别到邮箱文件夹: {friendly_folders}")
    except Exception as e:
        log_to_file(f"IMAP 连接或登录失败：{e}")
        return False

    search_query = f'SINCE {imap_start_str} BEFORE {imap_end_plus_one_str}'
    total_downloaded = 0

    for folder in folders:
        friendly_folder_title = get_friendly_folder_name(folder)

        # 扩充过滤：跳过已发送、已删除、草稿、垃圾箱等
        if any(skip in folder.upper() for skip in
               ['&XFJT0ZAB-', '&XFJFUMHJ-', '&XFJSIJZK-', '&G0L6P3UX-', '&U05BIY1EE6E-']):
            continue

        log_to_file(f"正在切换至文件夹: 【{friendly_folder_title}】")
        try:
            mail.select(f'"{folder}"', readonly=True)
        except Exception as e:
            log_to_file(f"无法打开文件夹 【{friendly_folder_title}】，跳过。")
            continue

        try:
            status, messages = mail.search('utf-8', search_query)
            if status != 'OK' or not messages or not messages[0]:
                status, messages = mail.search(None, search_query)
            if status != 'OK' or not messages or not messages[0]:
                continue
        except Exception as e:
            continue

        msg_nums = [num for num in messages[0].split(b' ') if num and num.isdigit()]
        if not msg_nums:
            continue

        log_to_file(f"在【{friendly_folder_title}】中发现 {len(msg_nums)} 封在日期范围内的邮件，正在提取标题...")

        # ⚡ 优化点 1：批量拉取所有邮件的头部信息（仅包含主题和基本元数据，极快）
        # 将所有编号组合成类似 b"1:506" 或者 b"1,2,3..."
        range_bytes = b",".join(msg_nums)
        try:
            # 仅获取 BODY[HEADER.FIELDS (SUBJECT)] 极大地减少网络I/O
            status, header_data = mail.fetch(range_bytes, '(BODY[HEADER.FIELDS (SUBJECT)])')
            if status != 'OK':
                header_data = []
        except Exception as e:
            log_to_file(f"批量提取标题失败，降级为逐封检查: {e}")
            header_data = []

        # 解析批量获取的标题映射
        subject_map = {}
        current_num = None
        for response_part in header_data:
            if isinstance(response_part, tuple):
                # 提取邮件编号
                num_match = re.search(r'^(\d+)\s+', response_part[0].decode('utf-8', errors='ignore'))
                if num_match:
                    current_num = num_match.group(1).encode()
                    header_msg = email.message_from_bytes(response_part[1])
                    subject_map[current_num] = decode_filename(header_msg['Subject']) or '无主题'

        # 开始遍历比对
        for num in msg_nums:
            try:
                # 优先从映射中拿标题，拿不到再单独去取（确保兼容性）
                if num in subject_map:
                    subject = subject_map[num]
                else:
                    status, msg_data = mail.fetch(num, '(BODY[HEADER.FIELDS (SUBJECT)])')
                    if status == 'OK' and msg_data[0]:
                        header_msg = email.message_from_bytes(msg_data[0][1])
                        subject = decode_filename(header_msg['Subject']) or '无主题'
                    else:
                        subject = '无主题'

                matched = False
                matched_keyword = "ALL"

                if subject_filters:
                    for f_word in subject_filters:
                        pattern = re.escape(f_word).replace(r'\*', '.*')
                        try:
                            if re.search(pattern, subject, re.IGNORECASE):
                                matched = True
                                matched_keyword = f_word
                                break
                        except re.error:
                            if f_word in subject:
                                matched = True
                                matched_keyword = f_word
                                break
                    if not matched:
                        continue  # 🎯 标题不匹配，直接跳过！绝不下载整封邮件！

                log_to_file(
                    f"命中匹配 -> 文件夹:[{friendly_folder_title}] | 命中词:[{matched_keyword}] | 主题:{subject}")

                # ⚡ 优化点 2：只有标题命中后，才拉取该邮件的完整内容（RFC822）
                status, msg_data = mail.fetch(num, '(RFC822)')
                if status != 'OK' or not msg_data or not msg_data[0]:
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                # 开始处理附件
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    filename = decode_filename(part.get_filename())
                    if filename and (not file_types or any(filename.lower().endswith(ext) for ext in file_types)):
                        filename = sanitize_filename(filename)
                        filepath = os.path.join(download_folder, filename)

                        if not os.path.exists(filepath):
                            with open(filepath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            log_to_file(f"   ↳ 📥 下载附件：{filename}")
                            total_downloaded += 1
                        else:
                            log_to_file(f"   ↳ ⏭️ 附件已存在，跳过：{filename}")
            except Exception as e:
                log_to_file(f"处理邮件编号 {num.decode()} 失败：{str(e)}")
                continue

    try:
        mail.close()
        mail.logout()
    except:
        pass

    log_to_file(f"扫描完成！共成功保存了 {total_downloaded} 个附件。")
    return True
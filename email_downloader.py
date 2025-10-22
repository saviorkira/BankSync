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
            if end_date.date() == datetime.now().date():
                update_log("提示: 结束日期为今天，邮件可能尚未到达")
            elif end_date > datetime.now():
                update_log("警告: 结束日期为未来日期，可能无邮件")
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
            subject_filters = ['']
        update_log(f"分割后的标题筛选关键词: {subject_filters}")

        is_running[0] = True
        email_download_button.disabled = True
        email_download_button.text = "下载中..."
        email_download_button.icon = ft.Icons.HOURGLASS_TOP
        email_download_button.update()
        update_log("开始下载邮件附件...")

        def worker():
            try:
                for subj_filter in subject_filters:
                    if subj_filter:
                        update_log(f"处理标题筛选: {subj_filter}")
                    else:
                        update_log("无标题筛选，下载所有匹配日期的附件")
                    success = download_attachments(
                        project_root,
                        start,
                        end,
                        file_types,
                        subj_filter,
                        email_download_path[0],
                        update_log
                    )
                    if success:
                        update_log(f"邮件附件下载完成（筛选: {subj_filter or 'ALL'}）。")
                    else:
                        update_log(f"邮件附件下载失败（筛选: {subj_filter or 'ALL'}），请检查日志。")
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

def download_attachments(project_root, start_date, end_date, file_types, subject_filter, download_folder, log_callback):
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
            log_to_file(f"文件名解码失败：{e}，使用原始名")
            return filename

    try:
        start = datetime.strptime(start_date, '%Y-%m-%d')
        end = datetime.strptime(end_date, '%Y-%m-%d')
        end_plus_one = (end + timedelta(days=1)).strftime('%d-%b-%Y')
        start = start.strftime('%d-%b-%Y')
    except ValueError as e:
        log_to_file(f"日期格式错误：{e}")
        return False

    log_to_file(f"搜索 {start} 至 {end} 的邮件")

    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    try:
        mail = imaplib.IMAP4_SSL(imap_server, port, ssl_context=context)
        mail.login(username, password)
        status, folders = mail.list()
        log_to_file(f"邮箱文件夹列表: {folders}")
        try:
            mail.select('INBOX')
        except Exception as e:
            log_to_file(f"选择 INBOX 失败：{e}，尝试其他文件夹")
            for folder in [b'"[Gmail]/All Mail"', b'INBOX.', b'"inbox"']:
                try:
                    mail.select(folder)
                    log_to_file(f"成功选择文件夹: {folder.decode()}")
                    break
                except:
                    continue
            else:
                log_to_file("无法选择有效文件夹")
                mail.logout()
                return False
    except Exception as e:
        log_to_file(f"IMAP 连接或登录失败：{e}")
        return False

    # 只按日期搜索
    search_query = f'SINCE {start} BEFORE {end_plus_one}'
    try:
        status, messages = mail.search('utf-8', search_query)
        log_to_file(f"搜索结果（UTF-8）: status={status}, messages={messages}")
        if status != 'OK' or not messages or not messages[0]:
            log_to_file("UTF-8 搜索无有效结果，尝试默认字符集")
            status, messages = mail.search(None, search_query)
            log_to_file(f"搜索结果（默认字符集）: status={status}, messages={messages}")
        if status != 'OK' or not messages or not messages[0]:
            log_to_file("无符合条件的邮件，可能是日期范围无邮件或服务器限制")
            if start == end_plus_one.strip():
                log_to_file("提示: 单日搜索可能因邮件尚未到达而为空")
            mail.logout()
            return True
    except Exception as e:
        log_to_file(f"搜索邮件失败：{str(e)}")
        mail.logout()
        return False

    messages = [num for num in messages[0].split(b' ') if num and num.isdigit()]
    log_to_file(f"过滤后邮件编号: {messages}")

    if not messages:
        log_to_file("无符合条件的邮件，退出。")
        if start == end_plus_one.strip():
            log_to_file("提示: 单日搜索可能因邮件尚未到达而为空")
        mail.logout()
        return True

    log_to_file(f"找到 {len(messages)} 封邮件（{start} 至 {end}）。")

    matched_emails = 0
    for num in messages:
        try:
            status, msg_data = mail.fetch(num, '(RFC822)')
            if status != 'OK' or not msg_data or not msg_data[0]:
                log_to_file(f"获取邮件 {num.decode()} 失败：无有效数据")
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = decode_filename(msg['Subject']) or '无主题'
            # 检查标题是否包含筛选关键词
            # if subject_filter and subject_filter not in subject:
            #     log_to_file(f"跳过邮件：{subject}（不匹配筛选 {subject_filter}）")
            #     continue
            if subject_filter:
                # 支持 * 通配符 => 转为正则的 .*
                pattern = re.escape(subject_filter).replace(r'\*', '.*')
                try:
                    if not re.search(pattern, subject, re.IGNORECASE):
                        log_to_file(f"跳过邮件：{subject}（不匹配筛选 {subject_filter}）")
                        continue
                except re.error as ex:
                    log_to_file(f"正则错误：{ex}，使用普通包含匹配")
                    if subject_filter not in subject:
                        continue

            matched_emails += 1
            log_to_file(f"处理邮件：{subject}")

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
                        log_to_file(f"下载附件：{filename}")
                    else:
                        log_to_file(f"附件已存在，跳过：{filename}")
        except Exception as e:
            log_to_file(f"处理邮件 {num.decode()} 失败：{str(e)}")
            continue

    if matched_emails == 0 and subject_filter:
        log_to_file(f"无邮件标题包含关键词：{subject_filter}")

    mail.close()
    mail.logout()
    log_to_file("下载完成！")
    return True
import imaplib
import email
import os
import ssl
import re
from email.header import decode_header
from datetime import datetime, timedelta

# 配置
IMAP_SERVER = 'imap.263.net'  # 国内，海外用 imapw.263.net
PORT = 993  # SSL
USERNAME = 'makunpeng@cqiti.com'  # 完整邮箱地址
PASSWORD = 'malum1ta1!!!'  # 网页生成的授权码
DOWNLOAD_FOLDER = './attachments'  # 保存路径

# 创建下载文件夹
if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)

# 日志文件
LOG_FILE = os.path.join(DOWNLOAD_FOLDER, 'download_log.txt')


def log(message):
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"{datetime.now()}: {message}\n")
    print(message)


# 清理文件名，替换非法字符
def sanitize_filename(filename):
    # 替换 Windows 非法字符：< > : " / \ | ? *
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


# 解码文件名，尝试多种编码
def decode_filename(filename):
    if not filename:
        return None
    try:
        decoded = decode_header(filename)[0][0]
        if isinstance(decoded, bytes):
            # 尝试 UTF-8、GB2312、GBK
            for encoding in ('utf-8', 'gb2312', 'gbk'):
                try:
                    return decoded.decode(encoding)
                except UnicodeDecodeError:
                    continue
            return decoded.decode('utf-8', errors='replace')  # 替换乱码
        return decoded
    except Exception as e:
        log(f"文件名解码失败：{e}，使用原始名")
        return filename


# 计算 5 天前的日期
since_date = (datetime.now() - timedelta(days=3)).strftime('%d-%b-%Y')
log(f"搜索 {since_date} 至今的邮件")

# 连接 IMAP
context = ssl.create_default_context()
context.check_hostname = False
context.verify_mode = ssl.CERT_NONE  # 测试用
mail = imaplib.IMAP4_SSL(IMAP_SERVER, PORT, ssl_context=context)
mail.login(USERNAME, PASSWORD)
mail.select('INBOX')

# 搜索最近 5 天的所有邮件
status, messages = mail.search(None, f'SINCE {since_date}')
messages = messages[0].split(b' ')

if not messages:
    log("最近 5 天无邮件，退出。")
else:
    log(f"找到 {len(messages)} 封邮件（{since_date} 至今）。")

    for num in messages:
        try:
            status, msg_data = mail.fetch(num, '(RFC822)')
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            # 解码主题
            subject = decode_filename(msg['Subject'])
            if not subject:
                subject = '无主题'
            log(f"处理邮件：{subject}")

            # 遍历附件
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                filename = decode_filename(part.get_filename())
                if filename:
                    filename = sanitize_filename(filename)
                    filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                    if not os.path.exists(filepath):
                        with open(filepath, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        log(f"下载附件：{filename}")
                    else:
                        log(f"附件已存在，跳过：{filename}")

            # 可选：标记已读
            # mail.store(num, '+FLAGS', '\\Seen')
        except Exception as e:
            log(f"处理邮件 {num.decode()} 失败：{e}")
            continue

mail.close()
mail.logout()
log("下载完成！")
# -*- coding: utf-8 -*-
import os
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import pythoncom
import time
import gc
import imaplib
import email
from email import policy
from email.parser import BytesParser

def get_imap_config(email, password):
    """自动发现IMAP配置"""
    domain = email.split('@')[1].lower()

    configs = {
        'gmail.com': ('imap.gmail.com', 993),
        'outlook.com': ('outlook.office365.com', 993),
        'qq.com': ('imap.qq.com', 993),
        '163.com': ('imap.163.com', 993),
        '126.com': ('imap.126.com', 993),
    }

    if domain in configs:
        host, port = configs[domain]
    else:
        host = f'imap.{domain}'
        port = 993

    return host, port

def save_msg(mail_item, msg_path):
    """保存邮件为MSG格式"""
    try:
        mail_item.SaveAs(msg_path, 3)  # olMSG = 3
        return True
    except Exception as e:
        print(f"  Save error: {e}", flush=True)
        return False

def imap_to_msg(email, password, output_dir, max_emails=1000):
    """从IMAP服务器下载邮件并保存为MSG格式"""
    print(f"Starting IMAP to MSG conversion...", flush=True)
    pythoncom.CoInitialize()

    # 复用Outlook实例
    outlook = None
    mail_item = None

    try:
        host, port = get_imap_config(email, password)
        print(f"Connecting to {host}:{port}...", flush=True)

        # 连接IMAP
        if port == 993:
            client = imaplib.IMAP4_SSL(host)
        else:
            client = imaplib.IMAP4(host)

        client.login(email, password)
        print("Connected!", flush=True)

        # 获取所有文件夹
        status, folders = client.list()
        print(f"Found {len(folders)} folders", flush=True)

        # 创建一次Outlook实例
        print("Initializing Outlook COM...", flush=True)
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Outlook initialized", flush=True)

        total = 0

        # 处理每个文件夹
        for folder_data in folders:
            if total >= max_emails:
                break

            try:
                # 解析文件夹名称
                parts = folder_data.decode().split('"')
                folder_name = parts[-2] if len(parts) > 1 else folder_data.decode().split()[-1].strip('"')

                if not folder_name:
                    continue

                print(f"Processing folder: {folder_name}", flush=True)

                # 选择文件夹
                status, messages = client.select(folder_name)
                if status != 'OK':
                    continue

                msg_count = int(messages[0])
                if msg_count == 0:
                    continue

                # 获取邮件ID列表
                status, msg_ids = client.search(None, 'ALL')
                if status != 'OK':
                    continue

                msg_id_list = msg_ids[0].split()
                if not msg_id_list:
                    continue

                # 创建输出子目录
                safe_folder_name = folder_name.replace('/', '_')
                folder_path = os.path.join(output_dir, safe_folder_name)
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)

                # 获取最新邮件（从后往前）
                for msg_id in msg_id_list[-max_emails:]:
                    if total >= max_emails:
                        break

                    try:
                        status, msg_data = client.fetch(msg_id, '(RFC822)')
                        if status != 'OK':
                            continue

                        # 解析邮件
                        msg_bytes = msg_data[0][1]
                        msg = email.message_from_bytes(msg_bytes)

                        # 获取主题
                        subject = msg.get('subject', '(No Subject)')
                        for c in '<>:"|?*':
                            subject = subject.replace(c, '_')
                        if len(subject) > 80:
                            subject = subject[:80]

                        # MSG文件路径
                        msg_path = os.path.join(folder_path, f"{subject}_{msg_id.decode()}.msg")

                        # 处理文件名重复
                        counter = 1
                        while os.path.exists(msg_path):
                            msg_path = os.path.join(folder_path, f"{subject}_{msg_id.decode()}_{counter}.msg")
                            counter += 1

                        # 复用Outlook实例创建邮件
                        mail_item = outlook.CreateItem(0)

                        # 设置基本属性
                        mail_item.Subject = subject

                        # 发件人
                        from_header = msg.get('from', '')
                        if from_header:
                            mail_item.Sender = from_header

                        # 收件人
                        to_header = msg.get('to', '')
                        if to_header:
                            mail_item.To = to_header

                        # 抄送
                        cc_header = msg.get('cc', '')
                        if cc_header:
                            mail_item.CC = cc_header

                        # 正文 - 尝试获取HTML或纯文本
                        if msg.is_multipart():
                            html_body = None
                            text_body = None
                            for part in msg.walk():
                                content_type = part.get_content_type()
                                if content_type == 'text/html' and not html_body:
                                    try:
                                        html_body = part.get_content()
                                    except:
                                        pass
                                elif content_type == 'text/plain' and not text_body:
                                    try:
                                        text_body = part.get_content()
                                    except:
                                        pass

                            if html_body:
                                mail_item.HTMLBody = html_body
                            elif text_body:
                                mail_item.Body = text_body
                        else:
                            try:
                                body = msg.get_content()
                                if msg.get_content_type() == 'text/html':
                                    mail_item.HTMLBody = body
                                else:
                                    mail_item.Body = body
                            except:
                                pass

                        # 添加附件
                        for part in msg.walk():
                            if part.get_content_disposition() == 'attachment':
                                try:
                                    filename = part.get_filename()
                                    if filename:
                                        # 保存到临时文件
                                        temp_path = os.path.join(os.environ['TEMP'], filename)
                                        with open(temp_path, 'wb') as f:
                                            f.write(part.get_payload(decode=True))
                                        mail_item.Attachments.Add(temp_path, 1, 0, filename)
                                        try:
                                            os.remove(temp_path)
                                        except:
                                            pass
                                except:
                                    pass

                        # 保存为MSG
                        if save_msg(mail_item, msg_path):
                            total += 1
                            if total % 10 == 0:
                                print(f"Progress: {total}/{max_emails}", flush=True)

                        # 清理当前邮件项目
                        del mail_item
                        mail_item = None
                        gc.collect()

                    except Exception as e:
                        print(f"  Email error: {e}", flush=True)
                        if mail_item:
                            del mail_item
                            mail_item = None
                        continue

            except Exception as e:
                print(f"  Folder error: {e}", flush=True)
                continue

        try:
            client.logout()
        except:
            pass

        print(f"Total: {total} messages saved to MSG", flush=True)
        return True

    except Exception as e:
        print(f"Error: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False
    finally:
        # 清理Outlook COM对象
        if mail_item:
            try:
                del mail_item
            except:
                pass
        if outlook:
            try:
                del outlook
            except:
                pass
        gc.collect()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: imap_to_msg.py <email> <password> <output_dir> [max_emails]")
        sys.exit(1)

    email = sys.argv[1]
    password = sys.argv[2]
    output_dir = sys.argv[3]
    max_emails = int(sys.argv[4]) if len(sys.argv) > 4 else 1000

    success = imap_to_msg(email, password, output_dir, max_emails)
    sys.exit(0 if success else 1)

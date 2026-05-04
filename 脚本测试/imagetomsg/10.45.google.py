# -*- coding: utf-8 -*-
import os
import sys
import io
import win32com.client
import pythoncom
import gc
import imaplib
import email
import re
from email import policy
from pathlib import Path

# 强制输出为 UTF-8 避免 Windows 控制台乱码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def sanitize_filename(name, max_length=120):
    """更严格的文件名清洗"""
    # 替换 Windows 禁止的字符
    name = re.sub(r'[\\/*?:"<>|]', '_', name)
    # 替换换行符等控制字符
    name = "".join(ch for ch in name if ch.isprintable())
    return name[:max_length].strip()

def get_imap_config(email_addr):
    domain = email_addr.split('@')[1].lower()
    configs = {
        'gmail.com': ('imap.gmail.com', 993),
        'outlook.com': ('outlook.office365.com', 993),
        'qq.com': ('imap.qq.com', 993),
        '163.com': ('imap.163.com', 993),
        '126.com': ('imap.126.com', 993),
    }
    return configs.get(domain, (f'imap.{domain}', 993))

def imap_to_msg(email_user, password, output_dir, max_emails=1000):
    print(f"[*] 启动任务: {email_user}", flush=True)
    pythoncom.CoInitialize()
    
    client = None
    outlook = None
    
    try:
        # 1. 连接 IMAP
        host, port = get_imap_config(email_user)
        client = imaplib.IMAP4_SSL(host, port) if port == 993 else imaplib.IMAP4(host, port)
        client.login(email_user, password)
        print(f"[+] 已连接到 IMAP 伺服器: {host}")

        # 2. 初始化 Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # 获取文件夹列表
        _, folders = client.list()
        total_saved = 0

        for folder_item in folders:
            if total_saved >= max_emails: break
            
            # 解析文件夹名 (处理带空格或中文的情况)
            folder_str = folder_item.decode()
            match = re.search(r'"([^"]+)"$', folder_str) or re.search(r' ([^ ]+)$', folder_str)
            folder_name = match.group(1) if match else "INBOX"
            
            status, messages = client.select(f'"{folder_name}"', readonly=True)
            if status != 'OK': continue

            _, data = client.search(None, 'ALL')
            msg_ids = data[0].split()
            if not msg_ids: continue

            print(f"--- 正在处理文件夹: {folder_name} (共 {len(msg_ids)} 封) ---")
            
            # 创建本地目录
            safe_folder = sanitize_filename(folder_name)
            target_path = Path(output_dir) / safe_folder
            target_path.mkdir(parents=True, exist_ok=True)

            # 从新到旧处理
            for msg_id in reversed(msg_ids):
                if total_saved >= max_emails: break
                
                try:
                    # 获取邮件内容
                    _, msg_data = client.fetch(msg_id, '(RFC822)')
                    raw_email = msg_data[0][1]
                    
                    # 使用 policy.default 自动处理编码和换行
                    msg_obj = email.message_from_bytes(raw_email, policy=policy.default)
                    
                    subject = msg_obj.get('subject', 'No Subject')
                    date_str = msg_obj.get('date', '')
                    clean_subject = sanitize_filename(subject)
                    
                    # 构造 MSG
                    mail_item = outlook.CreateItem(0)
                    mail_item.Subject = subject
                    mail_item.SentOnBehalfOfName = msg_obj.get('from', '')
                    mail_item.To = msg_obj.get('to', '')
                    mail_item.CC = msg_obj.get('cc', '')
                    
                    # 写入正文
                    body_part = msg_obj.get_body(preferencelist=('html', 'plain'))
                    if body_part:
                        if body_part.get_content_type() == 'text/html':
                            mail_item.HTMLBody = body_part.get_content()
                        else:
                            mail_item.Body = body_part.get_content()

                    # 处理附件
                    for attachment in msg_obj.iter_attachments():
                        fname = attachment.get_filename()
                        if fname:
                            temp_file = Path(os.environ['TEMP']) / fname
                            with open(temp_file, 'wb') as f:
                                f.write(attachment.get_payload(decode=True))
                            mail_item.Attachments.Add(str(temp_file))
                            os.remove(temp_file)

                    # 保存文件
                    file_name = f"{total_saved:04d}_{clean_subject}.msg"
                    save_full_path = target_path / file_name
                    mail_item.SaveAs(str(save_full_path.absolute()))
                    
                    total_saved += 1
                    if total_saved % 5 == 0:
                        print(f"进度: {total_saved}/{max_emails} 已保存")

                except Exception as e:
                    print(f"  [!] 邮件 ID {msg_id.decode()} 导出失败: {e}")
                finally:
                    # 及时释放 COM 对象，防止 Outlook 进程卡死
                    mail_item = None

        print(f"\n[√] 任务完成！共转换 {total_saved} 封邮件。")
        return True

    except Exception as e:
        print(f"[-] 关键错误: {e}")
        return False
    finally:
        if client: client.logout()
        pythoncom.CoUninitialize()
        gc.collect()

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("用法: python imap_to_msg.py <邮箱> <授权码/密码> <输出目录> [最大数量]")
        sys.exit(1)

    u, p, o = sys.argv[1], sys.argv[2], sys.argv[3]
    m = int(sys.argv[4]) if len(sys.argv) > 4 else 1000
    imap_to_msg(u, p, o, m)
# -*- coding: utf-8 -*-
import os
import sys
import io
import win32com.client
import pythoncom
import gc
import imaplib
import email as email_module
import time
import re
from email.header import decode_header
from email.utils import parseaddr, getaddresses, parsedate_to_datetime
from datetime import timezone

# 1. 环境编码与控制台支持
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
if os.name == 'nt':
    os.system('chcp 65001 > nul')

# 全局计数器
processed_count = 0

def get_imap_config(email, password):
    domain = email.split('@')[1].lower()
    configs = {
        'gmail.com': ('imap.gmail.com', 993),
        'outlook.com': ('outlook.office365.com', 993),
        'qq.com': ('imap.qq.com', 993),
        '163.com': ('imap.163.com', 993),
        '126.com': ('imap.126.com', 993),
    }
    return configs.get(domain, (f'imap.{domain}', 993))

def get_clean_header(header_raw):
    """彻底解决邮件头乱码和 b'...' 问题"""
    if not header_raw: return ""
    if isinstance(header_raw, bytes):
        header_raw = header_raw.decode('utf-8', errors='replace')
    try:
        parts = decode_header(header_raw)
        res = ""
        for part, enc in parts:
            if isinstance(part, bytes):
                for codec in [enc or 'utf-8', 'utf-8', 'gb18030', 'gbk', 'latin1']:
                    try:
                        res += part.decode(codec)
                        break
                    except: continue
            else: res += str(part)
        return res.strip()
    except: return str(header_raw)

def safe_str_uid(uid):
    """强制将 IMAP UID 转换为纯数字字符串，去除 b' ' 干扰"""
    if isinstance(uid, bytes):
        return uid.decode('utf-8').strip("'")
    return str(uid).strip("'")

def process_single_email(client, outlook, uid, folder_path):
    """处理单个邮件，返回是否成功"""
    global processed_count
    try:
        _, data = client.uid('fetch', uid, '(RFC822)')
        if not data or not data[0]:
            return False

        raw_msg = email_module.message_from_bytes(data[0][1])

        subject = get_clean_header(raw_msg.get('subject', ''))
        from_info = get_clean_header(raw_msg.get('from', ''))
        to_info = get_clean_header(raw_msg.get('to', ''))
        cc_info = get_clean_header(raw_msg.get('cc', ''))
        date_info = raw_msg.get('date', '')

        # 获取收件箱并创建邮件
        namespace = outlook.GetNamespace("MAPI")
        inbox_folder = namespace.GetDefaultFolder(6)
        mail_item = inbox_folder.Items.Add()
        mail_item.Subject = subject

        # 设置收件人 - 使用 getaddresses 正确解析
        try:
            to_addrs = getaddresses([to_info]) if to_info else []
            if to_addrs:
                # 只使用邮件地址，不包含显示名称
                to_list = [addr for name, addr in to_addrs]
                mail_item.To = "; ".join(to_list)
        except: pass

        # 设置抄送
        try:
            cc_addrs = getaddresses([cc_info]) if cc_info else []
            if cc_addrs:
                cc_list = [addr for name, addr in cc_addrs]
                mail_item.CC = "; ".join(cc_list)
        except: pass

        # 设置密送
        try:
            bcc_info = get_clean_header(raw_msg.get('bcc', ''))
            bcc_addrs = getaddresses([bcc_info]) if bcc_info else []
            if bcc_addrs:
                bcc_list = [addr for name, addr in bcc_addrs]
                mail_item.BCC = "; ".join(bcc_list)
        except: pass

        # 设置优先级
        try:
            priority = raw_msg.get('x-priority', '')
            importance_val = 1  # Normal
            if priority in ['1', '2']:
                importance_val = 2  # High
            elif priority == '5':
                importance_val = 0  # Low
            mail_item.Importance = importance_val
        except: pass

        # 设置发送时间
        try:
            date_str = raw_msg.get('date', '')
            if date_str:
                dt = parsedate_to_datetime(date_str)
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                mail_item.SentOn = dt.replace(tzinfo=None)
        except: pass

        # 先保存邮件，确保基础属性已设置
        mail_item.Save()

        # 写入发件人属性（需要在保存后设置）
        prop_accessor = mail_item.PropertyAccessor
        tag_prefix = "http://schemas.microsoft.com/mapi/proptag/"
        try:
            display_name, sender_addr = parseaddr(from_info)
            prop_accessor.SetProperty(f"{tag_prefix}0x0C1F001F", sender_addr)
            prop_accessor.SetProperty(f"{tag_prefix}0x0C1A001F", display_name or sender_addr)
            prop_accessor.SetProperty(f"{tag_prefix}0x0E070003", 1)
        except: pass

        # 获取并注入原始邮件头
        try:
            raw_headers = str(raw_msg.as_string()).split('\r\n\r\n')[0]
            if raw_headers and len(raw_headers) < 4000:
                prop_accessor.SetProperty(f"{tag_prefix}0x007D001F", raw_headers)
        except: pass

        # 注入 Message-ID
        try:
            msg_id = raw_msg.get('message-id', '')
            if msg_id:
                prop_accessor.SetProperty(f"{tag_prefix}0x0C1A001F", msg_id)
        except: pass

        # 再次保存
        mail_item.Save()

        # 处理正文和附件
        html_content = ""
        text_content = ""

        for part in raw_msg.walk():
            content_type = part.get_content_type()
            filename = part.get_filename()

            if filename:
                fname = get_clean_header(filename)
                safe_fname = re.sub(r'[\\/:*?"<>|]', '_', fname)
                temp_att_path = os.path.abspath(os.path.join(os.environ['TEMP'], f"tmp_{uid}_{safe_fname}"))

                payload = part.get_payload(decode=True)
                if payload:
                    with open(temp_att_path, 'wb') as f: f.write(payload)
                    mail_item.Attachments.Add(temp_att_path)
                    mail_item.Save()
                    os.remove(temp_att_path)

            elif content_type == 'text/html':
                payload = part.get_payload(decode=True)
                html_content = payload.decode(part.get_content_charset() or 'utf-8', errors='replace')
            elif content_type == 'text/plain':
                payload = part.get_payload(decode=True)
                text_content = payload.decode(part.get_content_charset() or 'utf-8', errors='replace')

        if html_content:
            mail_item.HTMLBody = html_content
        elif text_content:
            mail_item.Body = text_content

        mail_item.Save()

        # 文件命名
        clean_subject = re.sub(r'[\\/:*?"<>|]', '_', subject).strip()[:80] or "NoSubject"
        final_filename = f"{clean_subject}.msg"
        final_path = os.path.join(folder_path, final_filename)

        # 处理同名文件
        counter = 1
        while os.path.exists(final_path):
            final_path = os.path.join(folder_path, f"{clean_subject} ({counter}).msg")
            counter += 1

        mail_item.SaveAs(os.path.abspath(final_path), 9)
        mail_item.Delete()

        processed_count += 1
        print(f"  [{processed_count}] 成功保存: {os.path.basename(final_path)}")

        return True

    except Exception as e:
        print(f"  UID {uid} 出错: {e}")
        return False

def imap_to_msg_final(email_addr, password, output_dir, max_emails=99999, folder_filter=None):
    print(f"正在启动任务: {email_addr} (单线程模式)")
    pythoncom.CoInitialize()

    client = None
    outlook = None
    global processed_count
    processed_count = 0

    try:
        # 创建IMAP连接
        host, port = get_imap_config(email_addr, password)
        client = imaplib.IMAP4_SSL(host, port)
        client.login(email_addr, password)

        # 创建Outlook实例
        outlook = win32com.client.Dispatch("Outlook.Application")

        _, folders = client.list()
        active_folders = []
        for f_data in folders:
            f_raw = f_data.decode()
            match = re.search(r'\((.*?)\)\s+"(.*?)"\s+(.*)', f_raw)
            f_name = match.group(3).strip('"') if match else f_raw.split('"')[-2]
            if folder_filter and f_name not in folder_filter: continue

            res, _ = client.select(f'"{f_name}"', readonly=True)
            if res == 'OK':
                _, msg_data = client.uid('search', None, 'ALL')
                if msg_data[0]:
                    active_folders.append((f_name, len(msg_data[0].split())))

        # 收集所有邮件UID和目标文件夹
        all_emails = []
        for f_name, _ in active_folders:
            client.select(f'"{f_name}"', readonly=True)
            _, msg_data = client.uid('search', None, 'ALL')
            uids = msg_data[0].split()

            safe_folder_name = re.sub(r'[\\/:*?"<>|]', '_', f_name)
            folder_path = os.path.join(output_dir, safe_folder_name)
            if not os.path.exists(folder_path): os.makedirs(folder_path)

            for uid_raw in reversed(uids):
                if len(all_emails) >= max_emails: break
                uid = safe_str_uid(uid_raw)
                all_emails.append((uid, folder_path))

            if len(all_emails) >= max_emails: break

        print(f"共找到 {len(all_emails)} 封邮件，开始单线程处理...")

        # 单线程顺序处理
        for uid, folder_path in all_emails:
            process_single_email(client, outlook, uid, folder_path)

        print(f"\n任务完成！共导出 {processed_count} 封邮件。")

    finally:
        if client:
            try: client.logout()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python script.py <Email> <Password> <OutputDir> [MaxEmails] [Folder1,Folder2,...]")
    else:
        limit = int(sys.argv[4]) if len(sys.argv) > 4 else 100000
        f_filter = [f.strip() for f in sys.argv[5].split(',')] if len(sys.argv) > 5 else None
        imap_to_msg_final(sys.argv[1], sys.argv[2], sys.argv[3], limit, f_filter)
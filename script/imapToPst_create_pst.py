# -*- coding: utf-8 -*-
import os
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import win32com.client
import pythoncom
import time
import gc
import tempfile
from email import policy
from email.parser import BytesParser
from datetime import timezone
from email.utils import parsedate_to_datetime

def get_or_create_folder(root_folder, rel_path):
    """在指定的根文件夹下递归创建/获取层级目录"""
    current_folder = root_folder
    if rel_path == "." or not rel_path:
        return current_folder

    parts = rel_path.replace('\\', '/').split('/')
    for part in parts:
        if not part: continue
        try:
            current_folder = current_folder.Folders(part)
        except:
            try:
                current_folder = current_folder.Folders.Add(part)
            except:
                pass
    return current_folder

def super_fix_sender(mail_item, msg):
    """
    最强力的发件人注入：包含地址类型、显示名、SMTP地址
    """
    try:
        prop_accessor = mail_item.PropertyAccessor
        tag_prefix = "http://schemas.microsoft.com/mapi/proptag/"

        from_str = str(msg.get('from', ''))
        if not from_str: return

        # 定义需要注入的 MAPI 标签
        properties = {
            f"{tag_prefix}0x0C1A001F": from_str,      # 发件人姓名
            f"{tag_prefix}0x0C1F001F": from_str,      # 发件人地址
            f"{tag_prefix}0x0C1E001F": "SMTP",        # 地址类型
            f"{tag_prefix}0x0042001F": from_str,      # 代表姓名
            f"{tag_prefix}0x0065001F": from_str,      # 代表地址
            f"{tag_prefix}0x0064001F": "SMTP",        # 代表地址类型
        }

        for tag, value in properties.items():
            try:
                prop_accessor.SetProperty(tag, value)
            except: pass

        # 抹除草稿标志 (MSGFLAG_READ=1)
        prop_accessor.SetProperty(f"{tag_prefix}0x0E070003", 1)

        # 注入时间
        date_str = msg.get('date')
        if date_str:
            try:
                dt = parsedate_to_datetime(str(date_str))
                if dt.tzinfo is None: dt = dt.replace(tzinfo=timezone.utc)
                prop_accessor.SetProperty(f"{tag_prefix}0x00390040", dt) # 发送时间
                prop_accessor.SetProperty(f"{tag_prefix}0x0E060040", dt) # 接收时间
            except: pass

        mail_item.Save()
    except Exception as e:
        print(f"  [Sender Fix Error]: {e}", flush=True)

def create_pst_and_import(pst_path, eml_dir):
    print("Starting...", flush=True)

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        print("Outlook connected", flush=True)

        abs_pst = os.path.abspath(pst_path)
        abs_eml_dir = os.path.abspath(eml_dir)

        # 删除旧PST
        if os.path.exists(abs_pst):
            try:
                os.remove(abs_pst)
            except:
                pass

        # 创建PST
        print("Creating PST...", flush=True)
        try:
            pst_name = os.path.splitext(os.path.basename(abs_pst))[0]
            ns.Stores.AddPstStore(abs_pst, pst_name)
            print("Created with AddPstStore", flush=True)
        except:
            try:
                ns.AddStore(abs_pst)
                print("Created with AddStore", flush=True)
            except Exception as e:
                print(f"Create failed: {e}", flush=True)
                return

        time.sleep(2)

        # 定位PST根文件夹
        pst_root = None
        for folder in ns.Folders:
            try:
                if folder.Store and folder.Store.FilePath and \
                   os.path.abspath(folder.Store.FilePath).lower() == abs_pst.lower():
                    pst_root = folder
                    print(f"Found PST folder: {folder.Name}", flush=True)
                    break
            except: continue

        if not pst_root:
            print("ERROR: Cannot locate PST.", flush=True)
            return

        # 获取草稿箱作为中转
        local_drafts = ns.GetDefaultFolder(16)
        print("Using drafts as intermediate", flush=True)

        count = 0

        for root, dirs, files in os.walk(abs_eml_dir):
            rel_path = os.path.relpath(root, abs_eml_dir)
            target_folder = get_or_create_folder(pst_root, rel_path)

            eml_files = [f for f in files if f.lower().endswith('.eml')]
            for file in eml_files:
                try:
                    eml_path = os.path.join(root, file)
                    with open(eml_path, 'rb') as f:
                        msg = BytesParser(policy=policy.default).parse(f)

                    # 1. 中转站创建
                    mail = local_drafts.Items.Add(0)
                    mail.Subject = str(msg['subject'])[:250] if msg['subject'] else "(No Subject)"
                    mail.To = str(msg.get('to', ''))
                    mail.CC = str(msg.get('cc', ''))

                    body_part = msg.get_body(preferencelist=('html', 'plain'))
                    if body_part:
                        if body_part.get_content_type() == 'text/html':
                            mail.HTMLBody = body_part.get_content()
                        else:
                            mail.Body = body_part.get_content()

                    for part in msg.iter_attachments():
                        fname = part.get_filename()
                        if fname:
                            tmp = os.path.join(tempfile.gettempdir(), f"p_{int(time.time())}_{fname}")
                            with open(tmp, 'wb') as af: af.write(part.get_payload(decode=True))
                            mail.Attachments.Add(tmp)
                            try:
                                os.remove(tmp)
                            except:
                                pass

                    # 初次修复并保存
                    super_fix_sender(mail, msg)

                    # 2. 移动到 PST
                    moved_mail = mail.Move(target_folder)

                    # 3. 【关键】移动之后再次执行属性锁死！
                    super_fix_sender(moved_mail, msg)

                    count += 1
                    if count % 10 == 0:
                        print(f"Progress: {count}", flush=True)

                    del mail
                    del moved_mail

                except Exception as e:
                    print(f"  [Error] {file}: {e}", flush=True)

        print(f"Total emails imported: {count}", flush=True)
        print("SUCCESS", flush=True)

    except Exception as e:
        print(f"ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
    finally:
        gc.collect()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py <pst_path> <eml_dir>")
    else:
        create_pst_and_import(sys.argv[1], sys.argv[2])

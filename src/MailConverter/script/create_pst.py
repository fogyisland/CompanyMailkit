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
from datetime import timezone, datetime
from email.utils import parsedate_to_datetime
import re

# Outlook item types
olMailItem = 0
olContactItem = 2
olAppointmentItem = 1
olFolderContacts = 10
olFolderCalendar = 9

def parse_vcf(vcf_path):
    """解析VCF文件，返回字典"""
    with open(vcf_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    result = {}
    for line in content.split('\n'):
        if ':' in line:
            key, val = line.split(':', 1)
            key = key.split(';')[0].upper()
            val = val.strip()
            if key in result:
                if isinstance(result[key], list):
                    result[key].append(val)
                else:
                    result[key] = [result[key], val]
            else:
                result[key] = val
    return result

def create_contact_item(ns, contact_folder, vcf_path):
    """从VCF文件创建Outlook联系人"""
    try:
        vcf_data = parse_vcf(vcf_path)
        contact = contact_folder.Items.Add(olContactItem)

        # 姓名
        fn = vcf_data.get('FN', '')
        n = vcf_data.get('N', '')
        if n and ';' in n:
            parts = n.split(';')
            contact.FirstName = parts[1].strip() if len(parts) > 1 else ''
            contact.LastName = parts[0].strip() if len(parts) > 0 else ''
        if fn:
            contact.FullName = fn

        # 邮箱
        email = vcf_data.get('EMAIL')
        if email:
            if isinstance(email, list):
                contact.Email1Address = email[0]
                if len(email) > 1:
                    contact.Email2Address = email[1]
            else:
                contact.Email1Address = email

        # 电话
        tel = vcf_data.get('TEL', '')
        if isinstance(tel, list):
            for i, t in enumerate(tel):
                if i == 0:
                    contact.PrimaryTelephoneNumber = t
                elif i == 1:
                    contact.BusinessTelephoneNumber = t
                elif i == 2:
                    contact.MobileTelephoneNumber = t
        elif tel:
            contact.PrimaryTelephoneNumber = tel

        # 公司
        org = vcf_data.get('ORG')
        if org:
            contact.CompanyName = str(org).split(';')[0]

        # 职位
        title = vcf_data.get('TITLE')
        if title:
            contact.JobTitle = title

        # 部门
        dept = vcf_data.get('DEPT')
        if dept:
            contact.Department = dept

        contact.Save()
        return True
    except Exception as e:
        print(f"  [Contact Error] {vcf_path}: {e}", flush=True)
        return False

def parse_ics(ics_path):
    """解析ICS文件，返回事件数据"""
    with open(ics_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    # 提取VEVENT块
    event_match = re.search(r'BEGIN:VEVENT(.*?)END:VEVENT', content, re.DOTALL)
    if not event_match:
        return None

    event_content = event_match.group(1)
    result = {}

    for line in event_content.split('\n'):
        line = line.strip()
        if ':' in line:
            key, val = line.split(':', 1)
            val = val.strip()
            # 处理转义的字符
            val = val.replace('\\n', '\n').replace('\\,', ',').replace('\\;', ';').replace('\\\\', '\\')
            result[key.upper()] = val

    return result

def create_appointment_item(ns, calendar_folder, ics_path):
    """从ICS文件创建Outlook日历项"""
    try:
        event_data = parse_ics(ics_path)
        if not event_data:
            return False

        appt = calendar_folder.Items.Add(olAppointmentItem)

        # 主题
        summary = event_data.get('SUMMARY', '')
        appt.Subject = summary if summary else "(No Subject)"

        # 开始时间
        dtstart = event_data.get('DTSTART', '')
        if dtstart:
            try:
                if 'T' in dtstart:
                    dt_format = "%Y%m%dT%H%M%S"
                    if dtstart.endswith('Z'):
                        dt_format = "%Y%m%dT%H%M%SZ"
                    appt.Start = datetime.strptime(dtstart[:15], dt_format)
                else:
                    appt.Start = datetime.strptime(dtstart[:8], "%Y%m%d")
            except:
                appt.Start = datetime.now()

        # 结束时间
        dtend = event_data.get('DTEND', '')
        if dtend:
            try:
                if 'T' in dtend:
                    dt_format = "%Y%m%dT%H%M%S"
                    if dtend.endswith('Z'):
                        dt_format = "%Y%m%dT%H%M%SZ"
                    appt.End = datetime.strptime(dtend[:15], dt_format)
                else:
                    appt.End = datetime.strptime(dtend[:8], "%Y%m%d")
            except:
                pass

        # 地点
        location = event_data.get('LOCATION', '')
        if location:
            appt.Location = location

        # 正文
        description = event_data.get('DESCRIPTION', '')
        if description:
            appt.Body = description

        # 全天事件
        if 'T' not in (dtstart or ''):
            appt.AllDayEvent = True

        appt.Save()
        return True
    except Exception as e:
        print(f"  [Calendar Error] {ics_path}: {e}", flush=True)
        return False

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

def super_fix_sender(mail_item, msg, eml_path=None):
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
        else:
            # 如果没有Date头部，使用文件修改时间作为默认值
            try:
                file_mtime = os.path.getmtime(eml_path)
                dt = datetime.fromtimestamp(file_mtime, tz=timezone.utc)
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

        # 获取或创建联系人和日历文件夹
        try:
            contacts_folder = pst_root.Folders.Item("联系人")
        except:
            contacts_folder = pst_root.Folders.Add("联系人")

        try:
            calendar_folder = pst_root.Folders.Item("日历")
        except:
            calendar_folder = pst_root.Folders.Add("日历")

        count = 0
        contact_count = 0
        calendar_count = 0

        for root, dirs, files in os.walk(abs_eml_dir):
            rel_path = os.path.relpath(root, abs_eml_dir)

            # 特殊文件夹处理
            if rel_path == "联系人":
                # 处理VCF联系人文件
                vcf_files = [f for f in files if f.lower().endswith('.vcf')]
                for file in vcf_files:
                    try:
                        vcf_path = os.path.join(root, file)
                        if create_contact_item(ns, contacts_folder, vcf_path):
                            contact_count += 1
                            if contact_count % 10 == 0:
                                print(f"Contacts: {contact_count}", flush=True)
                    except Exception as e:
                        print(f"  [VCF Error] {file}: {e}", flush=True)
                continue

            if rel_path == "日历":
                # 处理ICS日历文件
                ics_files = [f for f in files if f.lower().endswith('.ics')]
                for file in ics_files:
                    try:
                        ics_path = os.path.join(root, file)
                        if create_appointment_item(ns, calendar_folder, ics_path):
                            calendar_count += 1
                            if calendar_count % 10 == 0:
                                print(f"Calendar: {calendar_count}", flush=True)
                    except Exception as e:
                        print(f"  [ICS Error] {file}: {e}", flush=True)
                continue

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

                    # 尝试多种方式获取邮件正文
                    body_set = False
                    body_part = msg.get_body(preferencelist=('html', 'plain'))
                    if body_part:
                        content = body_part.get_content()
                        if content and len(content.strip()) > 0:
                            if body_part.get_content_type() == 'text/html':
                                mail.HTMLBody = content
                            else:
                                mail.Body = content
                            body_set = True

                    # 如果正文为空，尝试其他方式获取
                    if not body_set:
                        # 尝试从多部分消息中直接获取
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            if content_type in ('text/html', 'text/plain'):
                                try:
                                    content = part.get_content()
                                    if content and len(str(content).strip()) > 0:
                                        if content_type == 'text/html':
                                            mail.HTMLBody = str(content)
                                        else:
                                            mail.Body = str(content)
                                        body_set = True
                                        break
                                except:
                                    try:
                                        payload = part.get_payload(decode=True)
                                        if payload:
                                            charset = part.get_content_charset() or 'utf-8'
                                            content = payload.decode(charset, errors='replace')
                                            if len(content.strip()) > 0:
                                                if content_type == 'text/html':
                                                    mail.HTMLBody = content
                                                else:
                                                    mail.Body = content
                                                body_set = True
                                                break
                                    except:
                                        pass

                    # 检查是否是日历邀请
                    if not body_set:
                        for part in msg.iter_attachments():
                            fname = part.get_filename()
                            if fname and (fname.lower().endswith('.ics') or 'calendar' in fname.lower()):
                                try:
                                    cal_content = part.get_payload(decode=True)
                                    if cal_content:
                                        mail.Body = "[日历邀请]\n" + cal_content.decode('utf-8', errors='replace')[:500]
                                        body_set = True
                                except:
                                    pass

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
                    super_fix_sender(mail, msg, eml_path)

                    # 2. 移动到 PST
                    moved_mail = mail.Move(target_folder)

                    # 3. 【关键】移动之后再次执行属性锁死！
                    super_fix_sender(moved_mail, msg, eml_path)

                    count += 1
                    if count % 10 == 0:
                        print(f"Progress: {count}", flush=True)

                    del mail
                    del moved_mail

                except Exception as e:
                    print(f"  [Error] {file}: {e}", flush=True)

        print(f"Total emails imported: {count}", flush=True)
        print(f"Total contacts imported: {contact_count}", flush=True)
        print(f"Total calendar events imported: {calendar_count}", flush=True)
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

# -*- coding: utf-8 -*-
"""通过EWS投递邮件到Exchange/Office 365，支持指定目标文件夹"""
import sys
import os

def deliver_to_ews(email, password, server_url, eml_path, target_folder=None):
    """投递邮件到EWS指定文件夹"""
    try:
        from exchangelib import Credentials, Account, Message, Folder
        from exchangelib.objects import ItemId
        from email.parser import Parser

        print(f"Connecting to EWS: {email}", flush=True)

        # 如果没有提供server_url，使用默认的Office 365 URL
        if not server_url or server_url == "auto":
            server_url = "outlook.office365.com"

        # 创建凭据
        credentials = Credentials(email, password)

        # 连接账户
        account = Account(
            primary_email_address=email,
            credentials=credentials,
            autodiscover=True,
            access_type=None
        )

        print(f"Connected to: {account.primary_smtp_address}", flush=True)

        # 读取EML文件
        with open(eml_path, 'r', encoding='utf-8', errors='ignore') as f:
            eml_content = f.read()

        # 简单解析EML
        subject = "(No Subject)"
        body = ""
        to_addrs = []
        cc_addrs = []
        in_body = False
        body_lines = []

        lines = eml_content.split('\n')
        for line in lines:
            if line.startswith('Subject:'):
                subject = line[8:].strip() or "(No Subject)"
            elif line.startswith('To:'):
                to_str = line[3:].strip()
                to_addrs = [t.strip() for t in to_str.split(',') if t.strip()]
            elif line.startswith('Cc:'):
                cc_str = line[3:].strip()
                cc_addrs = [t.strip() for t in cc_str.split(',') if t.strip()]
            elif line.startswith('\n') or line.startswith('\r\n'):
                in_body = True
            elif in_body:
                body_lines.append(line)

        body = '\n'.join(body_lines).strip()

        # 确定目标文件夹
        target = account.inbox
        if target_folder:
            # 尝试查找或创建目标文件夹
            try:
                # 尝试直接访问子文件夹
                folder_parts = target_folder.split('/')
                current = account.root
                for part in folder_parts:
                    if not part.strip():
                        continue
                    try:
                        current = current / part
                    except:
                        # 文件夹不存在，创建它
                        current = current.create_subfolder(part)
                target = current
                print(f"Target folder: {target.name}", flush=True)
            except Exception as e:
                print(f"Folder access error, using inbox: {e}", flush=True)
                target = account.inbox

        # 创建邮件并保存到目标文件夹
        msg = Message(
            account=account,
            subject=subject,
            body=body,
            to_recipients=to_addrs,
            cc_recipients=cc_addrs
        )

        # 保存到指定文件夹
        msg.folder = target
        msg.save()

        print(f"Saved to {target.name}: {subject}", flush=True)
        return True

    except ImportError:
        print("exchangelib not installed", flush=True)
        return False

    except Exception as e:
        print(f"EWS Error: {e}", flush=True)
        # 尝试使用Outlook COM作为备选
        try:
            return deliver_via_outlook(email, password, eml_path, target_folder)
        except:
            return False


def deliver_via_outlook(email, password, eml_path, target_folder=None):
    """通过Outlook COM投递（适用于本地Exchange混合模式）"""
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        # 查找目标邮箱文件夹
        target = None
        if target_folder:
            # 尝试查找文件夹
            try:
                # 遍历查找匹配的文件夹
                for folder in ns.Folders:
                    if email.lower() in folder.Name.lower():
                        try:
                            for subfolder in folder.Folders:
                                if target_folder.lower() in subfolder.Name.lower():
                                    target = subfolder
                                    break
                        except:
                            pass
            except:
                pass

        if not target:
            # 使用收件箱
            target = ns.GetDefaultFolder(6)  # 6 = inbox

        # 读取EML
        with open(eml_path, 'r', encoding='utf-8', errors='ignore') as f:
            eml_content = f.read()

        # 解析
        subject = "(No Subject)"
        body = ""
        to_addr = ""
        in_body = False
        body_lines = []

        lines = eml_content.split('\n')
        for line in lines:
            if line.startswith('Subject:'):
                subject = line[8:].strip() or "(No Subject)"
            elif line.startswith('To:'):
                to_addr = line[3:].strip()
            elif line.startswith('\n') or line.startswith('\r\n'):
                in_body = True
            elif in_body:
                body_lines.append(line)

        body = '\n'.join(body_lines).strip()

        # 创建邮件
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        if to_addr:
            mail.To = to_addr

        # 保存到目标文件夹
        mail.Save()
        mail.Move(target)

        print(f"Saved: {subject}", flush=True)
        return True

    except Exception as e:
        print(f"Outlook Error: {e}", flush=True)
        return False
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Usage: ews_deliver.py <email> <password> <server_url> <eml_path> [target_folder]")
        print("server_url can be 'auto' for Office 365 autodiscover")
        print("target_folder is optional, e.g., 'Archive/2024'")
        sys.exit(1)

    email = sys.argv[1]
    password = sys.argv[2]
    server_url = sys.argv[3]
    eml_path = sys.argv[4]
    target_folder = sys.argv[5] if len(sys.argv) > 5 else None

    success = deliver_to_ews(email, password, server_url, eml_path, target_folder)
    sys.exit(0 if success else 1)

# -*- coding: utf-8 -*-
"""实时投递单封邮件到PST"""
import sys
import os
import win32com.client
import pythoncom

def deliver_mail_to_pst(pst_path, eml_path):
    """投递单封邮件到PST"""
    pythoncom.CoInitialize()

    outlook = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        # 读取EML文件
        with open(eml_path, 'r', encoding='utf-8', errors='ignore') as f:
            eml_content = f.read()

        # 解析简单的邮件头
        subject = "(No Subject)"
        body = ""
        to_addr = ""

        lines = eml_content.split('\n')
        in_body = False
        body_lines = []

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

        # 查找PST文件夹
        pst_folder = None
        for folder in ns.Folders:
            try:
                if folder.Store and folder.Store.FilePath:
                    if os.path.abspath(folder.Store.FilePath).lower() == os.path.abspath(pst_path).lower():
                        pst_folder = folder
                        break
            except:
                pass

        if not pst_folder:
            print(f"ERROR: PST folder not found: {pst_path}", flush=True)
            return False

        # 创建邮件并移动到PST
        new_mail = outlook.CreateItem(0)
        new_mail.Subject = subject[:255]

        if body:
            new_mail.Body = body

        if to_addr:
            new_mail.To = to_addr

        # 保存并移动到PST
        new_mail.Save()
        new_mail.Move(pst_folder)

        print(f"Delivered: {subject}", flush=True)
        return True

    except Exception as e:
        print(f"ERROR: {e}", flush=True)
        return False
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: deliver_mail.py <pst_path> <eml_path>")
        sys.exit(1)

    pst_path = sys.argv[1]
    eml_path = sys.argv[2]

    success = deliver_mail_to_pst(pst_path, eml_path)
    sys.exit(0 if success else 1)

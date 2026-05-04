# -*- coding: utf-8 -*-
import sys
import os
import win32com.client
import pythoncom
import time
import gc
from email import policy
from email.generator import BytesGenerator
from io import BytesIO

def extract_emails_from_pst_or_ost(output_dir, store_path=None):
    """Extract emails from PST or OST to EML files"""
    print("Starting email extraction...", flush=True)
    pythoncom.CoInitialize()

    outlook = None
    ns = None
    try:
        print("Connecting to Outlook...", flush=True)
        new_outlook = False
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            print("Found running Outlook instance")
        except:
            print("No running Outlook, starting new instance...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            new_outlook = True
            # 等待Outlook启动完成
            time.sleep(5)

        # 如果是新启动的Outlook，等待更长时间确保完全加载
        if new_outlook:
            print("Waiting for Outlook to fully load...", flush=True)
            time.sleep(3)

        print("Getting MAPI namespace...", flush=True)
        ns = outlook.GetNamespace("MAPI")
        print("Outlook connected successfully", flush=True)

        # Find the store (PST or OST)
        target_folder = None
        store_name = None

        if store_path and os.path.exists(store_path):
            # Find by file path
            for folder in ns.Folders:
                try:
                    store = folder.Store
                    if store and store.FilePath:
                        if os.path.abspath(store.FilePath).lower() == os.path.abspath(store_path).lower():
                            target_folder = folder
                            store_name = folder.Name
                            file_type = "PST" if store.FilePath.lower().endswith('.pst') else "OST"
                            print(f"Found {file_type} by path: {folder.Name}", flush=True)
                            break
                except: pass
        else:
            # Find first available PST/OST
            for folder in ns.Folders:
                try:
                    store = folder.Store
                    if store and store.FilePath:
                        if '.pst' in store.FilePath.lower() or '.ost' in store.FilePath.lower():
                            target_folder = folder
                            store_name = folder.Name
                            file_type = "PST" if store.FilePath.lower().endswith('.pst') else "OST"
                            print(f"Found {file_type}: {folder.Name}", flush=True)
                            break
                except: pass

        if not target_folder and store_path:
            # 尝试自动添加PST文件
            print(f"Attempting to add PST file: {store_path}", flush=True)
            try:
                # 检查文件是否被锁定
                try:
                    with open(store_path, 'rb') as f:
                        f.read(1)
                    print("文件未被锁定", flush=True)
                except IOError as e:
                    print(f"警告: 文件可能被锁定: {e}", flush=True)

                # 添加PST文件到Outlook - 使用AddStoreEx
                # olStoreUnicode = 3
                try:
                    ns.Stores.Add(store_path)
                except:
                    # 如果AddStore失败，尝试AddStoreEx
                    try:
                        print("尝试使用 AddStoreEx...", flush=True)
                        # 0x00000003 = olStoreUnicode
                        ns.Stores.AddStoreEx(store_path, 3)
                    except Exception as e2:
                        raise e2

                print("PST file added successfully", flush=True)
                # 等待添加完成
                time.sleep(3)

                # 再次查找
                for folder in ns.Folders:
                    try:
                        store = folder.Store
                        if store and store.FilePath:
                            if os.path.abspath(store.FilePath).lower() == os.path.abspath(store_path).lower():
                                target_folder = folder
                                store_name = folder.Name
                                print(f"Found added PST: {folder.Name}", flush=True)
                                break
                    except: pass
            except Exception as e:
                print(f"Failed to add PST: {e}", flush=True)

        if not target_folder:
            print("ERROR: No PST/OST store found!", flush=True)
            print("Please manually add the PST/OST file in Outlook first, or use '文件 > 打开 > 添加账户'", flush=True)
            return False

        # Create output directory
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        print(f"Extracting to: {output_dir}", flush=True)

        total_emails = 0
        folder_count = 0

        # Recursively extract emails from all folders
        def extract_from_folder(folder, base_path):
            nonlocal total_emails, folder_count

            try:
                items = folder.Items
                count = items.Count
            except:
                count = 0

            if count > 0:
                folder_name = folder.Name
                if base_path:
                    folder_path = os.path.join(base_path, folder_name)
                else:
                    folder_path = folder_name

                print(f"Folder: {folder_name} ({count} emails)", flush=True)
                folder_count += 1

                # Create folder
                target_folder_path = output_dir
                if folder_path and folder_path != 'Outlook Data File':
                    target_folder_path = os.path.join(output_dir, folder_path)
                    if not os.path.exists(target_folder_path):
                        os.makedirs(target_folder_path)

                # Extract each email
                for i in range(1, count + 1):
                    try:
                        item = items(i)
                        if item.Class != 43:  # IPM.Note = Mail Item
                            continue

                        # Get email properties
                        subject = item.Subject or "No Subject"
                        # Clean filename
                        invalid_chars = '<>:"/\\|?*'
                        for c in invalid_chars:
                            subject = subject.replace(c, '_')

                        # Generate unique filename
                        eml_filename = f"{subject}_{i}.eml"
                        if len(eml_filename) > 200:
                            eml_filename = eml_filename[:200]
                        eml_path = os.path.join(target_folder_path, eml_filename)

                        # Handle duplicate filenames
                        counter = 1
                        base_name = eml_path
                        while os.path.exists(eml_path):
                            name, ext = os.path.splitext(base_name)
                            eml_path = f"{name}_{counter}{ext}"
                            counter += 1

                        # Try to save as EML using MAPI property
                        try:
                            # Method 1: Try to use PR_TRANSPORT_MESSAGE_HEADERS
                            # This gets the original MIME message
                            propAccessor = item.PropertyAccessor
                            try:
                                # PR_TRANSPORT_MESSAGE_HEADERS = 0x007D
                                headers = propAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
                                if headers:
                                    # Write as EML
                                    with open(eml_path, 'w', encoding='utf-8') as f:
                                        f.write(headers)
                                    total_emails += 1
                                    if total_emails % 50 == 0:
                                        print(f"Progress: {total_emails} emails", flush=True)
                                    continue
                            except:
                                pass

                            # Method 2: Construct EML manually from properties
                            eml_content = construct_eml(item)
                            with open(eml_path, 'w', encoding='utf-8') as f:
                                f.write(eml_content)
                            total_emails += 1

                            if total_emails % 50 == 0:
                                print(f"Progress: {total_emails} emails", flush=True)

                        except Exception as e:
                            print(f"Error saving {eml_filename}: {e}", flush=True)
                            continue

                    except Exception as e:
                        continue

            # Process subfolders
            try:
                for subfolder in folder.Folders:
                    extract_from_folder(subfolder, folder.Name if folder.Name != 'Outlook Data File' else '')
            except:
                pass

        # Start extraction from root
        extract_from_folder(target_folder, '')

        print(f"\n=== Extraction Complete ===", flush=True)
        print(f"Total emails: {total_emails}", flush=True)
        print(f"Folders: {folder_count}", flush=True)

        gc.collect()
        print("SUCCESS!", flush=True)
        return True

    except Exception as e:
        print(f"ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False
    finally:
        gc.collect()
        pythoncom.CoUninitialize()


def construct_eml(item):
    """Construct EML content from Outlook item"""
    lines = []

    # Subject
    subject = item.Subject or ""
    lines.append(f"Subject: {subject}")

    # From
    try:
        if item.SenderEmailAddress:
            sender_name = item.SenderName or item.SenderEmailAddress
            lines.append(f"From: {sender_name} <{item.SenderEmailAddress}>")
    except:
        pass

    # To
    try:
        if item.To:
            lines.append(f"To: {item.To}")
    except:
        pass

    # CC
    try:
        if item.CC:
            lines.append(f"Cc: {item.CC}")
    except:
        pass

    # BCC
    try:
        if item.BCC:
            lines.append(f"Bcc: {item.BCC}")
    except:
        pass

    # Date
    try:
        if item.SentOn:
            from email.utils import formatdate
            date_str = formatdate(item.SentOn.timestamp())
            lines.append(f"Date: {date_str}")
    except:
        pass

    # Message-ID
    try:
        if item.EntryID:
            lines.append(f"Message-ID: <{item.EntryID}>")
    except:
        pass

    # Priority
    try:
        if item.Importance == 2:  # High
            lines.append("X-Priority: 1")
        elif item.Importance == 0:  # Low
            lines.append("X-Priority: 5")
    except:
        pass

    # MIME-Version
    lines.append("MIME-Version: 1.0")

    # Content-Type
    has_html = False
    try:
        if hasattr(item, 'HTMLBody') and item.HTMLBody:
            has_html = True
    except:
        pass

    if has_html:
        lines.append("Content-Type: text/html; charset=utf-8")
        lines.append("Content-Transfer-Encoding: base64")
        lines.append("")
        try:
            import base64
            html_body = item.HTMLBody
            if isinstance(html_body, str):
                html_body = html_body.encode('utf-8')
            encoded = base64.b64encode(html_body).decode('ascii')
            # Wrap at 76 chars
            for i in range(0, len(encoded), 76):
                lines.append(encoded[i:i+76])
        except Exception as e:
            lines.append("(Could not encode HTML body)")
    else:
        # Plain text
        lines.append("Content-Type: text/plain; charset=utf-8")
        lines.append("Content-Transfer-Encoding: 7bit")
        lines.append("")
        try:
            body = item.Body or ""
            # Escape dots for SMTP compliance
            body = body.replace("\n.", "\n..")
            lines.append(body)
        except:
            pass

    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: extract_emails.py <output_dir> [pst_or_ost_path]")
        sys.exit(1)

    output_dir = sys.argv[1]
    store_path = sys.argv[2] if len(sys.argv) > 2 else None

    success = extract_emails_from_pst_or_ost(output_dir, store_path)
    sys.exit(0 if success else 1)

# -*- coding: utf-8 -*-
import sys
import os
import win32com.client
import pythoncom
import time
import re

def extract_emails_to_msg(output_dir, store_path=None):
    """Extract emails from PST or OST to MSG files"""
    print("Starting email extraction to MSG...", flush=True)
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
                # 使用AddStoreEx添加PST文件 (Unicode格式)
                # olStoreUnicode = 3
                try:
                    ns.AddStoreEx(store_path, 3)
                    print("PST file added successfully (AddStoreEx)", flush=True)
                except:
                    # 备选使用普通AddStore
                    ns.Stores.Add(store_path)
                    print("PST file added successfully (AddStore)", flush=True)

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
            print("Please manually add the PST/OST file in Outlook first", flush=True)
            return False

        # Debug: 列出所有可用的stores
        print("Available stores in Outlook:", flush=True)
        for folder in ns.Folders:
            try:
                if folder.Store and folder.Store.FilePath:
                    print(f"  - {folder.Name}: {folder.Store.FilePath}", flush=True)
            except: pass

        print(f"Target folder: {target_folder.Name}", flush=True)

        # Create output directory
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        print(f"Extracting to: {output_dir}", flush=True)

        total_emails = 0
        folder_count = 0

        # Recursively extract emails from all folders
        def extract_from_folder(folder, base_path):
            nonlocal total_emails, folder_count

            folder_name = folder.Name
            if base_path:
                folder_path = os.path.join(base_path, folder_name)
            else:
                folder_path = folder_name

            try:
                items = folder.Items
                count = items.Count
                print(f"Checking folder '{folder_name}': {count} items", flush=True)
            except Exception as e:
                count = 0
                print(f"Error getting items for folder '{folder_name}': {e}", flush=True)

            if count > 0:
                print(f"Folder: {folder_name} ({count} emails)", flush=True)
                folder_count += 1

                # Create folder
                target_folder_path = output_dir
                if folder_path and folder_path != 'Outlook Data File':
                    target_folder_path = os.path.join(output_dir, folder_path)
                    if not os.path.exists(target_folder_path):
                        os.makedirs(target_folder_path)

                # Extract each email as MSG
                for i in range(1, count + 1):
                    try:
                        item = items(i)
                        if item.Class != 43:  # IPM.Note = Mail Item
                            continue

                        # Get email subject for filename
                        subject = item.Subject or "No Subject"
                        # Clean filename
                        invalid_chars = '<>:"/\\|?*'
                        for c in invalid_chars:
                            subject = subject.replace(c, '_')

                        # Limit subject length
                        if len(subject) > 100:
                            subject = subject[:100]

                        # Generate unique filename
                        msg_filename = f"{subject}_{i}.msg"
                        msg_path = os.path.join(target_folder_path, msg_filename)

                        # Handle duplicate filenames
                        counter = 1
                        base_name = msg_path
                        while os.path.exists(msg_path):
                            name, ext = os.path.splitext(base_name)
                            msg_path = f"{name}_{counter}{ext}"
                            counter += 1

                        # Save as MSG
                        try:
                            item.SaveAs(os.path.abspath(msg_path), 9)  # 9 = MSG format
                            total_emails += 1
                            if total_emails % 50 == 0:
                                print(f"Progress: {total_emails} emails", flush=True)
                        except Exception as e:
                            print(f"Error saving {msg_filename}: {e}", flush=True)
                            continue

                    except Exception as e:
                        continue

            # Process subfolders - 始终遍历子文件夹，即使当前文件夹没有邮件
            try:
                print(f"Checking subfolders of '{folder.Name}'...", flush=True)
                for subfolder in folder.Folders:
                    try:
                        print(f"  Found subfolder: {subfolder.Name}", flush=True)
                    except:
                        pass
                    extract_from_folder(subfolder, folder_path)
            except Exception as e:
                print(f"Error processing subfolders: {e}", flush=True)

        # Start extraction from root
        extract_from_folder(target_folder, "")

        print(f"\nExtraction complete! Total: {total_emails} emails from {folder_count} folders", flush=True)
        return True

    except Exception as e:
        print(f"ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return False
    finally:
        try:
            if outlook:
                outlook.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_emails_to_msg.py <OutputDir> [StorePath]")
    else:
        output_dir = sys.argv[1]
        store_path = sys.argv[2] if len(sys.argv) > 2 else None
        extract_emails_to_msg(output_dir, store_path)

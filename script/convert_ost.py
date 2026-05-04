# -*- coding: utf-8 -*-
import sys
import os
import win32com.client
import pythoncom
import time
import gc

def convert_ost_to_pst(pst_path, ost_path=None):
    """Convert OST to PST"""
    print("Starting OST to PST conversion...", flush=True)
    pythoncom.CoInitialize()

    outlook = None
    ns = None
    pst_folder = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        print("Outlook connected", flush=True)

        # Find OST folder
        ost_folder = None
        if ost_path and os.path.exists(ost_path):
            for folder in ns.Folders:
                try:
                    store = folder.Store
                    if store and store.FilePath and os.path.abspath(store.FilePath).lower() == os.path.abspath(ost_path).lower():
                        ost_folder = folder
                        print(f"Found OST by path: {folder.Name}", flush=True)
                        break
                except: pass
        else:
            for folder in ns.Folders:
                try:
                    store = folder.Store
                    if store and store.FilePath and '.ost' in store.FilePath.lower():
                        ost_folder = folder
                        print(f"Found OST: {folder.Name}", flush=True)
                        break
                except: pass

        if not ost_folder:
            print("ERROR: No OST file found!", flush=True)
            return False

        pst_name = ost_folder.Name or "OST_Export"

        # Delete old PST
        if os.path.exists(pst_path):
            try: os.remove(pst_path)
            except: pass

        print("Creating PST...", flush=True)
        try:
            # Try AddPstStore with display name (preferred method)
            try:
                ns.Stores.AddPstStore(pst_path, pst_name)
                print("Created with AddPstStore(name)", flush=True)
            except:
                # Fallback to AddStore
                ns.AddStore(pst_path)
                print("Created with AddStore", flush=True)
        except Exception as e:
            print(f"Failed: {e}", flush=True)
            return False

        time.sleep(2)

        # Find PST folder
        pst_folder = None
        for folder in ns.Folders:
            try:
                if folder.Store and folder.Store.FilePath:
                    if os.path.abspath(folder.Store.FilePath).lower() == os.path.abspath(pst_path).lower():
                        pst_folder = folder
                        print(f"PST folder: {folder.Name}", flush=True)
                        break
            except: pass

        if not pst_folder:
            print("ERROR: PST folder not found", flush=True)
            return False

        total_emails = 0

        # Create Inbox folder in PST first
        inbox_target = None
        try:
            for f in pst_folder.Folders:
                if f.Name == '收件箱' or f.Name == 'Inbox':
                    inbox_target = f
                    break
        except: pass

        if not inbox_target:
            try:
                inbox_target = pst_folder.Folders.Add("收件箱")
            except:
                inbox_target = pst_folder

        # Copy from OST inbox directly first (most important)
        print("\n=== Copying from Inbox ===", flush=True)
        inbox_folder = None
        for subfolder in ost_folder.Folders:
            if '收件箱' in subfolder.Name or subfolder.Name == 'Inbox':
                inbox_folder = subfolder
                items = inbox_folder.Items
                count = items.Count
                print(f"Inbox: {count} emails", flush=True)

                # Sort by date descending to get most recent first
                try:
                    items.Sort("[ReceivedTime]", True)
                except: pass

                for i in range(1, count + 1):
                    try:
                        mail = items(i)
                        if mail.Class != 43:  # 43 = IPM.Note (Mail Item)
                            continue

                        # Copy email directly to PST Inbox folder
                        try:
                            copied_mail = mail.Copy()
                            copied_mail.Move(inbox_target)
                            total_emails += 1
                        except Exception as e:
                            # Fallback: create new mail manually
                            new_mail = outlook.CreateItem(0)
                            new_mail.Subject = mail.Subject or "(No Subject)"
                            try:
                                if mail.Body:
                                    new_mail.Body = mail.Body
                            except: pass
                            try:
                                if hasattr(mail, 'HTMLBody') and mail.HTMLBody:
                                    new_mail.HTMLBody = mail.HTMLBody
                            except: pass
                            try:
                                if mail.To:
                                    new_mail.To = mail.To
                            except: pass
                            try:
                                if mail.CC:
                                    new_mail.CC = mail.CC
                            except: pass
                            try:
                                if mail.BCC:
                                    new_mail.BCC = mail.BCC
                            except: pass
                            try:
                                if mail.SentOn:
                                    new_mail.SentOn = mail.SentOn
                            except: pass
                            try:
                                if mail.ReceivedTime:
                                    new_mail.ReceivedTime = mail.ReceivedTime
                            except: pass
                            # Set creation time to preserve original date
                            try:
                                new_mail.CreationTime = mail.CreationTime
                            except: pass
                            try:
                                new_mail.LastModificationTime = mail.LastModificationTime
                            except: pass

                            new_mail.Save()
                            new_mail.Move(inbox_target)
                            total_emails += 1

                        if total_emails % 20 == 0:
                            print(f"Progress: {total_emails}", flush=True)

                    except Exception as e:
                        continue

                print(f"Inbox done: {total_emails} emails", flush=True)
                break

        # Copy other important folders
        print("\n=== Copying other folders ===", flush=True)
        for subfolder in ost_folder.Folders:
            folder_name = subfolder.Name
            if '收件箱' in folder_name or folder_name == 'Inbox':
                continue

            items = subfolder.Items
            count = items.Count
            if count == 0:
                continue

            print(f"Folder: {folder_name} ({count} emails)", flush=True)

            # Create folder in PST
            target_folder = None
            try:
                for f in pst_folder.Folders:
                    if f.Name == folder_name:
                        target_folder = f
                        break
            except: pass

            if not target_folder:
                try:
                    target_folder = pst_folder.Folders.Add(folder_name)
                except:
                    target_folder = pst_folder

            # Sort by date
            try:
                items.Sort("[ReceivedTime]", True)
            except: pass

            # Copy emails
            folder_count = 0
            for i in range(1, count + 1):
                try:
                    mail = items(i)
                    if mail.Class != 43:
                        continue

                    try:
                        copied_mail = mail.Copy()
                        copied_mail.Move(target_folder)
                        folder_count += 1
                    except Exception as e:
                        # Fallback: create new mail manually
                        new_mail = outlook.CreateItem(0)
                        new_mail.Subject = mail.Subject or "(No Subject)"
                        try:
                            if mail.Body:
                                new_mail.Body = mail.Body
                        except: pass
                        try:
                            if mail.To:
                                new_mail.To = mail.To
                        except: pass
                        try:
                            if mail.CC:
                                new_mail.CC = mail.CC
                        except: pass
                        try:
                            if mail.SentOn:
                                new_mail.SentOn = mail.SentOn
                        except: pass
                        try:
                            if mail.ReceivedTime:
                                new_mail.ReceivedTime = mail.ReceivedTime
                        except: pass
                        try:
                            new_mail.CreationTime = mail.CreationTime
                        except: pass

                        new_mail.Save()
                        new_mail.Move(target_folder)
                        folder_count += 1

                except:
                    continue

            total_emails += folder_count
            print(f"  Done: {folder_count} emails from {folder_name}", flush=True)

        print(f"\nTotal: {total_emails} emails", flush=True)

        # Verify PST content
        time.sleep(1)
        print(f"PST folder items: {pst_folder.Items.Count}", flush=True)
        for f in pst_folder.Folders:
            print(f"  {f.Name}: {f.Items.Count} items", flush=True)

        gc.collect()
        time.sleep(2)

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

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: convert_ost.py <pst_path> [ost_path]")
        sys.exit(1)

    pst_path = sys.argv[1]
    ost_path = sys.argv[2] if len(sys.argv) > 2 else None

    success = convert_ost_to_pst(pst_path, ost_path)
    sys.exit(0 if success else 1)

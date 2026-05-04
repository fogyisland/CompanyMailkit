# -*- coding: utf-8 -*-
import win32com.client
import pythoncom

pythoncom.CoInitialize()

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    # Find PST
    for folder in ns.Folders:
        try:
            if folder.Store and "ost_backup" in str(folder.Store.FilePath).lower():
                print(f"PST: {folder.Name}")
                print(f"Items: {folder.Items.Count}")
                print("\nSubfolders:")
                for f in folder.Folders:
                    print(f"  {f.Name}: {f.Items.Count} items")
        except: pass

except Exception as e:
    print(f"Error: {e}")
finally:
    pythoncom.CoUninitialize()

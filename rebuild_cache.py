import win32com.client
import json
import os
import pythoncom

CACHE_FILE = r"C:\Users\mjc5422\.gemini\tmp\mjc5422\folder_cache.json"

def rebuild_cache():
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        new_cache = {}
        print("[*] 全ストアからフォルダ地図を再構築中...")
        for store in namespace.Stores:
            try:
                root = store.GetRootFolder()
                stack = [root]
                while stack:
                    folder = stack.pop()
                    # フォルダ名とEntryIDをマッピング
                    new_cache[folder.Name] = folder.EntryID
                    for sub in folder.Folders:
                        stack.append(sub)
            except: pass
        
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(new_cache, f)
        print(f"[*] 地図の更新完了！ {len(new_cache)} 個のフォルダを認識しました。")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    rebuild_cache()

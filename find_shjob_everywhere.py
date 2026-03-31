import win32com.client

def force_search_shjob():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        print("--- SHJOB の正確な位置を力技で捜索中 ---")
        for store in namespace.Stores:
            print(f"Checking Store: {store.DisplayName}")
            try:
                root = store.GetRootFolder()
                def scan_folders(folder):
                    try:
                        items = folder.Items
                        # 新しい順に100件ずつチェック（全部やると時間がかかるので）
                        limit = min(500, items.Count)
                        for i in range(items.Count, items.Count - limit, -1):
                            try:
                                item = items.Item(i)
                                if "SHJOB" in item.Subject.upper():
                                    print(f"発見！ {store.DisplayName} > {folder.Name} > {item.Subject}")
                            except: continue
                    except: pass
                    for sub in folder.Folders:
                        scan_folders(sub)

                scan_folders(root)
            except: continue
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    force_search_shjob()

import win32com.client

def list_folders(folder, indent=''):
    print(f"{indent}{folder.Name}")
    for sub in folder.Folders:
        try:
            list_folders(sub, indent + "  ")
        except:
            pass

def main():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        target_stores = ["受注工程(other)", "検討工程"]

        for store in namespace.Stores:
            # 部分一致でチェック
            if any(t in store.DisplayName for t in target_stores):
                print(f"\n--- Store: {store.DisplayName} ---")
                root = store.GetRootFolder()
                list_folders(root)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()

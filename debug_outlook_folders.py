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

        # Hynix という名前のストアを探す
        for store in namespace.Stores:
            if "Hynix" in store.DisplayName:
                print(f"--- Store: {store.DisplayName} ---")
                root = store.GetRootFolder()
                list_folders(root)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()

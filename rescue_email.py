import win32com.client

def rescue_email():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # 救出対象のメール情報
        target_subject = "【出荷DR開始通知】 三井_EWDTMH24D1 MG7H 64x2 118-C"
        dest_store_name = "三井"
        dest_folder_name = "EWDTMH24D0 MG7H 64x2 118-C"

        # 移動先の受信トレイ
        inbox = namespace.GetDefaultFolder(6)

        print(f"誤って移動したメール '{target_subject}' を救出中...")

        # 三井ストアからフォルダを探す
        for store in namespace.Stores:
            if dest_store_name in store.DisplayName:
                root = store.GetRootFolder()
                # 三井ストアの中を捜索
                def search_and_move_back(folder):
                    try:
                        # フォルダ名が一致するか
                        if dest_folder_name in folder.Name:
                            items = folder.Items
                            for i in range(items.Count, 0, -1):
                                item = items.Item(i)
                                if target_subject in item.Subject:
                                    print(f"発見！ 受信トレイに戻します...")
                                    item.Move(inbox)
                                    return True
                        for sub in folder.Folders:
                            if search_and_move_back(sub): return True
                    except: pass
                    return False

                if search_and_move_back(root):
                    print("救出完了！")
                    return
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    rescue_email()

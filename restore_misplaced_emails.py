import win32com.client
import json
import os

HISTORY_FILE = r"C:\Users\mjc5422\.gemini\tmp\mjc5422\sort_history.json"

def restore_emails():
    if not os.path.exists(HISTORY_FILE): return

    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            history = json.load(f)
    except: return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6) # 戻し先の受信トレイ

        print(f"--- {len(history)}通のメールを救出・復元中 ---")

        for record in reversed(history):
            # 15:00 以降の誤振り分けだけを対象にする
            if "15:00" in record["time"]:
                print(f"復元中: {record['subject']}")
                # 移動先のフォルダを探す
                found = False
                for store in namespace.Stores:
                    root = store.GetRootFolder()
                    def find_in_folders(folder):
                        if folder.Name == record["to"]:
                            items = folder.Items
                            # EntryIDで直接特定
                            try:
                                item = namespace.GetItemFromID(record["id"])
                                item.Move(inbox)
                                return True
                            except:
                                # IDで見つからない場合は件名で探す
                                for i in range(items.Count, 0, -1):
                                    it = items.Item(i)
                                    if it.Subject == record["subject"]:
                                        it.Move(inbox)
                                        return True
                        for sub in folder.Folders:
                            if find_in_folders(sub): return True
                        return False
                    if find_in_folders(root):
                        found = True
                        break
                if not found:
                    print(f"警告: 元フォルダ '{record['to']}' からメールが見つかりませんでした。")

        print("復元完了！")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    restore_emails()

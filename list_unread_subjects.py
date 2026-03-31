import win32com.client

def list_unread_emails_detailed():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        print("--- 未読メールの真の差出人を正確にチェック中 ---")
        found_count = 0
        for store in namespace.Stores:
            try:
                # 受信トレイを取得
                inbox = store.GetDefaultFolder(6) # 6 = olFolderInbox
                items = inbox.Items
                # 効率化のために未読でフィルタリング
                unread_items = items.Restrict("[UnRead] = True")

                # 新しい順に表示
                for i in range(unread_items.Count, 0, -1):
                    try:
                        item = unread_items.Item(i)
                        if item.Class == 43: # MailItem
                            sender_addr = ""
                            try:
                                # Exchange環境ならSenderEmailAddressが空の場合がある
                                sender_addr = item.SenderEmailAddress
                                if not sender_addr or "@" not in sender_addr:
                                    # SMTPアドレスを取得（Exchange環境用）
                                    sender_addr = item.Sender.GetExchangeUser().PrimarySmtpAddress
                            except:
                                sender_addr = item.SenderEmailAddress

                            found_count += 1
                            print(f"{found_count}: FROM:[{sender_addr}] SUBJECT: {item.Subject}")
                    except: continue
            except: continue

        if found_count == 0:
            print("未読メールが見つかりませんでした。")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    list_unread_emails_detailed()

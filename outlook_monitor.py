import win32com.client
import os
import subprocess
import time
import datetime
import outlook_sorter_v2  # 新しく作った仕分けエンジンを再利用

# 未読通知用のEntryID保存
PROCESSED_NOTIFY_FILE = r"C:\Users\mjc5422\.gemini\tmp\mjc5422\processed_notify.txt"

def show_toast(subject, sender):
    """デスクトップ通知（新岡さん宛の大事なメール用）"""
    ps_cmd = f'''
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");
    $notification = New-Object System.Windows.Forms.NotifyIcon;
    $notification.Icon = [System.Drawing.SystemIcons]::Information;
    $notification.BalloonTipTitle = "新岡さん宛の重要メール着信";
    $notification.BalloonTipText = "差出人: {sender}\n件名: {subject}";
    $notification.Visible = $true;
    $notification.ShowBalloonTip(10000);
    '''
    subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True)

def monitor_main():
    print(f"[{datetime.datetime.now()}] Outlookの自動仕分け・監視を開始しました...")

    if not os.path.exists(os.path.dirname(PROCESSED_NOTIFY_FILE)):
        os.makedirs(os.path.dirname(PROCESSED_NOTIFY_FILE), exist_ok=True)
    if not os.path.exists(PROCESSED_NOTIFY_FILE):
        with open(PROCESSED_NOTIFY_FILE, "w") as f: pass

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except:
        print("Outlookに接続できません。終了します。")
        return

    while True:
        try:
            # 【修正】自動仕分け機能は削除し、通知機能のみ残しました。
            # outlook_sorter.main()

            # 2. 「新岡さん」通知のチェック（仕分けられずに残ったものも含めて）
            with open(PROCESSED_NOTIFY_FILE, "r") as f:
                processed_ids = set(line.strip() for line in f)

            inbox = namespace.GetDefaultFolder(6) # メインの受信トレイ
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)

            # 直近20件だけ通知チェック
            limit = min(20, items.Count)
            with open(PROCESSED_NOTIFY_FILE, "a", encoding="utf-8") as f:
                for i in range(1, limit + 1):
                    try:
                        item = items.Item(i)
                        if item.Class == 43 and item.UnRead:
                            if "新岡さん" in item.Body and item.EntryID not in processed_ids: 
                                show_toast(item.Subject, item.SenderName)
                                f.write(item.EntryID + "\n")
                                processed_ids.add(item.EntryID)
                    except: continue

            # 3分おきにチェック（パパのPCの負荷にならないように）
            time.sleep(180)

        except Exception as e:
            print(f"監視中にエラーが発生しました: {e}")
            time.sleep(60)

if __name__ == "__main__":
    monitor_main()

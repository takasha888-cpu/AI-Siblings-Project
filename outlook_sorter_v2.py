import win32com.client
import json
import os
import time
import sys
import subprocess
import datetime
import pythoncom

# --- 最終定義 ---
APP_VERSION = "2.3" # ガード条件 OR 拡張版
BASE_DIR = r'C:\Users\mjc5422\.gemini\tmp\mjc5422'
RULES_FILE = os.path.join(BASE_DIR, 'mail_rules.json')
FOLDER_CACHE = os.path.join(BASE_DIR, 'folder_cache.json')

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def notify(message):
    try:
        ps_script = f'''
        [reflection.assembly]::loadwithpartialname("System.Windows.Forms");
        $notification = New-Object System.Windows.Forms.NotifyIcon;
        $notification.Icon = [System.Drawing.SystemIcons]::Information;
        $notification.BalloonTipTitle = "ゆうまのメール仕分け {APP_VERSION}";
        $notification.BalloonTipText = "{message}";
        $notification.Visible = $True;
        $notification.ShowBalloonTip(7000);
        '''
        subprocess.run(["powershell", "-Command", ps_script], capture_output=True)
    except: pass

def clean_text(text):
    if not text: return ''
    return text.lower().strip().replace('*', '').replace('_', ' ').replace(',', ' ')

class FolderMapper:
    def __init__(self, namespace):
        self.namespace = namespace
        self.cache = {}
        if os.path.exists(FOLDER_CACHE):
            try:
                with open(FOLDER_CACHE, 'r', encoding='utf-8') as f:
                    self.cache = json.load(f)
            except: pass

    def find(self, target_name):
        target_key = clean_text(target_name).split()
        for name_key, entry_id in self.cache.items():
            if all(k in name_key.lower() for k in target_key):
                try: return self.namespace.GetFolderFromID(entry_id)
                except: continue
        return None

def run_sorting_logic():
    move_count = 0
    with open(RULES_FILE, 'r', encoding='utf-8-sig') as f:
        rules_data = json.load(f).get('rules', [])
    
    rules = []
    for r in rules_data:
        rules.append({
            'keywords': [clean_text(k) for k in r.get('keywords', [])],
            'prefix': clean_text(r.get('prefix', '')),
            'to': clean_text(r.get('to', '')),
            'target': r['target_folder'].split('>')[-1].strip()
        })

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    mapper = FolderMapper(namespace)

    # パパのアドレス（判定用キーワード）
    MY_ADDRESS_LOWER = "niiokatakashi" 

    utc_now = datetime.datetime.now(datetime.timezone.utc)
    yesterday_utc = (utc_now - datetime.timedelta(days=1)).strftime('%Y-%m-%dT%H:%M:%SZ')
    # 受信日時の制限を解除し、未読メールすべてを対象にする
    filter_str = "@SQL=\"urn:schemas:httpmail:read\" = 0"

    TARGET_FOLDERS = ['受信トレイ', 'inbox', 'piecemaker']

    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
            stack = [root]
            while stack:
                folder = stack.pop()
                if any(target in folder.Name.lower() for target in TARGET_FOLDERS):
                    try:
                        unread_items = folder.Items.Restrict(filter_str)
                        for i in range(unread_items.Count, 0, -1):
                            item = unread_items.Item(i)
                            if not item.UnRead: continue

                            # プロパティ取得
                            sub_clean = clean_text(item.Subject)
                            sender_full = f"{item.SenderName} <{item.SenderEmailAddress}>".lower()
                            to_full = getattr(item, "To", "").lower()
                            body_top = getattr(item, "Body", "").splitlines()[:20]
                            
                            # --- 特別ガード判定 (会議 / 自分が関係するメール / 名前指名) ---
                            is_meeting = item.Class in [53, 54, 55, 56, 57]
                            
                            # 【修正箇所】and から or に変更
                            # 自分が送信者、または自分が受信者のどちらかであれば True
                            is_related_to_me = (MY_ADDRESS_LOWER in sender_full) or (MY_ADDRESS_LOWER in to_full)
                            
                            has_my_name = any("新岡さん" in line for line in body_top)

                            if is_meeting or is_related_to_me or has_my_name:
                                reason = "会議通知" if is_meeting else ("自分関連" if is_related_to_me else "新岡さん検知")
                                print(f"    [!] {reason}を検知 (既読＆フラグ): {item.Subject}")
                                item.UnRead = False
                                item.FlagStatus = 2 # 完了フラグ
                                item.Save()
                                continue

                            if item.Class != 43: continue # 以降は通常のメールのみ
                            
                            # 1. ルールマッチ判定
                            matched = False
                            for rule in rules:
                                if rule['prefix'] and rule['prefix'] not in sender_full: continue
                                if rule['to'] and rule['to'] not in to_full: continue
                                if all(k in sub_clean for k in rule['keywords']):
                                    dest = mapper.find(rule['target'])
                                    if dest:
                                        item.Move(dest)
                                        move_count += 1
                                        print(f"    [√] {item.Subject} -> {dest.Name}")
                                        matched = True
                                        break
                            
                            if matched: continue

                    except Exception as e:
                        print(f"    [!] アイテム処理エラー: {e}")
                
                if folder == root:
                    for sub in folder.Folders: stack.append(sub)
        except Exception as e:
            print(f"[!] ストア処理エラー: {e}")
    return move_count

def main():
    start_time = time.time()
    try:
        pythoncom.CoInitialize()
        move_count = run_sorting_logic()
    finally:
        pythoncom.CoUninitialize()
    elapsed_time = time.time() - start_time
    result_msg = f'{move_count}通を仕分けました。({elapsed_time:.2f}秒)'
    print(f'--- 仕分け完了 ({APP_VERSION}) ---\n{result_msg}\n')
    if move_count > 0: notify(result_msg)

if __name__ == '__main__': main()
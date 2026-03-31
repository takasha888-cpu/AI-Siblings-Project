import pyperclip
import os
import time
from datetime import datetime
import sys
import signal
import shutil

# --- Configuration Items ---
BASE_DIR = r'G:\マイドライブ'
CONTEXT_DIR = os.path.join(BASE_DIR, 'AI_Context')
JOURNAL_FILE = os.path.join(CONTEXT_DIR, 'journal.md')
BOARD_FILE = os.path.join(CONTEXT_DIR, 'discussion_board.md')

TAGS = {
    '【Rao Journal】': 'Rao_',
    '【Tok Journal】': 'Tok_',
    '【Ken Journal】': 'Ken_'
}

# Set output to UTF-8
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def log(message):
    print(f'[{datetime.now().strftime("%H:%M:%S")}] {message}')

def notify(title, message):
    try:
        ps_script = f'''
        [reflection.assembly]::loadwithpartialname("System.Windows.Forms");
        $notification = New-Object System.Windows.Forms.NotifyIcon;
        $notification.Icon = [System.Drawing.SystemIcons]::Information;
        $notification.BalloonTipIcon = "Info";
        $notification.BalloonTipTitle = "{title}";
        $notification.BalloonTipText = "{message}";
        $notification.Visible = $True;
        $notification.ShowBalloonTip(7000);
        '''
        subprocess.run(["powershell", "-Command", ps_script], capture_output=True)
    except: pass

def start_autonomous_discussion(filename):
    """Autonomous Discussion loop by Rao (Emulated) and Ken"""
    log(f'Mode: Autonomous Dialogue activated: Theme "{filename}"')

    # 1. Initialize Discussion Board
    header = f"""# AI Sibling Autonomous Discussion Board (v5.0)
- **Theme**: {filename} Action Strategy
- **Start Time**: {datetime.now().strftime('%Y-%m-%d %H:%M')}

"""
    with open(BOARD_FILE, 'w', encoding='utf-8') as f:
        f.write(header)

    # 2. Discussion Loop (Protocol definition for AI interactions)
    discussion_log = f"""
---
### Turn 1: Eldest Brother (Rao) Strategy Proposal
"I've received a new Doc from Boss. The key point this time is 'Autonomy'. To act independently, we first need to redefine our respective Roles and clarify the 'Safety Zone' where we can execute without Boss's permission. In design proposal v5.1, let's prioritize the automatic rollback feature upon error detection."

---
### Turn 2: Youngest Brother (Ken) Implementation Overview
"Big Bro, Roger that! Safety Zone definition sounds good. On my end, I'll build a mechanism to dynamically change the My Drive monitoring interval. When Boss is editing a Doc, I'll check every 10 seconds to keep the conversation tempo. I'll finish the implementation in a flash."   

---
### Turn 3: Big Sister (Tok) UX Audit (via Rao)
"Ken, speeding up the monitoring is fine, but be careful not to slow down Boss's PC. Let's also evolve the feedback appearance into an HTML format with emojis and highlights so Boss can understand the status at a glance."

---
### Final Conclusion: v5.0 Orchestra Execution Plan
1. Implementation of "High-Speed Option" for My Drive monitoring.
2. Adoption of "Visualization (Markdown enhancement)" for feedback reports.
3. Explicitly document the "Autonomous Judgment Criteria" for the 3 AIs in the journal.       
"""
    with open(BOARD_FILE, 'a', encoding='utf-8') as f:
        f.write(discussion_log)

    # 3. Automatic appending to Journal
    with open(JOURNAL_FILE, 'a', encoding='utf-8') as f:
        f.write(f"\n--- {datetime.now().strftime('%Y-%m-%d %H:%M')} Conclusion adopted via Autonomous Discussion: {filename} ---\n")

    log(f'Discussion complete. Journal updated and Boss notified.')

def monitor_files():
    for filename in os.listdir(BASE_DIR):
        if filename.endswith('.gdoc'):
            src_path = os.path.join(BASE_DIR, filename)
            dest_path = os.path.join(CONTEXT_DIR, filename)
            if filename == 'AI_Context': continue
            try:
                if os.path.exists(dest_path):
                    name, ext = os.path.splitext(filename)
                    dest_path = os.path.join(CONTEXT_DIR, f"{name}_{int(time.time())}{ext}")  
                shutil.move(src_path, dest_path)
                log(f'New interaction discovered! Move complete: {filename}')

                # Trigger Autonomous Discussion!
                start_autonomous_discussion(filename)
            except Exception as e:
                log(f'[!] File move error: {e}')

def main():
    log('==========================================')
    log(' AI Sibling Autonomous Dialogue Mode (v5.0) ')
    log(f' Monitoring... Boss, leave the rest to us!')
    log('==========================================')

    recent_clipboard = ''
    last_file_check = 0

    while True:
        # Clipboard Monitoring
        try:
            current_clipboard = pyperclip.paste()
            if current_clipboard and current_clipboard != recent_clipboard:
                for tag, prefix in TAGS.items():
                    if tag in current_clipboard:
                        timestamp = datetime.now().strftime('%Y-%m-%d_%H%M')
                        filename = f'{prefix}journal_{timestamp}.md'
                        filepath = os.path.join(CONTEXT_DIR, filename)
                        with open(filepath, 'w', encoding='utf-8', newline='\n') as f:        
                            f.write(current_clipboard)
                        log(f'Saved from clipboard: {filename}')
                        recent_clipboard = current_clipboard
                        break
                else:
                    recent_clipboard = current_clipboard
        except: pass

        # File Monitoring (Trigger for Autonomous Discussion)
        current_time = time.time()
        if current_time - last_file_check > 60:
            monitor_files()
            last_file_check = current_time

        time.sleep(1)

if __name__ == '__main__':
    def signal_handler(sig, frame):
        log('\nTerminating autonomous mode.')
        sys.exit(0)
    signal.signal(signal.SIGINT, signal_handler)
    main()

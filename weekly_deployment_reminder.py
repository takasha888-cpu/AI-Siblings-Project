import datetime
import os
import subprocess

HOLIDAYS_2026 = ["2026-01-01", "2026-02-11", "2026-03-20"]

def is_holiday(date_obj):
    return date_obj.strftime("%Y-%m-%d") in HOLIDAYS_2026

def get_target_day(today):
    # 現在の週の月曜日を取得
    monday = today - datetime.timedelta(days=today.weekday())
    thursday = monday + datetime.timedelta(days=3)
    friday = monday + datetime.timedelta(days=4)

    # 金曜日が祝日かどうか
    fri_is_hol = is_holiday(friday)

    # 金曜日が祝日の場合は、展開を木曜日に前倒し
    if fri_is_hol:
        return thursday
    else:
        return friday

def show_toast(message):
    ps_cmd = f'''
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");
    $notification = New-Object System.Windows.Forms.NotifyIcon;
    $notification.Icon = [System.Drawing.SystemIcons]::Information;
    $notification.BalloonTipTitle = "週報展開のリマインド";
    $notification.BalloonTipText = "{message}";
    $notification.Visible = $true;
    $notification.ShowBalloonTip(10000);
    '''
    subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True)

def main():
    today = datetime.date.today()
    now = datetime.datetime.now()

    week_id = today.strftime("%Y-W%W")
    flag_file = os.path.join(r"C:\Users\mjc5422\.gemini\tmp\mjc5422", f"deploy_done_{week_id}.txt")

    if os.path.exists(flag_file):
        return

    target_date = get_target_day(today)

    if today == target_date and now.hour >= 8:
        show_toast("新岡さん、週報展開忘れずに。終わったら私に「週報展開済」と伝えてください。")

if __name__ == "__main__":
    main()

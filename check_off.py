import time
import re
from datetime import datetime, timedelta
import win32com.client
from plyer import notification

OVERTIME_SUBJECT = "滿公司規定工時尚未下班刷卡通知"

def get_target_folder(outlook):
    """優先尋找名為『打卡信件』的資料夾，找不到則回傳收件匣"""
    inbox = outlook.GetDefaultFolder(6)
    try:
        return inbox.Folders["打卡信件"]
    except Exception:
        try:
            # 嘗試在根目錄尋找
            for folder in outlook.Folders:
                try:
                    return folder.Folders["打卡信件"]
                except Exception:
                    continue
        except Exception:
            pass
    return inbox

def parse_time_from_subject(subject, target_date):
    """
    從主旨解析打卡時間，例如：'2026/09/14 於08:44刷卡'
    若解析成功則回傳該 datetime 物件
    """
    match = re.search(r'於\s*(\d{1,2}):(\d{2})', subject)
    if match:
        hour, minute = int(match.group(1)), int(match.group(2))
        return datetime(target_date.year, target_date.month, target_date.day, hour, minute, 0)
    return None

def check_clock_in_for_day(folder, today):
    """抓取當日早上 12:00 前的上班打卡信（依主旨時間解析）"""
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)
    
    for message in messages:
        try:
            received_time = message.ReceivedTime
            if hasattr(received_time, 'tzinfo') and received_time.tzinfo is not None:
                received_time = received_time.replace(tzinfo=None)
            
            if received_time.date() == today:
                subject = str(message.Subject)
                if "刷卡" in subject and "未下班" not in subject:
                    dt = parse_time_from_subject(subject, today)
                    # 早上 12:00 前的刷卡視為上班卡
                    if dt and dt.hour < 12:
                        return dt
        except Exception:
            continue
    return None

def has_clock_out_email(folder, today, clock_in_time):
    """檢查是否已經收到當日的下班打卡信"""
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)
    
    for message in messages:
        try:
            received_time = message.ReceivedTime
            if hasattr(received_time, 'tzinfo') and received_time.tzinfo is not None:
                received_time = received_time.replace(tzinfo=None)
            
            if received_time.date() == today:
                subject = str(message.Subject)
                if "刷卡" in subject and "未下班" not in subject:
                    dt = parse_time_from_subject(subject, today)
                    if dt and dt > clock_in_time and dt.hour >= 12:
                        return True
        except Exception:
            continue
    return False

def has_overtime_email(folder, check_date):
    """檢查特定日期是否有收到滿工時尚未刷卡的超時信件"""
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)
    
    for message in messages:
        try:
            received_time = message.ReceivedTime
            if hasattr(received_time, 'tzinfo') and received_time.tzinfo is not None:
                received_time = received_time.replace(tzinfo=None)
            
            if received_time.date() == check_date:
                if OVERTIME_SUBJECT in str(message.Subject):
                    return True
        except Exception:
            continue
    return False

def run_daily_cycle():
    now = datetime.now()
    today = now.date()

    # 週末（六日）不執行主要打卡監控
    if now.weekday() >= 5:
        print(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] 今天是週末，休眠至隔日...")
        time.sleep(3600)
        return

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = get_target_folder(outlook)

    # 1. 檢查前一個工作天是否有超時通知信，並於早上 09:00 進行補登提醒
    prev_days = 3 if now.weekday() == 0 else 1  # 若今天是週一，往前推三天（週五）
    prev_work_date = today - timedelta(days=prev_days)
    overtime_notified = False

    # 2. 早上 08:30 前不執行上班信偵測
    earliest_check_time = datetime(today.year, today.month, today.day, 8, 30, 0)
    while datetime.now() < earliest_check_time:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 尚未到達 08:30，稍候中...")
        time.sleep(60)

    # 3. 等待當日上班打卡信
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 開始偵測當日上班打卡信...")
    clock_in_time = None
    while not clock_in_time:
        clock_in_time = check_clock_in_for_day(folder, today)
        if not clock_in_time:
            time.sleep(60)
        
        # 途中若到達早上 09:00，檢查是否需要提醒補登超時
        curr = datetime.now()
        remind_time = datetime(today.year, today.month, today.day, 9, 0, 0)
        if not overtime_notified and curr >= remind_time and curr < remind_time + timedelta(minutes=10):
            if has_overtime_email(folder, prev_work_date):
                print("【提醒】前一工作天有收到超時通知，發送補登提醒！")
                try:
                    notification.notify(
                        title='⚠️ 忘刷/超時補登提醒',
                        message=f'您在 {prev_work_date.strftime("%m/%d")} 有收到滿工時尚未下班通知，請記得於系統補登超時！',
                        timeout=20
                    )
                except Exception as e:
                    print(f"發送補登通知失敗: {e}")
            overtime_notified = True

    # 4. 上班時間 + 9 小時（8小時工時 + 1小時午休）
    target_time = clock_in_time + timedelta(hours=9)
    print(f"【成功抓取】上班時間: {clock_in_time.strftime('%H:%M:%S')} (主旨時間)")
    print(f"【預計提醒】下班提醒開始時間: {target_time.strftime('%H:%M:%S')}")

    # 5. 等待到達下班時間
    while datetime.now() < target_time:
        # 若在下班時間前已提早刷卡打卡，則直接結束今日下班監控
        if has_clock_out_email(folder, today, clock_in_time):
            print("【提早打卡】已偵測到下班打卡信，今日提醒結束。")
            break
        time.sleep(30)

    # 6. 下班提醒階段：每分鐘提醒一次，最多持續 10 分鐘（共 10 次）
    alert_start = target_time
    alert_end = alert_start + timedelta(minutes=10)
    last_notify_ts = 0

    print(f"[{datetime.now().strftime('%H:%M:%S')}] 進入下班提醒階段（限時 10 分鐘）...")
    while datetime.now() <= alert_end:
        # 檢查是否已收到下班打卡信
        if has_clock_out_email(folder, today, clock_in_time):
            print("【已收到下班打卡信】下班打卡完成，停止提醒！")
            break

        # 每 60 秒（1分鐘）跳出一次通知
        now_ts = time.time()
        if now_ts - last_notify_ts >= 60:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] 發送下班打卡提醒...")
            try:
                notification.notify(
                    title='⏰ 下班打卡提醒',
                    message='已經滿 9 小時囉！尚未偵測到下班打卡信，請記得去刷卡！',
                    timeout=10
                )
            except Exception as e:
                print(f"通知發送失敗: {e}")
            last_notify_ts = now_ts

        time.sleep(10)

    print("今日下班提醒流程已結束，程式將休眠至隔日早上 08:30 後再行偵測（保持背景開啟）...")

    # 7. 休眠至隔日 08:30
    tomorrow = today + timedelta(days=1)
    next_wake_time = datetime(tomorrow.year, tomorrow.month, tomorrow.day, 8, 30, 0)
    sleep_seconds = (next_wake_time - datetime.now()).total_seconds()
    if sleep_seconds > 0:
        time.sleep(sleep_seconds)

def main():
    print("=== Outlook 打卡監控服務已啟動 ===")
    while True:
        try:
            run_daily_cycle()
        except Exception as e:
            print(f"發生非預期錯誤: {e}，1 分鐘後重試...")
            time.sleep(60)

if __name__ == "__main__":
    main()
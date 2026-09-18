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
    從主旨解析打卡時間（格式如：'2026/09/14 於08:44刷卡'）
    解析成功則回傳該 datetime 物件
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
                    if dt and dt.hour < 12:
                        return dt
        except Exception:
            continue
    return None

def has_clock_out_email(folder, today, clock_in_time):
    """檢查是否已經收到當日的下班打卡信（依主旨時間解析，需晚於上班時間且在 12:00 之後）"""
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

    # 週末（六、日）不執行監控，休眠 1 小時後再次確認
    if now.weekday() >= 5:
        print(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] 今天是週末，休眠中...")
        time.sleep(3600)
        return

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = get_target_folder(outlook)

    # 前一個工作日判定：週一檢查週五（差 3 天），其餘往前推 1 天
    prev_days = 3 if now.weekday() == 0 else 1
    prev_work_date = today - timedelta(days=prev_days)
    overtime_notified = False

    # 1. 早上 08:30 前進入待機狀態
    earliest_check_time = datetime(today.year, today.month, today.day, 8, 30, 0)
    while datetime.now() < earliest_check_time:
        time.sleep(60)

    # 2. 等待當日上班打卡信
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 開始偵測當日上班打卡信...")
    clock_in_time = None
    while not clock_in_time:
        clock_in_time = check_clock_in_for_day(folder, today)
        if not clock_in_time:
            time.sleep(60)
        
        # 早上 09:00 檢查前一工作天是否有超時通知
        curr = datetime.now()
        remind_time = datetime(today.year, today.month, today.day, 9, 0, 0)
        if not overtime_notified and curr >= remind_time and curr < remind_time + timedelta(minutes=10):
            if has_overtime_email(folder, prev_work_date):
                print("【提醒】前一工作天收到超時通知，發送補登提醒！")
                try:
                    notification.notify(
                        title='⚠️ 忘刷/超時補登提醒',
                        message=f'您在 {prev_work_date.strftime("%m/%d")} 有收到滿工時尚未下班通知，請記得補登超時！',
                        timeout=20
                    )
                except Exception as e:
                    print(f"發送補登通知失敗: {e}")
            overtime_notified = True

    # 3. 推算下班時間（上班打卡時間 + 9 小時）
    target_time = clock_in_time + timedelta(hours=9)
    print(f"【成功抓取】上班時間: {clock_in_time.strftime('%H:%M:%S')} (主旨時間)")
    print(f"【預計提醒】下班提醒開始時間: {target_time.strftime('%H:%M:%S')}")

    # 4. 等候滿 9 小時（若提前收到下班卡信件則直接標記完成）
    clock_out_done = False
    while datetime.now() < target_time:
        if has_clock_out_email(folder, today, clock_in_time):
            print("【提早打卡】已偵測到下班打卡信，今日提醒流程結束。")
            clock_out_done = True
            break
        time.sleep(30)

    # 5. 下班提醒階段：未打卡則每 60 秒提醒一次，上限 10 分鐘（共 10 次）
    if not clock_out_done:
        alert_end = target_time + timedelta(minutes=10)
        last_notify_ts = 0

        print(f"[{datetime.now().strftime('%H:%M:%S')}] 進入下班提醒階段（最多持續 10 分鐘）...")
        while datetime.now() <= alert_end:
            if has_clock_out_email(folder, today, clock_in_time):
                print("【已收到下班打卡信】下班打卡完成，停止提醒！")
                clock_out_done = True
                break

            now_ts = time.time()
            if now_ts - last_notify_ts >= 60:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] 發送下班打卡提醒...")
                try:
                    notification.notify(
                        title='⏰ 下班打卡提醒',
                        message='工作已滿 9 小時！尚未偵測到下班打卡信，請記得刷卡！',
                        timeout=10
                    )
                except Exception as e:
                    print(f"通知發送失敗: {e}")
                last_notify_ts = now_ts

            time.sleep(10)

    print("今日打卡與提醒流程結束，常駐等待至隔日 08:30（請勿關閉視窗）...")

    # 6. 安全休眠至隔日早上 08:30（支援電腦休眠喚醒）
    tomorrow = today + timedelta(days=1)
    next_wake_time = datetime(tomorrow.year, tomorrow.month, tomorrow.day, 8, 30, 0)
    
    while datetime.now() < next_wake_time:
        time.sleep(30)  # 每 30 秒檢查一次當前時間，休眠喚醒後能立刻接軌

def main():
    print("=== Outlook 打卡監控服務已啟動 ===")
    while True:
        try:
            run_daily_cycle()
        except Exception as e:
            print(f"執行時發生非預期錯誤: {e}，60 秒後重試...")
            time.sleep(60)

if __name__ == "__main__":
    main()
import time
from datetime import datetime, timedelta
import win32com.client
from plyer import notification

def main():
    # 1. 檢查今天是否為週末 (5 = 星期六, 6 = 星期日)
    today_weekday = datetime.now().weekday()
    if today_weekday >= 5:
        print("今天是週末（六日），不執行打卡提醒，程式自動結束。")
        return

    print("開始監控 Outlook 上班打卡信件...")

    def get_clock_in_time():
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6) # 收件匣
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)
            
            today = datetime.now().date()
            for message in messages:
                try:
                    received_time = message.ReceivedTime
                    if received_time.date() == today:
                        sender = message.SenderName
                        subject = message.Subject
                        if ("PIC考勤系統" in sender or "考勤" in sender or "刷卡" in subject) and "下班" not in subject:
                            if hasattr(received_time, 'tzinfo') and received_time.tzinfo is not None:
                                received_time = received_time.replace(tzinfo=None)
                            return received_time
                except:
                    continue
        except Exception as e:
            print(f"讀取 Outlook 錯誤: {e}")
        return None

    # 2. 迴圈等待直到收到今天的上班打卡信
    clock_in_time = None
    while not clock_in_time:
        clock_in_time = get_clock_in_time()
        if not clock_in_time:
            print("尚未收到今天的上班打卡信，10 分鐘後重試...")
            time.sleep(600)

    target_time = clock_in_time + timedelta(hours=8)
    print(f"【成功抓取】上班時間: {clock_in_time.strftime('%H:%M:%S')}")
    print(f"【預計提醒】下班打卡時間: {target_time.strftime('%H:%M:%S')}")

    # 3. 進入計時，直到滿 8 小時後開始每 2 分鐘提醒
    last_notify_time = 0
    task_finished = False
    
    while True:
        now = datetime.now()
        
        # 4. 檢查是否已經收到「下班打卡」的信
        if not task_finished and check_clock_out_received(clock_in_time):
            print("【偵測到下班打卡信】打卡完成！提醒已解除，程式將保持開啟（可手動關閉終端機）。")
            task_finished = True  # 改為完成狀態，停止繼續發送通知，但不結束程式

        # 5. 如果還沒打卡、超過 8 小時，且距離上次通知已經超過 2 分鐘
        if not task_finished and now >= target_time:
            current_timestamp = time.time()
            if current_timestamp - last_notify_time >= 60:
                print("時間到！發送提醒通知...")
                try:
                    notification.notify(
                        title='⏰ 下班打卡提醒',
                        message='已經工作滿 8 小時囉！尚未偵測到下班打卡信，請記得去刷卡！',
                        timeout=15
                    )
                except Exception as e:
                    print(f"發送通知失敗: {e}")
                last_notify_time = current_timestamp

        time.sleep(30) # 每 30 秒檢查一次

def check_clock_out_received(clock_in_time):
    """檢查收件匣是否有比上班時間更新、且符合下班打卡的信件"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        today = datetime.now().date()
        for message in messages:
            try:
                received_time = message.ReceivedTime
                if received_time.date() == today:
                    if hasattr(received_time, 'tzinfo') and received_time.tzinfo is not None:
                        received_time = received_time.replace(tzinfo=None)
                    
                    if received_time > clock_in_time:
                        subject = message.Subject
                        sender = message.SenderName
                        if "下班" in subject or ("考勤" in sender and "刷卡" in subject):
                            return True
            except:
                continue
    except Exception as e:
        print(f"檢查下班信件發生錯誤: {e}")
    return False

if __name__ == "__main__":
    main()
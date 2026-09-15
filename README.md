# Outlook 上下班打卡提醒小工具

監控本地 Outlook 收件匣中的上班打卡通知信，自動計算 8 小時工作時間，並於滿 8 小時後透過 Windows 桌面通知每 2 分鐘提醒一次，直到偵測到下班打卡信件為止。

## 功能特點
- 自動排除週末（六日不執行）。
- 自動抓取當天最新上班打卡時間並推算 8 小時下班點。
- 滿 8 小時且未打下班卡時，每 120 秒發送桌面彈出通知。
- 偵測到下班打卡信後停止發送通知，程式維持執行直到手動關閉。

## 環境需求
- Windows 10 / 11
- 已安裝並登入的 Outlook 桌面版
- Python 3.10+

## 安裝方式
```powershell
pip install -r requirements.txt
```

## 執行方式
```powershell
python check_off.py
```
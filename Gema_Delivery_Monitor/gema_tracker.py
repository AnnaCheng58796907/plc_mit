import pandas as pd
import requests
import shutil
import os
import sys
import configparser
from datetime import datetime, timedelta

# --- 取得 .exe 所在的資料夾路徑 ---
if getattr(sys, 'frozen', False):
    # 如果是打包後的 exe，取得 exe 所在的目錄
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 如果是原始 python 腳本，取得腳本所在的目錄
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

config_path = os.path.join(BASE_DIR, 'config.ini')
config = configparser.ConfigParser()

# 讀取設定檔
if not os.path.exists(config_path):
    print(f"找不到設定檔: {config_path}")
    input("請按任意鍵退出...")
    exit()

config.read(config_path, encoding='utf-8')

# 從外部設定檔抓取變數
SERVER_FILE = config.get('PATH', 'server_file')
TG_TOKEN = config.get('TELEGRAM', 'token')
TG_ID = config.get('TELEGRAM', 'chat_id')
REMIND_DAYS = config.getint('SETTINGS', 'remind_days', fallback=1)

def send_tg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_ID, "text": msg, "parse_mode": "Markdown"}
    try:
        requests.post(url, data=payload, timeout=10)
    except:
        pass

def run_task():
    temp_local = os.path.join(BASE_DIR, 'temp_sync.xlsx')
    try:
        if not os.path.exists(SERVER_FILE):
            print(f"無法存取路徑: {SERVER_FILE}")
            return
        
        shutil.copy2(SERVER_FILE, temp_local)
        df = pd.read_excel(temp_local)
        
        # 篩選邏輯 (與測試時相同)
        df['預計回廠'] = pd.to_datetime(df['預計回廠'], errors='coerce')
        today = datetime.now().date()
        target = today + timedelta(days=REMIND_DAYS)
        
        alert_df = df[(df['實到日期'].isna()) & (df['預計回廠'].dt.date <= target)]
        
        if not alert_df.empty:
            msg = "📦 *【進貨提醒】*\n"
            for _, row in alert_df.iterrows():
                d = row['預計回廠'].strftime('%m/%d')
                msg += f"📅 {d} | 🏭 {row.get('廠商','?')} | 📄 {row.get('工單號碼','?')}\n"
            send_tg(msg)
            print("通知發送成功！")
        else:
            print("目前無待追蹤項目。")
            
    except Exception as e:
        print(f"錯誤: {e}")
    finally:
        if os.path.exists(temp_local):
            os.remove(temp_local)

if __name__ == "__main__":
    run_task()
    print("程式執行完畢。")
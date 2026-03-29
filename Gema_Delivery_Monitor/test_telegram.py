import requests
import configparser
import os

# --- 設定檔讀取邏輯 ---
CONFIG_FILE = 'config.ini'
config = configparser.ConfigParser()

def check_and_send():
    # 1. 檢查設定檔是否存在
    if not os.path.exists(CONFIG_FILE):
        # 如果不存在，自動建立一個「範本」
        config['TELEGRAM'] = {
            'token': '在此填入Token',
            'chat_id': '在此填入數字ID'
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            config.write(f)
        print(f"首次執行：已產生 {CONFIG_FILE} 檔案。")
        print("💡 請到資料夾中用『記事本』打開該檔案，填入你的資訊後儲存，再重新執行本程式。")
        return

    # 2. 讀取設定檔
    config.read(CONFIG_FILE, encoding='utf-8')
    TOKEN = config.get('TELEGRAM', 'token')
    ID = config.get('TELEGRAM', 'chat_id')

    # 3. 測試發送
    print("正在嘗試連線至 Telegram...")
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    payload = {
        'chat_id': ID,
        'text': '🚀 測試連線成功！這是我從外部設定檔讀取的訊息。'
    }

    try:
        response = requests.post(url, data=payload)
        if response.status_code == 200:
            print("✅ 成功！請查看手機 Telegram 機器人。")
        else:
            print(f"❌ 失敗！錯誤代碼：{response.status_code}")
            print(f"內容：{response.text}")
    except Exception as e:
        print(f"💥 無法連線，可能是網路環境限制：{e}")

if __name__ == "__main__":
    check_and_send()
import pymcprotocol
import pandas as pd
import time
from datetime import datetime

# 設定test_PLC連線參數
# 請將IP位址修改為你的FX3U/FX5U PLC實際的IP位址
# 埠號通常為5007 (用於MC Protocol)
PLC_IP = "127.0.0.1"  # 範例IP，請自行修改
PLC_PORT = 5007  # GX Works2模擬器預設埠號

# 設定要讀取的資料區域和數量
# 例如，這裡讀取D1000到D1009，共10個暫存器
READ_DEVICE = "D1000"
READ_COUNT = 10

# 設定CSV檔案儲存路徑
# 檔案會以"plc_data_YYYY-MM-DD_HH-MM-SS.csv"命名
CSV_FOLDER = "D:\GitHub\plc_mit\pymcprotocol_test"  # 請自行修改為你想要儲存的資料夾路徑


def connect_and_read_plc():
    """
    連線到PLC並讀取指定資料
    """
    try:
        # 建立MC Protocol通訊物件
        mc = pymcprotocol.Type3E()

        print(f"嘗試連線至PLC：{PLC_IP}:{PLC_PORT}...")
        mc.connect(PLC_IP, PLC_PORT)
        print("連線成功！")

        # 讀取資料
        # 'd' 代表D暫存器，'1000' 是起始位址，'10' 是讀取數量
        print(f"正在讀取 {READ_DEVICE} 開始的 {READ_COUNT} 個暫存器...")
        read_values = mc.batchread_wordunits(headdevice=READ_DEVICE,
                                              readsize=READ_COUNT)
        print("讀取完成！")
        # 關閉連線
        mc.close()
        print("已中斷PLC連線。")
        return read_values
        
    except Exception as e:
        print(f"連線或讀取PLC時發生錯誤：{e}")
        return None


def save_to_csv(data_list):
    """
    將讀取的資料儲存為CSV檔案
    """
    if data_list is None:
        print("沒有資料可以儲存。")
        return
        
    try:
        # 建立一個DataFrame
        # 這裡假設讀取的是D1000到D1009
        columns = [f"D{1000 + i}" for i in range(len(data_list))]
        df = pd.DataFrame([data_list], columns=columns)
        
        # 生成帶有時間戳記的檔案名稱
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"plc_data_{timestamp}.csv"
        filepath = CSV_FOLDER + "\\" + filename

        # 將DataFrame儲存為CSV檔案
        df.to_csv(filepath, index=False)
        
        print(f"資料已成功儲存至：{filepath}")
        
    except Exception as e:
        print(f"儲存CSV檔案時發生錯誤：{e}")


# 主程式區塊
if __name__ == "__main__":
    
    # 建立CSV檔案儲存資料夾
    # 這行程式碼會確保指定的資料夾存在，如果不存在會自動建立
    import os
    os.makedirs(CSV_FOLDER, exist_ok=True)
    
    # 開始執行主迴圈，每隔一段時間讀取一次資料
    while True:
        print("\n------------------------------------")
        
        # 呼叫函數連線並讀取PLC資料
        plc_data = connect_and_read_plc()
        
        # 如果成功讀取到資料，則呼叫函數儲存為CSV
        if plc_data is not None:
            save_to_csv(plc_data)
        
        # 設定每次讀取資料的間隔時間
        # 這裡設定每5秒執行一次，你可以根據需求調整
        INTERVAL_SECONDS = 5
        print(f"等待 {INTERVAL_SECONDS} 秒後進行下一次讀取...")
        time.sleep(INTERVAL_SECONDS)
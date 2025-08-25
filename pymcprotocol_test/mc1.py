import pymcprotocol
import pandas as pd
import time
from datetime import datetime
import os

# 設定test_PLC連線參數
# 請將IP位址修改為你的FX3U/FX5U PLC實際的IP位址
# 埠號通常為5007 (用於MC Protocol)
PLC_IP = "127.0.0.1"
# 範例IP，請自行修改
PLC_PORT = 5007
# GX Works2模擬器預設埠號

# 設定要讀取的資料區域和數量
# D暫存器: 讀取D1000到D1009，共10個
READ_WORD_DEVICE = "D1000"
READ_WORD_COUNT = 10

# M點位: 讀取M1000到M1009，共10個
READ_BIT_DEVICE = "M1000"
READ_BIT_COUNT = 10

# 設定CSV檔案儲存路徑
# 檔案會以"plc_data_YYYY-MM-DD_HH-MM-SS.csv"命名
CSV_FOLDER = r"D:\GitHub\plc_mit\pymcprotocol_test"  # 請自行修改為你想要儲存的資料夾路徑


def connect_and_read_plc():
    """
    連線到PLC並讀取指定資料
    """
    try:
        # 建立MC Protocol通訊物件
        mc = pymcprotocol.Type3E()
        
        # 設定連線參數
        mc.setaccessopt(commtype="binary")
        
        print(f"嘗試連線至PLC：{PLC_IP}:{PLC_PORT}...")
        mc.connect(ip=PLC_IP, port=PLC_PORT)
        print("連線成功！")

        # 讀取D暫存器資料
        print(f"正在讀取 {READ_WORD_DEVICE} 開始的 {READ_WORD_COUNT} 個暫存器...")
        word_values = mc.batchread_wordunits(headdevice=READ_WORD_DEVICE,
                                             readsize=READ_WORD_COUNT)
        print("D暫存器讀取完成！")

        # 讀取M點位資料
        print(f"正在讀取 {READ_BIT_DEVICE} 開始的 {READ_BIT_COUNT} 個點位...")
        bit_values = mc.batchread_bitunits(headdevice=READ_BIT_DEVICE,
                                           readsize=READ_BIT_COUNT)
        print("M點位讀取完成！")

        # 關閉連線
        mc.close()
        print("已中斷PLC連線。")

        # 返回一個字典，包含兩種資料
        return {"word_data": word_values, "bit_data": bit_values}
 
    except Exception as e:
        print(f"連線或讀取PLC時發生錯誤：{e}")
        return None


def save_to_csv(data_dict):
    """
    將讀取的資料儲存為CSV檔案
    """
    if data_dict is None:
        print("沒有資料可以儲存。")
        return

    try:
        word_data = data_dict.get("word_data", [])
        bit_data = data_dict.get("bit_data", [])

        # 建立欄位名稱
        word_columns = [f"D{1000 + i}" for i in range(len(word_data))]
        bit_columns = [f"M{1000 + i}" for i in range(len(bit_data))]
        all_columns = word_columns + bit_columns

        # 將兩種資料合併為一行
        all_data = word_data + bit_data

        # 建立一個DataFrame
        df = pd.DataFrame([all_data], columns=all_columns)

        # 生成帶有時間戳記的檔案名稱
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"plc_data_{timestamp}.csv"
        filepath = os.path.join(CSV_FOLDER, filename)

        # 將DataFrame儲存為CSV檔案
        df.to_csv(filepath, index=False)

        print(f"資料已成功儲存至：{filepath}")

    except Exception as e:
        print(f"儲存CSV檔案時發生錯誤：{e}")


# 主程式區塊
if __name__ == "__main__":

    # 建立CSV檔案儲存資料夾
    os.makedirs(CSV_FOLDER, exist_ok=True)

    # 開始執行主迴圈，每隔一段時間讀取一次資料
    while True:
        print("\n------------------------------------")

        # 呼叫函數連線並讀取PLC資料
        plc_data_dict = connect_and_read_plc()

        # 如果成功讀取到資料，則呼叫函數儲存為CSV
        if plc_data_dict is not None:
            save_to_csv(plc_data_dict)

        # 設定每次讀取資料的間隔時間
        INTERVAL_SECONDS = 5
        print(f"等待 {INTERVAL_SECONDS} 秒後進行下一次讀取...")
        time.sleep(INTERVAL_SECONDS)
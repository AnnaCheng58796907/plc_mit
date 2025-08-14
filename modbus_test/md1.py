import pandas as pd
import time
from datetime import datetime
from pymodbus.client import ModbusTcpClient

# ----------------------------------------------------
# 設定PLC連線參數
# ----------------------------------------------------
PLC_IP = "192.168.1.100"  # 請修改為你的FX PLC實際IP位址
PLC_PORT = 502             # Modbus TCP的標準埠號
MODBUS_ADDRESS = 1000      # Modbus的起始位址，例如讀取D1000，請根據手冊調整
READ_COUNT = 10            # 要讀取的暫存器數量

# ----------------------------------------------------
# 設定CSV檔案儲存路徑
# ----------------------------------------------------
CSV_FOLDER = "C:\\PLC_Data_Modbus\\"  # 請自行修改為你想要儲存的資料夾路徑


def connect_and_read_plc():
    """
    連線到PLC並讀取指定資料
    """
    try:
        # 建立Modbus TCP客戶端物件
        client = ModbusTcpClient(PLC_IP, port=PLC_PORT)
        
        print(f"嘗試連線至PLC：{PLC_IP}:{PLC_PORT}...")
        if client.connect():
            print("連線成功！")
        else:
            print("連線失敗，請檢查IP、埠號或網路設定。")
            return None

        # 讀取Holding Registers
        print(f"正在讀取 Modbus 位址 {MODBUS_ADDRESS} 開始的 {READ_COUNT} 個暫存器...")
        # read_holding_registers(address, count, unit)
        # unit=1 是 Modbus Slave ID，通常為1
        result = client.read_holding_registers(address=MODBUS_ADDRESS,
                                               count=READ_COUNT, slave=1)

        if result.isError():
            print(f"讀取資料時發生錯誤：{result}")
            client.close()
            return None
        
        read_values = result.registers
        print(f"讀取完成！資料：{read_values}")
        
        # 關閉連線
        client.close()
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
        # 這裡假設讀取的資料對應到D暫存器
        # 根據你的PLC暫存器調整
        columns = [f"D{1000 + i}" for i in range(len(data_list))]
        df = pd.DataFrame([data_list], columns=columns)
        
        # 生成帶有時間戳記的檔案名稱
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"plc_data_{timestamp}.csv"
        filepath = CSV_FOLDER + filename

        # 將DataFrame儲存為CSV檔案
        df.to_csv(filepath, index=False)
        
        print(f"資料已成功儲存至：{filepath}")
        
    except Exception as e:
        print(f"儲存CSV檔案時發生錯誤：{e}")


# 主程式區塊
if __name__ == "__main__":
    
    # 建立CSV檔案儲存資料夾
    import os
    os.makedirs(CSV_FOLDER, exist_ok=True)
    
    # 開始執行主迴圈
    while True:
        print("\n" + "="*50)
        
        plc_data = connect_and_read_plc()
        
        if plc_data is not None:
            save_to_csv(plc_data)
        
        INTERVAL_SECONDS = 5
        print(f"等待 {INTERVAL_SECONDS} 秒後進行下一次讀取...")
        time.sleep(INTERVAL_SECONDS)
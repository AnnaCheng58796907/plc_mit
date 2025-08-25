import pymcprotocol
import pandas as pd
import time
from datetime import datetime
import os

# 設定PLC連線參數
PLC_IP = "127.0.0.1"
PLC_PORT = 5007

# 設定要讀取的資料區域和數量 - 改為從較低的位址開始
READ_WORD_DEVICE = "D0"  # 改為D0，更可能存在
READ_WORD_COUNT = 10

READ_BIT_DEVICE = "M0"  # 改為M0，更可能存在
READ_BIT_COUNT = 10

# 設定CSV檔案儲存路徑
CSV_FOLDER = r"D:\GitHub\plc_mit\pymcprotocol_test"


def test_basic_communication(mc):
    """測試基本通訊功能"""
    print("🔍 測試基本通訊...")
    
    # 測試 echo
    try:
        mc.echo_test()
        print("✅ Echo 測試成功")
        return True
    except Exception as e:
        print(f"❌ Echo 測試失敗: {e}")
        return False


def test_different_addresses(mc):
    """測試不同的暫存器位址"""
    test_addresses = [
        ("D0", 1),
        ("D10", 1), 
        ("D100", 1),
        ("D1000", 1)
    ]
    
    print("🔍 測試不同暫存器位址...")
    successful_address = None
    
    for addr, size in test_addresses:
        try:
            print(f"  測試 {addr}...")
            data = mc.batchread_wordunits(headdevice=addr, readsize=size)
            print(f"  ✅ {addr} 讀取成功: {data}")
            if successful_address is None:
                successful_address = addr
        except Exception as e:
            print(f"  ❌ {addr} 讀取失敗: {e}")
    
    return successful_address


def connect_and_read_plc():
    """
    連線到PLC並讀取指定資料
    """
    mc = None
    try:
        # 建立MC Protocol通訊物件
        mc = pymcprotocol.Type3E()
        
        # 設定通訊參數
        mc.setaccessopt(commtype="binary")
        
        # 增加超時時間到10秒
        mc.soc_timeout = 10.0
        
        print(f"嘗試連線至PLC：{PLC_IP}:{PLC_PORT}...")
        mc.connect(ip=PLC_IP, port=PLC_PORT)
        print("✅ 連線成功！")
        
        # 測試基本通訊
        if not test_basic_communication(mc):
            print("⚠️  基本通訊測試失敗，但繼續嘗試讀取資料...")
        
        # 測試不同位址
        working_address = test_different_addresses(mc)
        
        if working_address:
            print(f"\n✅ 找到可用的暫存器位址: {working_address}")
            
            # 使用可用的位址讀取更多資料
            try:
                print(f"讀取 {working_address} 開始的 {READ_WORD_COUNT} 個暫存器...")
                word_values = mc.batchread_wordunits(
                    headdevice=working_address,
                    readsize=READ_WORD_COUNT
                )
                print(f"✅ D暫存器讀取成功: {word_values}")
                
                # 嘗試讀取M點位
                try:
                    print(f"讀取 {READ_BIT_DEVICE} 開始的 {READ_BIT_COUNT} 個點位...")
                    bit_values = mc.batchread_bitunits(
                        headdevice=READ_BIT_DEVICE,
                        readsize=READ_BIT_COUNT
                    )
                    print(f"✅ M點位讀取成功: {bit_values}")
                except Exception as e:
                    print(f"❌ M點位讀取失敗: {e}")
                    bit_values = []
                
                return {"word_data": word_values, "bit_data": bit_values}
                
            except Exception as e:
                print(f"❌ 批量讀取失敗: {e}")
        else:
            print("❌ 找不到可用的暫存器位址")
            
        return None
        
    except Exception as e:
        print(f"❌ 連線或讀取PLC時發生錯誤：{e}")
        return None
    finally:
        if mc:
            try:
                mc.close()
                print("🔌 已中斷PLC連線")
            except:
                pass


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
        word_columns = [f"D{i}" for i in range(len(word_data))]
        bit_columns = [f"M{i}" for i in range(len(bit_data))]
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

        print(f"✅ 資料已成功儲存至：{filepath}")

    except Exception as e:
        print(f"❌ 儲存CSV檔案時發生錯誤：{e}")


# 主程式區塊
if __name__ == "__main__":
    
    # 建立CSV檔案儲存資料夾
    os.makedirs(CSV_FOLDER, exist_ok=True)
    
    print("=== PLC 連線診斷工具 ===")
    print("按 Ctrl+C 停止程式")
    
    success_count = 0
    total_attempts = 0
    
    try:
        # 執行一次完整的測試
        print("\n" + "="*50)
        total_attempts += 1
        
        # 呼叫函數連線並讀取PLC資料
        plc_data_dict = connect_and_read_plc()
        
        # 如果成功讀取到資料，則呼叫函數儲存為CSV
        if plc_data_dict is not None:
            success_count += 1
            save_to_csv(plc_data_dict)
            print(f"\n✅ 成功率: {success_count}/{total_attempts}")
            print("診斷完成！PLC 通訊正常。")
        else:
            print(f"\n❌ 成功率: {success_count}/{total_attempts}")
            print("診斷結果：PLC 通訊有問題，請檢查：")
            print("1. GX Works2 模擬器是否正確啟動")
            print("2. MC Protocol 設定是否正確")
            print("3. 暫存器位址範圍是否存在")
            
    except KeyboardInterrupt:
        print("\n\n程式已停止")
    except Exception as e:
        print(f"\n程式執行錯誤：{e}")

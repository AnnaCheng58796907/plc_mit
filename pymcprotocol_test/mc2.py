import pymcprotocol

# # 步驟 1: 建立連線物件
# # 連接到 GX Works2 模擬器，使用預設的 IP 和 Port
# # 'SLMP' 代表三菱的 SLMP (Seamless Message Protocol)，是 MC 協定的底層協定
# mc = pymcprotocol.Type3E('127.0.0.1', 5007)

# # 步驟 2: 連線到 PLC
# # 連線到模擬 PLC
# try:
#     # 這裡的 'connect()' 會嘗試建立連線
#     mc.connect()
#     print("成功連線到 GX Works2 模擬器！")

#     # 步驟 3: 讀取暫存器資料
#     # 使用 'batch_read_word' 來讀取一組字元 (Word)
#     # 這裡的參數意義如下：
#     # - 'D': 代表要讀取的暫存器類型是 D 暫存器
#     # - 0: 從 D0 暫存器開始讀取
#     # - 10: 總共要讀取 10 個暫存器 (從 D0 到 D9)
#     d_data = mc.batch_read_word('D', 0, 10)

#     # 步驟 4: 顯示讀取到的資料
#     print("讀取到的 D0 至 D9 暫存器資料為：")
#     for i, data in enumerate(d_data):
#         print(f"D{i}: {data}")

#     # 步驟 5: 寫入資料到暫存器
#     # 使用 'batch_write_word' 來寫入一組字元
#     # 這裡的參數意義如下：
#     # - 'D': 代表要寫入的暫存器類型
#     # - 10: 從 D10 暫存器開始寫入
#     # - [100, 200]: 要寫入的資料，100 會寫入 D10，200 會寫入 D11
#     mc.batch_write_word('D', 10, [100, 200])
#     print("\n成功寫入資料 [100, 200] 到 D10 和 D11 暫存器。")

#     # 再次讀取 D10 到 D11，以確認資料是否正確寫入
#     d_data_written = mc.batch_read_word('D', 10, 2)
#     print("再次讀取 D10 至 D11 暫存器，確認資料：")
#     for i, data in enumerate(d_data_written):
#         print(f"D{i+10}: {data}")

# except Exception as e:
#     print(f"連線或讀寫錯誤: {e}")

# finally:
#     # 步驟 6: 關閉連線
#     # 確保連線在程式結束時被正常關閉
#     if 'mc' in locals() and mc.is_connect():
#         mc.close()
#         print("\n連線已關閉。")

# 這行程式碼不應該有錯

mc = pymcprotocol.Type3E()

try:
    mc.connect('127.0.0.1', 5007)
    print("成功連線到 PLC！")
    # 進行您的讀寫操作
    d_data = mc.batchread_wordunits(headdevice='D0', readsize=10)
    print("讀取到的資料:", d_data)

except Exception as e:
    print(f"連線失敗: {e}")
finally:
    try:
        mc.close()
        print("連線已關閉。")
    except:
        pass

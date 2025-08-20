import pymcprotocol

def test_connection():
    """簡單測試PLC連線"""
    mc = pymcprotocol.Type3E()
    
    try:
        print("嘗試連線到 GX Works2 模擬器...")
        mc.connect('127.0.0.1', 5007)
        print("✅ 連線成功！")
        
        # 嘗試讀取簡單的D0暫存器
        print("嘗試讀取 D0...")
        try:
            data = mc.batchread_wordunits(headdevice='D0', readsize=1)
            print(f"✅ 成功讀取 D0: {data}")
        except Exception as e:
            print(f"❌ 讀取D0失敗: {e}")
            
        # 嘗試讀取 D100
        print("嘗試讀取 D100...")
        try:
            data = mc.batchread_wordunits(headdevice='D100', readsize=1)
            print(f"✅ 成功讀取 D100: {data}")
        except Exception as e:
            print(f"❌ 讀取D100失敗: {e}")
            
        # 嘗試讀取 D1000
        print("嘗試讀取 D1000...")
        try:
            data = mc.batchread_wordunits(headdevice='D1000', readsize=1)
            print(f"✅ 成功讀取 D1000: {data}")
        except Exception as e:
            print(f"❌ 讀取D1000失敗: {e}")
            
    except Exception as e:
        print(f"❌ 連線失敗: {e}")
    finally:
        try:
            mc.close()
            print("🔌 連線已關閉")
        except:
            pass

if __name__ == "__main__":
    test_connection()

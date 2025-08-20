import pymcprotocol

def test_with_timeout():
    """測試不同的超時設定"""
    mc = pymcprotocol.Type3E()
    
    try:
        print("嘗試連線到 GX Works2 模擬器...")
        mc.connect('127.0.0.1', 5007)
        print("✅ 連線成功！")
        
        # 設定較長的超時時間（預設是3秒，改為10秒）
        mc.soc_timeout = 10.0
        print(f"設定超時時間為: {mc.soc_timeout} 秒")
        
        # 嘗試讀取不同的暫存器
        test_addresses = ['D0', 'D10', 'D100', 'D500', 'D1000']
        
        for addr in test_addresses:
            print(f"\n嘗試讀取 {addr}...")
            try:
                data = mc.batchread_wordunits(headdevice=addr, readsize=1)
                print(f"✅ 成功讀取 {addr}: {data}")
                break  # 如果成功讀取，就停止測試
            except Exception as e:
                print(f"❌ 讀取{addr}失敗: {e}")
                
        # 嘗試echo測試
        print("\n嘗試 Echo 測試...")
        try:
            mc.echo_test()
            print("✅ Echo 測試成功")
        except Exception as e:
            print(f"❌ Echo 測試失敗: {e}")
            
    except Exception as e:
        print(f"❌ 連線失敗: {e}")
    finally:
        try:
            mc.close()
            print("\n🔌 連線已關閉")
        except:
            pass

if __name__ == "__main__":
    test_with_timeout()

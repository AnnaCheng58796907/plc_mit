import tkinter as tk
from tkinter import messagebox
import math

def calculate():
    try:
        # 讀取輸入數值
        t = float(entry_t.get())        # 板厚
        r = float(entry_r.get())        # 內半徑
        angle = float(entry_a.get())    # 折彎角度
        flange_a = float(entry_fa.get()) # 邊長 A
        flange_b = float(entry_fb.get()) # 邊長 B
        k = float(entry_k.get())        # K-Factor
        
        # 讀取開孔數據 (選填)
        hole_d = entry_hd.get()
        hole_x = entry_hx.get()

        # 核心計算：折彎補償 (BA)
        rad = math.radians(angle)
        ba = rad * (r + (k * t))
        
        # 總展開長度 L = (A - R - T) + (B - R - T) + BA
        # 註：此處假設 A, B 為外部尺寸
        total_l = (flange_a - r - t) + (flange_b - r - t) + ba
        
        result_text = f"折彎補償 (BA): {ba:.3f} mm\n總展開長度 (L): {total_l:.3f} mm"
        
        # 開孔安全性檢查
        if hole_d and hole_x:
            h_d = float(hole_d)
            h_x = float(hole_x)
            safe_dist = 2 * t
            actual_dist = h_x - (h_d / 2)
            if actual_dist < safe_dist:
                result_text += f"\n\n⚠ 警告：孔位太近！\n建議安全距離需 > {safe_dist}mm"
            else:
                result_text += "\n\n✅ 開孔位置安全"

        label_result.config(text=result_text, fg="#007ACC")
        
    except ValueError:
        messagebox.showerror("輸入錯誤", "請確保所有欄位皆為數字")

# --- 建立視窗介面 ---
root = tk.Tk()
root.title("板金展開計算器 v1.0")
root.geometry("400x550")

# 標題
tk.Label(root, text="板金參數輸入", font=("Arial", 14, "bold")).pack(pady=10)

# 建立輸入表單函式
def create_input(label_text, default_val):
    frame = tk.Frame(root)
    frame.pack(fill="x", padx=30, pady=5)
    tk.Label(frame, text=label_text, width=15, anchor="w").pack(side="left")
    entry = tk.Entry(frame)
    entry.insert(0, default_val)
    entry.pack(side="right", expand=True, fill="x")
    return entry

entry_t = create_input("板厚 (T):", "2.0")
entry_r = create_input("折彎內半徑 (R):", "2.0")
entry_a = create_input("折彎角度 (°):", "90")
entry_fa = create_input("外部邊長 A:", "50")
entry_fb = create_input("外部邊長 B:", "50")
entry_k = create_input("K-Factor:", "0.44")

tk.Label(root, text="開孔選項 (選填)", font=("Arial", 10, "italic")).pack(pady=10)
entry_hd = create_input("孔徑 (D):", "")
entry_hx = create_input("孔中心距折彎線:", "")

# 計算按鈕
btn_calc = tk.Button(root, text="開始計算", command=calculate, bg="#007ACC", fg="white", font=("Arial", 12, "bold"))
btn_calc.pack(pady=20, ipadx=20)

# 結果顯示區
label_result = tk.Label(root, text="等待計算...", font=("Arial", 11), justify="left")
label_result.pack(pady=10)

root.mainloop()
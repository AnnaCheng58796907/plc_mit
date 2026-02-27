import tkinter as tk
from tkinter import messagebox, ttk
import math

# 定義不同材料的預設 K-Factor
MATERIAL_DATA = {
    "黑鐵 (Cold Rolled Steel)": 0.44,
    "不鏽鋼 (Stainless Steel)": 0.35,
    "鋁板 (Aluminum)": 0.40,
    "自定義 (Custom)": 0.44
}

def on_material_change(event):
    """當下拉選單切換時，自動更新 K-Factor 欄位"""
    selected_material = combo_material.get()
    k_val = MATERIAL_DATA.get(selected_material, 0.44)
    entry_k.delete(0, tk.END)
    entry_k.insert(0, str(k_val))

def calculate():
    try:
        # 讀取輸入數值
        t = float(entry_t.get())        # 板厚
        r = float(entry_r.get())        # 內半徑
        angle = float(entry_a.get())    # 折彎角度
        flange_a = float(entry_fa.get()) # 邊長 A
        flange_b = float(entry_fb.get()) # 邊長 B
        k = float(entry_k.get())        # K-Factor
        
        # 核心計算：折彎補償 (BA)
        rad = math.radians(angle)
        ba = rad * (r + (k * t))
        
        # 總展開長度 L (假設 A, B 為外部尺寸)
        # 公式：L = (A - Setback) + (B - Setback) + BA
        # 90度時 Setback = R + T
        total_l = (flange_a - r - t) + (flange_b - r - t) + ba
        
        result_text = f"材料: {combo_material.get()}\n"
        result_text += f"折彎補償 (BA): {ba:.3f} mm\n"
        result_text += f"總展開長度 (L): {total_l:.3f} mm"
        
        # 開孔安全性檢查
        hole_d = entry_hd.get()
        hole_x = entry_hx.get()
        if hole_d and hole_x:
            h_d = float(hole_d)
            h_x = float(hole_x)
            safe_dist = 2 * t
            actual_dist = h_x - (h_d / 2)
            if actual_dist < safe_dist:
                result_text += f"\n\n⚠ 警告：孔位太近折彎區！\n孔邊緣距 R 起點僅 {actual_dist:.2f}mm\n(建議需 > {safe_dist}mm)"
            else:
                result_text += "\n\n✅ 開孔位置安全"

        label_result.config(text=result_text, fg="#007ACC")
        
    except ValueError:
        messagebox.showerror("輸入錯誤", "請確保所有欄位皆為數字")

# --- 建立視窗介面 ---
root = tk.Tk()
root.title("專業板金展開計算器 v1.1")
root.geometry("450x650")

# 標題
tk.Label(root, text="板金加工參數設定", font=("Arial", 14, "bold")).pack(pady=15)

# 1. 材料選擇下拉選單
frame_mat = tk.Frame(root)
frame_mat.pack(fill="x", padx=30, pady=5)
tk.Label(frame_mat, text="選擇材料:", width=15, anchor="w").pack(side="left")
combo_material = ttk.Combobox(frame_mat, values=list(MATERIAL_DATA.keys()), state="readonly")
combo_material.set("黑鐵 (Cold Rolled Steel)")
combo_material.pack(side="right", expand=True, fill="x")
combo_material.bind("<<ComboboxSelected>>", on_material_change)

# 建立輸入表單函式
def create_input(label_text, default_val):
    frame = tk.Frame(root)
    frame.pack(fill="x", padx=30, pady=5)
    tk.Label(frame, text=label_text, width=15, anchor="w").pack(side="left")
    entry = tk.Entry(frame)
    entry.insert(0, default_val)
    entry.pack(side="right", expand=True, fill="x")
    return entry

entry_t = create_input("板厚 (T) mm:", "2.0")
entry_r = create_input("折彎內半徑 (R) mm:", "2.0")
entry_a = create_input("折彎角度 (°):", "90")
entry_fa = create_input("外部邊長 A mm:", "50")
entry_fb = create_input("外部邊長 B mm:", "50")
entry_k = create_input("K-Factor:", "0.44")

# 開孔選項
tk.Label(root, text="開孔選項 (選填)", font=("Arial", 10, "italic"), fg="gray").pack(pady=10)
entry_hd = create_input("孔直徑 (D) mm:", "")
entry_hx = create_input("孔中心距折彎線 mm:", "")

# 計算按鈕
btn_calc = tk.Button(root, text="執行展開計算", command=calculate, bg="#28a745", fg="white", font=("Arial", 12, "bold"))
btn_calc.pack(pady=20, ipadx=30)

# 結果顯示區
result_frame = tk.LabelFrame(root, text="計算結果", padx=10, pady=10)
result_frame.pack(padx=30, fill="both", expand=True)
label_result = tk.Label(result_frame, text="請輸入數據後點擊計算", font=("Consolas", 10), justify="left")
label_result.pack()

root.mainloop()
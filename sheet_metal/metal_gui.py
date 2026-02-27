import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import math
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 連動板金計算器 v2.0")
        self.root.geometry("450x700")
        
        # 1. 讀取 Excel 數據
        self.load_excel_data()
        
        # 2. 建立 UI 介面
        self.create_widgets()

    def load_excel_data(self):
        file_path = "bend_parameters.xlsx"
        if not os.path.exists(file_path):
            messagebox.showerror("錯誤", f"找不到 {file_path}\n請先建立 Excel 檔。")
            self.df = pd.DataFrame(columns=["材質", "厚度", "折彎係數"])
        else:
            try:
                self.df = pd.read_excel(file_path)
            except Exception as e:
                messagebox.showerror("讀取失敗", f"無法讀取 Excel: {e}")

    def create_widgets(self):
        tk.Label(self.root, text="從 Excel 自動選用係數", font=("Arial", 14, "bold")).pack(pady=15)

        # 材質下拉選單
        tk.Label(self.root, text="1. 選擇材質:").pack()
        self.materials = self.df["材質"].unique().tolist()
        self.combo_mat = ttk.Combobox(self.root, values=self.materials, state="readonly")
        self.combo_mat.pack(pady=5)
        self.combo_mat.bind("<<ComboboxSelected>>", self.update_thickness_options)

        # 厚度下拉選單
        tk.Label(self.root, text="2. 選擇厚度 (mm):").pack()
        self.combo_thick = ttk.Combobox(self.root, state="readonly")
        self.combo_thick.pack(pady=5)
        self.combo_thick.bind("<<ComboboxSelected>>", self.update_k_factor)

        # K-Factor (自動填入)
        self.entry_k = self.create_input("自動填入 K-Factor:", "0.0")
        
        # 其他基本參數
        self.entry_r = self.create_input("內半徑 (R) mm:", "2.0")
        self.entry_a = self.create_input("折彎角度 (°):", "90")
        self.entry_fa = self.create_input("外部邊長 A mm:", "50")
        self.entry_fb = self.create_input("外部邊長 B mm:", "50")

        # 計算按鈕
        btn_calc = tk.Button(self.root, text="執行計算", command=self.calculate, bg="#007ACC", fg="white")
        btn_calc.pack(pady=20, ipadx=30)

        # 結果顯示
        self.label_result = tk.Label(self.root, text="結果顯示區", font=("Consolas", 10), justify="left")
        self.label_result.pack()

    def create_input(self, label_text, default_val):
        frame = tk.Frame(self.root)
        frame.pack(fill="x", padx=40, pady=2)
        tk.Label(frame, text=label_text, width=15, anchor="w").pack(side="left")
        entry = tk.Entry(frame)
        entry.insert(0, default_val)
        entry.pack(side="right", expand=True, fill="x")
        return entry

    def update_thickness_options(self, event):
        selected_mat = self.combo_mat.get()
        # 篩選該材質擁有的厚度
        thicknesses = self.df[self.df["材質"] == selected_mat]["厚度"].tolist()
        self.combo_thick.config(values=thicknesses)
        self.combo_thick.set("")
        self.entry_k.delete(0, tk.END)

    def update_k_factor(self, event):
        selected_mat = self.combo_mat.get()
        selected_thick = float(self.combo_thick.get())
        # 找出對應的 K-Factor
        k_val = self.df[(self.df["材質"] == selected_mat) & (self.df["厚度"] == selected_thick)]["折彎係數"].values[0]
        self.entry_k.delete(0, tk.END)
        self.entry_k.insert(0, str(k_val))

    def calculate(self):
        try:
            t = float(self.combo_thick.get())
            k = float(self.entry_k.get())
            r = float(self.entry_r.get())
            angle = float(self.entry_a.get())
            fa = float(self.entry_fa.get())
            fb = float(self.entry_fb.get())

            rad = math.radians(angle)
            ba = rad * (r + (k * t))
            total_l = (fa - r - t) + (fb - r - t) + ba
            
            self.label_result.config(text=f"折彎補償 (BA): {ba:.3f}\n總展開長度 (L): {total_l:.3f}")
        except:
            messagebox.showerror("錯誤", "請確認材質與厚度已選擇，且各欄位輸入正確。")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金展開計算器 (自定義公式版)")
        self.root.geometry("450x650")
        
        # 讀取資料數據
        self.load_excel_data()
        self.create_widgets()

    def load_excel_data(self):
        # 檔案需與 .py 檔放在同一個資料夾
        file_path = "bend_parameters.xlsx"
        if not os.path.exists(file_path):
            messagebox.showwarning("找不到檔案", f"未找到 {file_path}，請確認檔案位置。")
            self.df = pd.DataFrame(columns=["材質 (Material)", "厚度 (Thickness)", "折彎係數 (K-Factor)"])
        else:
            try:
                self.df = pd.read_excel(file_path)
            except Exception as e:
                messagebox.showerror("錯誤", f"無法讀取 Excel: {e}")

    def create_widgets(self):
        # 標題
        tk.Label(self.root, text="板金展開計算 (連線 Excel)", font=("Microsoft JhengHei", 14, "bold")).pack(pady=15)

        # 選擇材質
        tk.Label(self.root, text="1. 選擇材質:").pack()
        materials = self.df["材質 (Material)"].unique().tolist() if not self.df.empty else []
        self.combo_mat = ttk.Combobox(self.root, values=materials, state="readonly")
        self.combo_mat.pack(pady=5)
        self.combo_mat.bind("<<ComboboxSelected>>", self.update_thickness)

        # 選擇厚度
        tk.Label(self.root, text="2. 選擇厚度 (T):").pack()
        self.combo_thick = ttk.Combobox(self.root, state="readonly")
        self.combo_thick.pack(pady=5)
        self.combo_thick.bind("<<ComboboxSelected>>", self.update_k_val)

        # 自動顯示折彎係數
        self.entry_k = self.create_input("折彎係數 (K):", "0.0")
        
        # 輸入邊長
        self.entry_fa = self.create_input("外部邊長 (A):", "0.0")
        self.entry_fb = self.create_input("外部邊長 (B):", "0.0")

        # 計算按鈕
        btn_calc = tk.Button(self.root, text="執行計算 (A+B-2T+K)", command=self.calculate, 
                             bg="#28a745", fg="white", font=("Arial", 11, "bold"))
        btn_calc.pack(pady=20, ipadx=40)

        # 結果區
        self.res_frame = tk.LabelFrame(self.root, text="計算結果", padx=10, pady=10)
        self.res_frame.pack(padx=30, fill="x")
        self.label_res = tk.Label(self.res_frame, text="等待計算...", font=("Consolas", 12), fg="blue")
        self.label_res.pack()

    def create_input(self, label_text, default_val):
        frame = tk.Frame(self.root)
        frame.pack(fill="x", padx=50, pady=5)
        tk.Label(frame, text=label_text, width=15, anchor="w").pack(side="left")
        entry = tk.Entry(frame)
        entry.insert(0, default_val)
        entry.pack(side="right", expand=True, fill="x")
        return entry

    def update_thickness(self, event):
        sel_mat = self.combo_mat.get()
        thick_list = self.df[self.df["材質 (Material)"] == sel_mat]["厚度 (Thickness)"].tolist()
        self.combo_thick.config(values=thick_list)
        self.combo_thick.set("")
        self.entry_k.delete(0, tk.END)

    def update_k_val(self, event):
        sel_mat = self.combo_mat.get()
        sel_t = float(self.combo_thick.get())
        # 從 DataFrame 抓取對應係數
        k_val = self.df[(self.df["材質 (Material)"] == sel_mat) & (self.df["厚度 (Thickness)"] == sel_t)]["折彎係數 (K-Factor)"].values[0]
        self.entry_k.delete(0, tk.END)
        self.entry_k.insert(0, str(k_val))

    def calculate(self):
        try:
            # 取得數值
            a = float(self.entry_fa.get())
            b = float(self.entry_fb.get())
            t = float(self.combo_thick.get())
            k = float(self.entry_k.get())

            # 執行你的公式: (A+B) - 2T + K
            ans = (a + b) - (2 * t) + k
            
            self.label_res.config(text=f"總展開長度: {ans:.3f} mm")
        except Exception:
            messagebox.showerror("錯誤", "請檢查數值是否正確填寫")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
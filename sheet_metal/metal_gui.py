import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金多角度展開計算器 v4.0")
        self.root.geometry("550x850")
        
        self.excel_file = "bend_parameters.xlsx"
        self.angle_entries = [] # 儲存動態生成的角度輸入框
        self.load_excel_data()
        self.create_widgets()

    def load_excel_data(self):
        if not os.path.exists(self.excel_file):
            messagebox.showwarning("警告", f"找不到 {self.excel_file}")
            self.df = pd.DataFrame()
        else:
            try:
                self.df = pd.read_excel(self.excel_file, index_col=0)
                self.df.columns = [str(col) for col in self.df.columns]
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取失敗: {e}")

    def create_widgets(self):
        tk.Label(self.root, text="板金多折彎-各別角度計算", font=("Microsoft JhengHei", 14, "bold")).pack(pady=10)

        # --- 第一區：Excel 參數 ---
        group_excel = tk.LabelFrame(self.root, text="1. 材質與厚度 (Excel 檢索)", padx=10, pady=10)
        group_excel.pack(padx=20, fill="x")

        tk.Label(group_excel, text="材質:").grid(row=0, column=0, sticky="w")
        self.combo_mat = ttk.Combobox(group_excel, values=self.df.index.tolist(), state="readonly")
        self.combo_mat.grid(row=0, column=1, sticky="ew", padx=5)
        self.combo_mat.bind("<<ComboboxSelected>>", self.on_selection_change)

        tk.Label(group_excel, text="厚度 (T):").grid(row=1, column=0, sticky="w")
        self.combo_thick = ttk.Combobox(group_excel, values=self.df.columns.tolist(), state="readonly")
        self.combo_thick.grid(row=1, column=1, sticky="ew", padx=5)
        self.combo_thick.bind("<<ComboboxSelected>>", self.on_selection_change)

        tk.Label(group_excel, text="90°標準K值:").grid(row=2, column=0, sticky="w")
        self.entry_k90 = tk.Entry(group_excel, bg="#f0f0f0")
        self.entry_k90.grid(row=2, column=1, sticky="ew", padx=5)

        # --- 第二區：折彎數量設定 ---
        group_config = tk.Frame(self.root, padx=20, pady=10)
        group_config.pack(fill="x")

        tk.Label(group_config, text="外部邊長總和 (ΣA):").grid(row=0, column=0, sticky="w")
        self.entry_sum = tk.Entry(group_config)
        self.entry_sum.grid(row=0, column=1, sticky="ew", padx=5)

        tk.Label(group_config, text="折彎次數 (n):").grid(row=1, column=0, sticky="w")
        self.spin_n = tk.Spinbox(group_config, from_=1, to=10, command=self.update_angle_inputs)
        self.spin_n.grid(row=1, column=1, sticky="w", padx=5)
        # 綁定鍵盤輸入事件，手動輸入數字也會更新
        self.spin_n.bind("<KeyRelease>", lambda e: self.update_angle_inputs())

        # --- 第三區：動態角度輸入區 ---
        self.angle_frame = tk.LabelFrame(self.root, text="2. 設定各別折彎角度", padx=10, pady=10)
        self.angle_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        self.update_angle_inputs() # 初始化顯示一個輸入框

        # --- 計算與結果 ---
        self.btn_calc = tk.Button(self.root, text="開始計算展開長度", command=self.calculate, 
                                  bg="#0056b3", fg="white", font=("Arial", 11, "bold"))
        self.btn_calc.pack(pady=10, ipadx=50)

        self.res_label = tk.Label(self.root, text="結果: --", font=("Consolas", 12, "bold"), fg="#d9534f")
        self.res_label.pack(pady=10)

    def on_selection_change(self, event):
        mat = self.combo_mat.get()
        thick = self.combo_thick.get()
        if mat and thick:
            try:
                k90 = self.df.loc[mat, thick]
                self.entry_k90.delete(0, tk.END)
                self.entry_k90.insert(0, str(k90))
            except: pass

    def update_angle_inputs(self):
        """根據折彎次數動態增減角度輸入框"""
        # 清除舊的輸入框
        for widget in self.angle_frame.winfo_children():
            widget.destroy()
        self.angle_entries = []

        try:
            n = int(self.spin_n.get())
        except: n = 1

        for i in range(n):
            row_f = tk.Frame(self.angle_frame)
            row_f.pack(fill="x", pady=2)
            tk.Label(row_f, text=f"第 {i+1} 個角度 (°):", width=15).pack(side="left")
            ent = tk.Entry(row_f)
            ent.insert(0, "90") # 預設 90 度
            ent.pack(side="right", expand=True, fill="x")
            self.angle_entries.append(ent)

    def calculate(self):
        try:
            sigma_a = float(self.entry_sum.get())
            t = float(self.combo_thick.get())
            k90 = float(self.entry_k90.get())
            
            total_k_adj = 0
            n = len(self.angle_entries)
            
            # 遍歷每個角度輸入框計算補償
            for ent in self.angle_entries:
                angle = float(ent.get())
                # 公式: (K90 / 90) * (180 - Angle)
                total_k_adj += (k90 / 90) * (180 - angle)

            # 總展開長度 = ΣA - (n * 2 * T) + ΣK_adj
            result = sigma_a - (n * 2 * t) + total_k_adj
            
            self.res_label.config(text=f"總展開長度: {result:.3f} mm\n(總扣除: {n*2*t:.2f}, 總補償: {total_k_adj:.3f})")
        except Exception:
            messagebox.showerror("錯誤", "請檢查所有數值是否輸入正確")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
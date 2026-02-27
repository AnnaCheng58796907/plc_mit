import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金係數對照計算器 v3.0")
        self.root.geometry("500x750")
        
        self.excel_file = "bend_parameters.xlsx"
        self.load_excel_data()
        self.create_widgets()

    def load_excel_data(self):
        if not os.path.exists(self.excel_file):
            messagebox.showwarning("警告", f"找不到 {self.excel_file}\n請建立格式為「縱向材質、橫向厚度」的 Excel。")
            self.df = pd.DataFrame()
        else:
            try:
                # 讀取 Excel，將第一列(材質)設為索引
                self.df = pd.read_excel(self.excel_file, index_col=0)
                # 確保欄位名稱(厚度)為字串，方便下拉選單使用
                self.df.columns = [str(col) for col in self.df.columns]
            except Exception as e:
                messagebox.showerror("錯誤", f"Excel 讀取失敗: {e}")

    def create_widgets(self):
        tk.Label(self.root, text="板金展開計算 (Excel 矩陣對照)", font=("Microsoft JhengHei", 14, "bold")).pack(pady=15)

        # --- 參數選擇區 ---
        group_param = tk.LabelFrame(self.root, text="自動係數查詢", padx=15, pady=15)
        group_param.pack(padx=20, fill="x")

        # 材質選擇
        tk.Label(group_param, text="1. 選擇材質 (縱向):").pack(anchor="w")
        materials = self.df.index.tolist() if not self.df.empty else []
        self.combo_mat = ttk.Combobox(group_param, values=materials, state="readonly")
        self.combo_mat.pack(fill="x", pady=5)
        self.combo_mat.bind("<<ComboboxSelected>>", self.on_selection_change)

        # 厚度選擇
        tk.Label(group_param, text="2. 選擇厚度 (橫向):").pack(anchor="w")
        thicknesses = self.df.columns.tolist() if not self.df.empty else []
        self.combo_thick = ttk.Combobox(group_param, values=thicknesses, state="readonly")
        self.combo_thick.pack(fill="x", pady=5)
        self.combo_thick.bind("<<ComboboxSelected>>", self.on_selection_change)

        # 顯示查找到的 K 值
        tk.Label(group_param, text="對應折彎係數 (K):").pack(anchor="w")
        self.entry_k = tk.Entry(group_param, font=("Arial", 11, "bold"), fg="red", bg="#eee")
        self.entry_k.pack(fill="x", pady=5)

        # --- 尺寸計算區 ---
        group_calc = tk.LabelFrame(self.root, text="尺寸計算", padx=15, pady=15)
        group_calc.pack(padx=20, pady=15, fill="x")

        tk.Label(group_calc, text="外部邊長總和 (ΣA):").pack(anchor="w")
        self.entry_sum = tk.Entry(group_calc)
        self.entry_sum.pack(fill="x", pady=5)

        tk.Label(group_calc, text="折彎次數 (n):").pack(anchor="w")
        self.spin_n = tk.Spinbox(group_calc, from_=1, to=20)
        self.spin_n.pack(fill="x", pady=5)

        # 按鈕
        self.btn_calc = tk.Button(self.root, text="執行公式計算", command=self.calculate, 
                                  bg="#28a745", fg="white", font=("Microsoft JhengHei", 12, "bold"))
        self.btn_calc.pack(pady=10, ipadx=30)

        # 結果
        self.res_label = tk.Label(self.root, text="結果: -- mm", font=("Consolas", 14, "bold"), fg="#0056b3")
        self.res_label.pack(pady=20)

    def on_selection_change(self, event):
        """當材質或厚度改變時，自動去 Excel 矩陣找數值"""
        mat = self.combo_mat.get()
        thick = self.combo_thick.get()
        
        if mat and thick:
            try:
                # 根據橫縱座標查找
                k_val = self.df.loc[mat, thick]
                self.entry_k.delete(0, tk.END)
                self.entry_k.insert(0, str(k_val))
            except Exception:
                self.entry_k.delete(0, tk.END)
                self.entry_k.insert(0, "無資料")

    def calculate(self):
        try:
            # 讀取數值
            sigma_a = float(self.entry_sum.get())
            t = float(self.combo_thick.get())
            k = float(self.entry_k.get())
            n = int(self.spin_n.get())

            # 您的公式: (ΣA) - (n * 2T) + (n * K)
            result = sigma_a - (n * 2 * t) + (n * k)
            self.res_label.config(text=f"總展開長度: {result:.3f} mm")
        except:
            messagebox.showerror("錯誤", "請檢查輸入數值與 Excel 選項是否正確")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
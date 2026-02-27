import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金全功能計算器 v3.5 (角度換算版)")
        self.root.geometry("500x800")
        
        self.excel_file = "bend_parameters.xlsx"
        self.load_excel_data()
        self.create_widgets()

    def load_excel_data(self):
        if not os.path.exists(self.excel_file):
            messagebox.showwarning("警告", f"找不到 {self.excel_file}")
            self.df = pd.DataFrame()
        else:
            try:
                # 讀取矩陣式 Excel
                self.df = pd.read_excel(self.excel_file, index_col=0)
                self.df.columns = [str(col) for col in self.df.columns]
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取失敗: {e}")

    def create_widgets(self):
        tk.Label(self.root, text="板金展開計算器 (動態角度)", font=("Microsoft JhengHei", 14, "bold")).pack(pady=15)

        # --- 第一區：自動搜尋 (與之前相同) ---
        group_search = tk.LabelFrame(self.root, text="Step 1: 參數檢索", padx=15, pady=10)
        group_search.pack(padx=20, fill="x")

        tk.Label(group_search, text="選擇材質:").pack(anchor="w")
        self.combo_mat = ttk.Combobox(group_search, values=self.df.index.tolist(), state="readonly")
        self.combo_mat.pack(fill="x", pady=2)
        self.combo_mat.bind("<<ComboboxSelected>>", self.on_selection_change)

        tk.Label(group_search, text="選擇厚度 (T):").pack(anchor="w")
        self.combo_thick = ttk.Combobox(group_search, values=self.df.columns.tolist(), state="readonly")
        self.combo_thick.pack(fill="x", pady=2)
        self.combo_thick.bind("<<ComboboxSelected>>", self.on_selection_change)

        tk.Label(group_search, text="90度標準係數 (K90):").pack(anchor="w")
        self.entry_k90 = tk.Entry(group_search, bg="#f0f0f0")
        self.entry_k90.pack(fill="x", pady=2)

        # --- 第二區：角度與尺寸 ---
        group_calc = tk.LabelFrame(self.root, text="Step 2: 尺寸與角度輸入", padx=15, pady=10)
        group_calc.pack(padx=20, pady=10, fill="x")

        tk.Label(group_calc, text="折彎角度 (度):").pack(anchor="w")
        self.entry_angle = tk.Entry(group_calc)
        self.entry_angle.insert(0, "90") # 預設 90
        self.entry_angle.pack(fill="x", pady=2)

        tk.Label(group_calc, text="外部邊長總和 (ΣA):").pack(anchor="w")
        self.entry_sum = tk.Entry(group_calc)
        self.entry_sum.pack(fill="x", pady=2)

        tk.Label(group_calc, text="折彎次數 (n):").pack(anchor="w")
        self.spin_n = tk.Spinbox(group_calc, from_=1, to=20)
        self.spin_n.pack(fill="x", pady=2)

        # --- 計算按鈕 ---
        self.btn_calc = tk.Button(self.root, text="執行公式計算", command=self.calculate, 
                                  bg="#28a745", fg="white", font=("Microsoft JhengHei", 12, "bold"))
        self.btn_calc.pack(pady=15, ipadx=40)

        # --- 結果顯示 ---
        self.res_frame = tk.LabelFrame(self.root, text="計算報告", padx=10, pady=10)
        self.res_frame.pack(padx=20, fill="both", expand=True)
        self.label_res = tk.Label(self.res_frame, text="等待計算...", font=("Consolas", 10), justify="left")
        self.label_res.pack(anchor="w")

    def on_selection_change(self, event):
        mat = self.combo_mat.get()
        thick = self.combo_thick.get()
        if mat and thick:
            try:
                k90 = self.df.loc[mat, thick]
                self.entry_k90.delete(0, tk.END)
                self.entry_k90.insert(0, str(k90))
            except:
                pass

    def calculate(self):
        try:
            # 取得基礎值
            t = float(self.combo_thick.get())
            k90 = float(self.entry_k90.get())
            angle = float(self.entry_angle.get())
            sigma_a = float(self.entry_sum.get())
            n = int(self.spin_n.get())

            # 1. 計算該角度下的調整係數 K_adj
            # 公式: (K90 / 90) * (180 - Angle)
            k_adj = (k90 / 90) * (180 - angle)

            # 2. 計算總展開長度
            # 公式: ΣA - (n * 2T) + (n * K_adj)
            result = sigma_a - (n * 2 * t) + (n * k_adj)

            report = (
                f"角度調整係數 (K'): {k_adj:.3f}\n"
                f"總扣除值 (n*2T): -{n * 2 * t:.2f} mm\n"
                f"總補償值 (n*K'): +{n * k_adj:.2f} mm\n"
                f"--------------------------------\n"
                f"最終展開長度: {result:.3f} mm"
            )
            self.label_res.config(text=report)

        except Exception as e:
            messagebox.showerror("計算錯誤", "請確認所有欄位已填寫正確數值")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
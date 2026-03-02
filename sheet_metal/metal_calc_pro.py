import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金內邊展開專業計算器 v6.0")
        self.root.geometry("600x850")
        
        # 檔案路徑與資料儲存
        self.excel_file = "bend_parameters.xlsx"
        self.side_entries = []   # 儲存內邊長度輸入框
        self.angle_entries = []  # 儲存折彎角度輸入框
        
        self.load_excel_data()
        self.create_widgets()

    def load_excel_data(self):
        """從 Excel 讀取材質/厚度矩陣"""
        if os.path.exists(self.excel_file):
            try:
                # 讀取第一個工作表，將第一列設為索引(材質)
                self.df = pd.read_excel(self.excel_file, sheet_name=0, index_col=0)
                # 強制轉換標題(厚度)為字串，避免查詢錯誤
                self.df.columns = [str(col) for col in self.df.columns]
            except Exception as e:
                print(f"Excel 讀取錯誤: {e}")
                self.df = pd.DataFrame()
        else:
            self.df = pd.DataFrame()

    def create_widgets(self):
        # --- 標題區 ---
        header = tk.Label(self.root, text="板金展開計算 (內邊相加法)", font=("Microsoft JhengHei", 16, "bold"), fg="#1A237E")
        header.pack(pady=15)

        # --- 第一區：自動搜尋 (從 Excel 抓取 K90) ---
        group_k = tk.LabelFrame(self.root, text="1. 參數檢索 (Excel 自動連動)", padx=15, pady=10)
        group_k.pack(padx=20, fill="x")

        # 材質與厚度選擇
        f_select = tk.Frame(group_k)
        f_select.pack(fill="x")
        
        tk.Label(f_select, text="材質:").grid(row=0, column=0, sticky="w")
        self.c_mat = ttk.Combobox(f_select, values=self.df.index.tolist(), state="readonly")
        self.c_mat.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.c_mat.bind("<<ComboboxSelected>>", self.update_k90)

        tk.Label(f_select, text="厚度 (T):").grid(row=1, column=0, sticky="w")
        self.c_thick = ttk.Combobox(f_select, values=self.df.columns.tolist(), state="readonly")
        self.c_thick.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.c_thick.bind("<<ComboboxSelected>>", self.update_k90)

        # 顯示 K90 數值
        tk.Label(group_k, text="90° 標準補償係數 (K90):").pack(anchor="w", pady=(5,0))
        self.e_k90 = tk.Entry(group_k, font=("Arial", 11, "bold"), fg="blue", bg="#F5F5F5")
        self.e_k90.pack(fill="x", pady=2)

        # --- 第二區：結構設定 ---
        group_cfg = tk.Frame(self.root, padx=20, pady=10)
        group_cfg.pack(fill="x")
        
        tk.Label(group_cfg, text="折彎次數 (n):", font=("Microsoft JhengHei", 10)).pack(side="left")
        self.s_n = tk.Spinbox(group_cfg, from_=1, to=15, width=8, command=self.refresh_inputs)
        self.s_n.pack(side="left", padx=10)
        # 鍵盤輸入後也要更新介面
        self.s_n.bind("<KeyRelease>", lambda e: self.refresh_inputs())

        # --- 第三區：動態輸入區 (內邊與角度) ---
        # 使用 Canvas 加上 Scrollbar 以防輸入框太多
        self.input_container = tk.LabelFrame(self.root, text="2. 輸入內邊長度與折彎角度", padx=15, pady=10)
        self.input_container.pack(padx=20, pady=5, fill="both", expand=True)
        
        self.refresh_inputs()

        # --- 第四區：計算與結果 ---
        calc_btn = tk.Button(self.root, text="執行公式計算 (Σ內邊 + ΣK')", command=self.calculate, 
                             bg="#2E7D32", fg="white", font=("Microsoft JhengHei", 12, "bold"))
        calc_btn.pack(pady=15, ipadx=60, ipady=5)

        self.res_frame = tk.LabelFrame(self.root, text="計算報告", padx=15, pady=15)
        self.res_frame.pack(padx=20, pady=10, fill="x")
        
        self.l_res = tk.Label(self.res_frame, text="請輸入數值後點擊計算", font=("Consolas", 11), justify="left")
        self.l_res.pack(anchor="w")

    def update_k90(self, event):
        """根據材質厚度查找 Excel 值"""
        m, t = self.c_mat.get(), self.c_thick.get()
        if m and t:
            try:
                val = self.df.loc[m, t]
                self.e_k90.delete(0, tk.END)
                self.e_k90.insert(0, str(val))
            except: pass

    def refresh_inputs(self):
        """動態生成 L1, Angle1, L2... 的 UI"""
        # 清除舊的
        for w in self.input_container.winfo_children(): w.destroy()
        self.side_entries = []
        self.angle_entries = []

        try:
            n = int(self.s_n.get())
        except: n = 1

        # 內邊長度 L (n+1 個), 角度 Angle (n 個)
        for i in range(n + 1):
            # 內邊長度輸入
            f_s = tk.Frame(self.input_container)
            f_s.pack(fill="x", pady=2)
            tk.Label(f_s, text=f"內邊長度 L{i+1} (mm):", width=18, anchor="w").pack(side="left")
            ent_s = tk.Entry(f_s)
            ent_s.pack(side="right", expand=True, fill="x")
            self.side_entries.append(ent_s)

            # 如果不是最後一段，加入折彎角度輸入
            if i < n:
                f_a = tk.Frame(self.input_container)
                f_a.pack(fill="x", pady=2)
                tk.Label(f_a, text=f" ↳ 折彎角度 {i+1} (°):", width=18, anchor="w", fg="#757575").pack(side="left")
                ent_a = tk.Entry(f_a)
                ent_a.insert(0, "90")
                ent_a.pack(side="right", expand=True, fill="x")
                self.angle_entries.append(ent_a)

    def calculate(self):
        """核心計算邏輯"""
        try:
            k90 = float(self.e_k90.get())
            
            # 1. 累加所有內邊長度
            sum_sides = sum([float(e.get() or 0) for e in self.side_entries])
            
            # 2. 計算各角度補償
            sum_k_adj = 0
            detail_info = ""
            
            for i, e in enumerate(self.angle_entries):
                angle = float(e.get() or 90)
                # 角度補償公式: (K90 / 90) * (180 - Angle)
                k_adj = (k90 / 90) * (180 - angle)
                sum_k_adj += k_adj
                
                # 判斷角度類型
                shape = "鈍角(平)" if angle < 90 else ("直角" if angle == 90 else "銳角(尖)")
                detail_info += f"  - 折彎{i+1}: {angle}° ({shape}) → K'={k_adj:.3f}\n"

            total_length = sum_sides + sum_k_adj
            
            # 輸出報告
            report = (
                f"【分析結果】\n"
                f"{detail_info}"
                f"--------------------------------\n"
                f"內邊總和 (ΣL): {sum_sides:.2f} mm\n"
                f"總補償值 (ΣK'): {sum_k_adj:.3f} mm\n"
                f"================================\n"
                f"展開總長度: {total_length:.3f} mm"
            )
            self.l_res.config(text=report, fg="black")
            
        except ValueError:
            messagebox.showerror("輸入錯誤", "請檢查是否所有欄位都已填入數字")
        except Exception as e:
            messagebox.showerror("錯誤", f"計算發生問題: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
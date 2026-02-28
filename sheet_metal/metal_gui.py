import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金設計綜合工具 v5.5")
        self.root.geometry("600x850")
        
        self.excel_file = "bend_parameters.xlsx"
        self.angle_entries = [] # 儲存展開分頁的角度輸入框
        
        self.load_all_data()
        
        # 建立分頁控鍵
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")
        
        # 建立兩個主要分頁
        self.tab_bend = tk.Frame(self.notebook)
        self.tab_hw = tk.Frame(self.notebook)
        
        self.notebook.add(self.tab_bend, text=" 📐 板金展開計算 ")
        self.notebook.add(self.tab_hw, text=" 🔩 鉚合零件查詢 ")
        
        self.setup_bend_tab()      # 設置分頁 1 (展開)
        self.setup_hardware_tab()  # 設置分頁 2 (硬體)

    def load_all_data(self):
        """讀取 Excel 中的兩個矩陣工作表"""
        if not os.path.exists(self.excel_file):
            messagebox.showwarning("警告", f"找不到 {self.excel_file}\n請確認 Excel 包含兩個分頁。")
            self.df_bend = pd.DataFrame()
            self.df_hw = pd.DataFrame()
        else:
            try:
                # 讀取工作表1: 折彎參數 (假設在第1頁)
                self.df_bend = pd.read_excel(self.excel_file, sheet_name=0, index_col=0)
                self.df_bend.columns = [str(col) for col in self.df_bend.columns]
                
                # 讀取工作表2: Hardware (矩陣格式)
                self.df_hw = pd.read_excel(self.excel_file, sheet_name="Hardware", index_col=0)
                self.df_hw.columns = [str(col) for col in self.df_hw.columns]
            except Exception as e:
                messagebox.showerror("錯誤", f"Excel 讀取失敗: {e}")

    # --- 分頁 1: 多角度展開計算 ---
    def setup_bend_tab(self):
        container = tk.Frame(self.tab_bend, padx=20, pady=10)
        container.pack(fill="both", expand=True)

        tk.Label(container, text="板金參數與折彎設定", font=("Arial", 12, "bold")).pack(pady=5)
        
        # 選擇材質/厚度
        f_top = tk.Frame(container)
        f_top.pack(fill="x")
        
        tk.Label(f_top, text="材質:").grid(row=0, column=0, sticky="w")
        self.c_mat = ttk.Combobox(f_top, values=self.df_bend.index.tolist(), state="readonly")
        self.c_mat.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        self.c_mat.bind("<<ComboboxSelected>>", self.update_bend_k)

        tk.Label(f_top, text="厚度:").grid(row=1, column=0, sticky="w")
        self.c_thick = ttk.Combobox(f_top, values=self.df_bend.columns.tolist(), state="readonly")
        self.c_thick.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        self.c_thick.bind("<<ComboboxSelected>>", self.update_bend_k)

        tk.Label(f_top, text="90° K值:").grid(row=2, column=0, sticky="w")
        self.e_k90 = tk.Entry(f_top, bg="#eee")
        self.e_k90.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        # 尺寸與次數
        tk.Label(container, text="外部邊長總和:").pack(anchor="w", pady=(10,0))
        self.e_sum_a = tk.Entry(container)
        self.e_sum_a.pack(fill="x", pady=2)

        tk.Label(container, text="折彎次數 (n):").pack(anchor="w")
        self.s_n = tk.Spinbox(container, from_=1, to=10, command=self.refresh_angles)
        self.s_n.pack(fill="x", pady=2)
        self.s_n.bind("<KeyRelease>", lambda e: self.refresh_angles())

        self.angle_area = tk.LabelFrame(container, text="各折彎角度 (°)", padx=10, pady=10)
        self.angle_area.pack(fill="both", expand=True, pady=10)
        self.refresh_angles()

        tk.Button(container, text="計算展開長度", bg="#28a745", fg="white", font=("Arial", 11, "bold"),
                  command=self.calc_bend).pack(fill="x", pady=10)
        self.l_bend_res = tk.Label(container, text="結果: --", font=("Arial", 12, "bold"), fg="blue")
        self.l_bend_res.pack()

    # --- 分頁 2: 鉚合開孔查詢 (矩陣對照) ---
    def setup_hardware_tab(self):
        container = tk.Frame(self.tab_hw, padx=30, pady=20)
        container.pack(fill="both", expand=True)

        tk.Label(container, text="零件開孔規格檢索", font=("Arial", 14, "bold")).pack(pady=10)

        # 種類選擇 (縱向)
        tk.Label(container, text="1. 選擇零件種類 (如螺帽/螺柱):").pack(anchor="w")
        hw_types = self.df_hw.index.tolist() if not self.df_hw.empty else []
        self.c_hw_type = ttk.Combobox(container, values=hw_types, state="readonly")
        self.c_hw_type.pack(fill="x", pady=5)
        self.c_hw_type.bind("<<ComboboxSelected>>", self.lookup_hardware)

        # 規格選擇 (橫向)
        tk.Label(container, text="2. 選擇規格尺寸 (如 M3, 1/8\"):").pack(anchor="w")
        hw_specs = self.df_hw.columns.tolist() if not self.df_hw.empty else []
        self.c_hw_spec = ttk.Combobox(container, values=hw_specs, state="readonly")
        self.c_hw_spec.pack(fill="x", pady=5)
        self.c_hw_spec.bind("<<ComboboxSelected>>", self.lookup_hardware)

        # 顯示結果
        self.hw_res_frame = tk.LabelFrame(container, text="查詢結果", padx=20, pady=20)
        self.hw_res_frame.pack(fill="x", pady=30)
        
        self.l_hw_hole = tk.Label(self.hw_res_frame, text="建議開孔: --", font=("Arial", 20, "bold"), fg="#d9534f")
        self.l_hw_hole.pack()

    # --- 邏輯處理 ---
    def update_bend_k(self, event):
        m, t = self.c_mat.get(), self.c_thick.get()
        if m and t:
            self.e_k90.delete(0, tk.END)
            self.e_k90.insert(0, str(self.df_bend.loc[m, t]))

    def refresh_angles(self):
        for w in self.angle_area.winfo_children(): w.destroy()
        self.angle_entries = []
        try: n = int(self.s_n.get())
        except: n = 1
        for i in range(n):
            f = tk.Frame(self.angle_area); f.pack(fill="x", pady=1)
            tk.Label(f, text=f"折彎 {i+1} 角度:").pack(side="left")
            e = tk.Entry(f); e.insert(0, "90"); e.pack(side="right", expand=True, fill="x")
            self.angle_entries.append(e)

    def calc_bend(self):
        try:
            sum_a = float(self.e_sum_a.get())
            t = float(self.c_thick.get())
            k90 = float(self.e_k90.get())
            total_k = sum([(k90/90)*(180-float(e.get())) for e in self.angle_entries])
            n = len(self.angle_entries)
            res = sum_a - (n*2*t) + total_k
            self.l_bend_res.config(text=f"總展開長度: {res:.3f} mm")
        except: messagebox.showerror("錯誤", "請檢查輸入數值")

    def lookup_hardware(self, event):
        t, s = self.c_hw_type.get(), self.c_hw_spec.get()
        if t and s:
            val = self.df_hw.loc[t, s]
            self.l_hw_hole.config(text=f"Ø {val} mm" if str(val) != "nan" else "無對應資料")

if __name__ == "__main__":
    root = tk.Tk(); app = MetalApp(root); root.mainloop()
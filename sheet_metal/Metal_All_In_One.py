import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金設計全功能工具箱 v7.0")
        self.root.geometry("600x900")
        
        self.excel_file = "bend_parameters.xlsx"
        self.side_entries = []
        self.angle_entries = []
        
        self.load_all_data()
        
        # 建立分頁控鍵
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")
        
        self.tab_bend = tk.Frame(self.notebook)
        self.tab_hw = tk.Frame(self.notebook)
        
        self.notebook.add(self.tab_bend, text=" 📐 展開計算 (內邊法) ")
        self.notebook.add(self.tab_hw, text=" 🔩 鉚合零件開孔 ")
        
        self.setup_bend_tab()
        self.setup_hw_tab()

    def load_all_data(self):
        """讀取 Excel 中的兩個工作表"""
        if os.path.exists(self.excel_file):
            try:
                # 讀取折彎參數
                self.df_bend = pd.read_excel(self.excel_file, sheet_name=0, index_col=0)
                self.df_bend.columns = [str(col) for col in self.df_bend.columns]
                
                # 讀取鉚合參數 (Hardware 分頁)
                self.df_hw = pd.read_excel(self.excel_file, sheet_name="Hardware", index_col=0)
                self.df_hw.columns = [str(col) for col in self.df_hw.columns]
            except Exception as e:
                print(f"Excel 讀取失敗: {e}")
                self.df_bend = pd.DataFrame()
                self.df_hw = pd.DataFrame()
        else:
            self.df_bend = pd.DataFrame()
            self.df_hw = pd.DataFrame()

    # --- 分頁 1: 展開計算邏輯 ---
    def setup_bend_tab(self):
        container = tk.Frame(self.tab_bend, padx=20, pady=10)
        container.pack(fill="both", expand=True)

        tk.Label(container, text="內邊長度累加 + 角度補償", font=("Arial", 14, "bold")).pack(pady=10)

        # K90 檢索區
        f_k = tk.LabelFrame(container, text="1. 取得 K90 係數", padx=10, pady=10)
        f_k.pack(fill="x", pady=5)
        
        tk.Label(f_k, text="材質:").grid(row=0, column=0, sticky="w")
        self.c_mat = ttk.Combobox(f_k, values=self.df_bend.index.tolist(), state="readonly")
        self.c_mat.grid(row=0, column=1, sticky="ew", padx=5)
        self.c_mat.bind("<<ComboboxSelected>>", self.update_k90)

        tk.Label(f_k, text="厚度:").grid(row=1, column=0, sticky="w")
        self.c_thick = ttk.Combobox(f_k, values=self.df_bend.columns.tolist(), state="readonly")
        self.c_thick.grid(row=1, column=1, sticky="ew", padx=5)
        self.c_thick.bind("<<ComboboxSelected>>", self.update_k90)

        self.e_k90 = tk.Entry(f_k, font=("Arial", 11, "bold"), fg="blue")
        self.e_k90.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        tk.Label(f_k, text="90° K值:").grid(row=2, column=0, sticky="w")

        # 折彎輸入區
        tk.Label(container, text="折彎次數 (n):").pack(anchor="w", pady=(10,0))
        self.s_n = tk.Spinbox(container, from_=1, to=15, command=self.refresh_bend_ui)
        self.s_n.pack(fill="x")
        self.s_n.bind("<KeyRelease>", lambda e: self.refresh_bend_ui())

        self.bend_area = tk.LabelFrame(container, text="2. 輸入尺寸與角度", padx=10, pady=10)
        self.bend_area.pack(fill="both", expand=True, pady=10)
        
        self.refresh_bend_ui()

        tk.Button(container, text="執行計算", bg="#28a745", fg="white", font=("Arial", 12, "bold"),
                  command=self.calc_bend).pack(fill="x", pady=10)
        self.l_bend_res = tk.Label(container, text="結果: --", font=("Arial", 12, "bold"), fg="red")
        self.l_bend_res.pack()

    def update_k90(self, event):
        m, t = self.c_mat.get(), self.c_thick.get()
        if m and t:
            try:
                self.e_k90.delete(0, tk.END)
                self.e_k90.insert(0, str(self.df_bend.loc[m, t]))
            except: pass

    def refresh_bend_ui(self):
        for w in self.bend_area.winfo_children(): w.destroy()
        self.side_entries = []
        self.angle_entries = []
        try: n = int(self.s_n.get())
        except: n = 1
        for i in range(n + 1):
            f_s = tk.Frame(self.bend_area); f_s.pack(fill="x", pady=1)
            tk.Label(f_s, text=f"內邊 L{i+1}:", width=12).pack(side="left")
            e_s = tk.Entry(f_s); e_s.pack(side="right", expand=True, fill="x")
            self.side_entries.append(e_s)
            if i < n:
                f_a = tk.Frame(self.bend_area); f_a.pack(fill="x", pady=1)
                tk.Label(f_a, text=f"  ↳ 角度 {i+1}:", width=12, fg="gray").pack(side="left")
                e_a = tk.Entry(f_a); e_a.insert(0, "90"); e_a.pack(side="right", expand=True, fill="x")
                self.angle_entries.append(e_a)

    def calc_bend(self):
        try:
            k90 = float(self.e_k90.get())
            sum_s = sum([float(e.get() or 0) for e in self.side_entries])
            sum_k = sum([(k90/90)*(180-float(e.get() or 90)) for e in self.angle_entries])
            self.l_bend_res.config(text=f"總展開長度: {sum_s + sum_k:.3f} mm")
        except: messagebox.showerror("錯誤", "請檢查數值輸入")

    # --- 分頁 2: 鉚合零件查詢 ---
    def setup_hw_tab(self):
        container = tk.Frame(self.tab_hw, padx=30, pady=30)
        container.pack(fill="both", expand=True)

        tk.Label(container, text="零件開孔尺寸檢索 (矩陣對照)", font=("Arial", 14, "bold")).pack(pady=10)

        f_query = tk.LabelFrame(container, text="查詢條件", padx=20, pady=20)
        f_query.pack(fill="x", pady=10)

        tk.Label(f_query, text="1. 選擇種類 (螺帽/螺柱/拉釘):").pack(anchor="w")
        hw_types = self.df_hw.index.tolist() if not self.df_hw.empty else []
        self.c_hw_type = ttk.Combobox(f_query, values=hw_types, state="readonly")
        self.c_hw_type.pack(fill="x", pady=5)
        self.c_hw_type.bind("<<ComboboxSelected>>", self.lookup_hw)

        tk.Label(f_query, text="2. 選擇規格 (M2 / M3 / 1/8\"):").pack(anchor="w")
        hw_specs = self.df_hw.columns.tolist() if not self.df_hw.empty else []
        self.c_hw_spec = ttk.Combobox(f_query, values=hw_specs, state="readonly")
        self.c_hw_spec.pack(fill="x", pady=5)
        self.c_hw_spec.bind("<<ComboboxSelected>>", self.lookup_hw)

        self.res_hw_frame = tk.LabelFrame(container, text="查詢結果", padx=20, pady=20)
        self.res_hw_frame.pack(fill="x", pady=30)
        self.l_hw_res = tk.Label(self.res_hw_res_frame if hasattr(self, 'res_hw_res_frame') else self.res_hw_frame, 
                                 text="--", font=("Arial", 32, "bold"), fg="#D32F2F")
        self.l_hw_res.pack()

    def lookup_hw(self, event):
        t, s = self.c_hw_type.get(), self.c_hw_spec.get()
        if t and s:
            try:
                val = self.df_hw.loc[t, s]
                self.l_hw_res.config(text=f"Ø {val}" if str(val) != "nan" else "無資料")
            except: self.l_hw_res.config(text="查詢失敗")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金設計全功能工具箱 v8.0")
        self.root.geometry("650x900")
        
        self.excel_file = "bend_parameters.xlsx"
        self.load_all_data()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")
        
        # 建立三個分頁
        self.tab_bend = tk.Frame(self.notebook)
        self.tab_special = tk.Frame(self.notebook)
        self.tab_hw = tk.Frame(self.notebook)
        
        self.notebook.add(self.tab_bend, text=" 📐 動態角度計算 ")
        self.notebook.add(self.tab_special, text=" 🌀 特殊固定折彎 ")
        self.notebook.add(self.tab_hw, text=" 🔩 鉚合零件查詢 ")
        
        self.setup_bend_tab()
        self.setup_special_tab()
        self.setup_hw_tab()

    def load_all_data(self):
        """讀取 Excel 中的三個工作表"""
        if os.path.exists(self.excel_file):
            try:
                self.df_bend = pd.read_excel(self.excel_file, sheet_name=0, index_col=0)
                self.df_hw = pd.read_excel(self.excel_file, sheet_name="Hardware", index_col=0)
                self.df_special = pd.read_excel(self.excel_file, sheet_name="Special", index_col=0)
                
                # 清洗欄位標題為字串
                for df in [self.df_bend, self.df_hw, self.df_special]:
                    df.columns = [str(col) for col in df.columns]
            except Exception as e:
                messagebox.showwarning("Excel錯誤", f"讀取分頁失敗，請檢查分頁名稱(Hardware, Special): {e}")
                self.df_bend = self.df_hw = self.df_special = pd.DataFrame()
        else:
            self.df_bend = self.df_hw = self.df_special = pd.DataFrame()

    # --- 分頁 1: 動態角度計算 (內邊法) ---
    def setup_bend_tab(self):
        container = tk.Frame(self.tab_bend, padx=20, pady=10)
        container.pack(fill="both", expand=True)
        tk.Label(container, text="Σ內邊長 + 角度變動補償", font=("Arial", 12, "bold")).pack(pady=5)
        
        # K90 檢索
        f_k = tk.LabelFrame(container, text="1. 取得 K90 係數")
        f_k.pack(fill="x", pady=5, padx=5)
        self.c_mat = ttk.Combobox(f_k, values=self.df_bend.index.tolist(), state="readonly")
        self.c_mat.pack(side="left", padx=5, pady=5)
        self.c_thick = ttk.Combobox(f_k, values=self.df_bend.columns.tolist(), state="readonly")
        self.c_thick.pack(side="left", padx=5)
        self.e_k90 = tk.Entry(f_k, width=10, fg="blue", font=("Arial", 10, "bold"))
        self.e_k90.pack(side="left", padx=5)
        self.c_mat.bind("<<ComboboxSelected>>", self.update_k90)
        self.c_thick.bind("<<ComboboxSelected>>", self.update_k90)

        # 折彎輸入
        tk.Label(container, text="折彎次數 (n):").pack(anchor="w")
        self.s_n = tk.Spinbox(container, from_=1, to=10, command=self.refresh_bend_ui)
        self.s_n.pack(fill="x")
        self.bend_area = tk.LabelFrame(container, text="2. 輸入尺寸/角度", padx=10, pady=10)
        self.bend_area.pack(fill="both", expand=True, pady=5)
        self.refresh_bend_ui()

        tk.Button(container, text="計算展開", bg="#28a745", fg="white", command=self.calc_bend).pack(fill="x", pady=5)
        self.l_res = tk.Label(container, text="結果: --", font=("Arial", 14, "bold"), fg="red")
        self.l_res.pack()

    def update_k90(self, event):
        m, t = self.c_mat.get(), self.c_thick.get()
        if m and t:
            try:
                self.e_k90.delete(0, tk.END)
                self.e_k90.insert(0, str(self.df_bend.loc[m, t]))
            except: pass

    def refresh_bend_ui(self):
        for w in self.bend_area.winfo_children(): w.destroy()
        self.sides = []; self.angles = []
        try: n = int(self.s_n.get())
        except: n = 1
        for i in range(n + 1):
            tk.Label(self.bend_area, text=f"內邊 L{i+1}:").pack()
            e_s = tk.Entry(self.bend_area); e_s.pack(fill="x")
            self.sides.append(e_s)
            if i < n:
                tk.Label(self.bend_area, text=f"  ↳ 角度 {i+1}:", fg="gray").pack()
                e_a = tk.Entry(self.bend_area); e_a.insert(0, "90"); e_a.pack(fill="x")
                self.angles.append(e_a)

    def calc_bend(self):
        try:
            k90 = float(self.e_k90.get())
            s_sum = sum([float(e.get() or 0) for e in self.sides])
            k_sum = sum([(k90/90)*(180-float(e.get() or 90)) for e in self.angles])
            self.l_res.config(text=f"總長: {s_sum + k_sum:.3f} mm")
        except: messagebox.showerror("錯誤", "數值輸入有誤")

    # --- 分頁 2: 特殊固定折彎 (讀取 Special 分頁) ---
    def setup_special_tab(self):
        container = tk.Frame(self.tab_special, padx=30, pady=20)
        container.pack(fill="both", expand=True)
        tk.Label(container, text="特定工法係數查詢", font=("Arial", 14, "bold")).pack(pady=10)

        f_sp = tk.LabelFrame(container, text="選擇特殊折彎項目", padx=20, pady=20)
        f_sp.pack(fill="x")

        tk.Label(f_sp, text="折彎類型:").pack(anchor="w")
        self.c_sp_type = ttk.Combobox(f_sp, values=self.df_special.index.tolist(), state="readonly")
        self.c_sp_type.pack(fill="x", pady=5)

        tk.Label(f_sp, text="板厚:").pack(anchor="w")
        self.c_sp_thick = ttk.Combobox(f_sp, values=self.df_special.columns.tolist(), state="readonly")
        self.c_sp_thick.pack(fill="x", pady=5)

        self.l_sp_res = tk.Label(container, text="固定補償值: --", font=("Arial", 18, "bold"), fg="#D32F2F")
        self.l_sp_res.pack(pady=30)
        
        self.c_sp_type.bind("<<ComboboxSelected>>", self.lookup_special)
        self.c_sp_thick.bind("<<ComboboxSelected>>", self.lookup_special)

    def lookup_special(self, event):
        t, th = self.c_sp_type.get(), self.c_sp_thick.get()
        if t and th:
            try:
                val = self.df_special.loc[t, th]
                self.l_sp_res.config(text=f"固定補償值: {val} mm")
            except: self.l_sp_res.config(text="查無資料")

    # --- 分頁 3: 鉚合零件 (維持原樣) ---
    def setup_hw_tab(self):
        container = tk.Frame(self.tab_hw, padx=30, pady=20)
        container.pack(fill="both", expand=True)
        tk.Label(container, text="零件開孔規格檢索", font=("Arial", 14, "bold")).pack(pady=10)
        
        self.c_hw_type = ttk.Combobox(container, values=self.df_hw.index.tolist(), state="readonly")
        self.c_hw_type.pack(fill="x", pady=10)
        self.c_hw_spec = ttk.Combobox(container, values=self.df_hw.columns.tolist(), state="readonly")
        self.c_hw_spec.pack(fill="x", pady=10)
        
        self.l_hw_res = tk.Label(container, text="Ø --", font=("Arial", 32, "bold"), fg="#D32F2F")
        self.l_hw_res.pack(pady=30)
        
        self.c_hw_type.bind("<<ComboboxSelected>>", self.lookup_hw)
        self.c_hw_spec.bind("<<ComboboxSelected>>", self.lookup_hw)

    def lookup_hw(self, event):
        t, s = self.c_hw_type.get(), self.c_hw_spec.get()
        if t and s:
            try:
                val = self.df_hw.loc[t, s]
                self.l_hw_res.config(text=f"Ø {val}")
            except: self.l_hw_res.config(text="--")

if __name__ == "__main__":
    root = tk.Tk(); app = MetalApp(root); root.mainloop()
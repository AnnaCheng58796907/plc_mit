import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金設計全功能工具箱 v9.0 - 內邊相加版")
        self.root.geometry("650x900")
        
        self.excel_file = "bend_parameters.xlsx"
        self.all_sheets = {}
        self.side_entries = []
        self.angle_entries = []
        
        self.load_all_excel_sheets()
        
        # 建立分頁系統
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")
        
        self.tab_bend = tk.Frame(self.notebook)
        self.tab_special = tk.Frame(self.notebook)
        self.tab_hw = tk.Frame(self.notebook)
        
        self.notebook.add(self.tab_bend, text=" 📐 動態角度展開 ")
        self.notebook.add(self.tab_special, text=" 🌀 特殊固定折彎 ")
        self.notebook.add(self.tab_hw, text=" 🔩 鉚合零件查詢 ")
        
        self.setup_bend_tab()
        self.setup_special_tab()
        self.setup_hw_tab()

    def load_all_excel_sheets(self):
        """讀取 Excel 內所有分頁"""
        if os.path.exists(self.excel_file):
            try:
                xls = pd.ExcelFile(self.excel_file)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name, index_col=0)
                    df.columns = [str(col) for col in df.columns]
                    self.all_sheets[sheet_name] = df
            except Exception as e:
                messagebox.showerror("Excel 錯誤", f"讀取失敗: {e}")
        else:
            messagebox.showwarning("檔案缺失", f"找不到 {self.excel_file}\n請確認檔案存在於程式同資料夾。")

    # --- 分頁 1: 動態角度展開 (內邊法 + 自選係數表) ---
    def setup_bend_tab(self):
        container = tk.Frame(self.tab_bend, padx=20, pady=10)
        container.pack(fill="both", expand=True)

        # 1. 選擇係數表
        f_sheet = tk.LabelFrame(container, text="1. 選擇係數來源 (工作表)", padx=10, pady=5)
        f_sheet.pack(fill="x", pady=5)
        
        bend_sheets = [s for s in self.all_sheets.keys() if s not in ["Hardware", "Special"]]
        self.c_sheet_sel = ttk.Combobox(f_sheet, values=bend_sheets, state="readonly")
        self.c_sheet_sel.pack(fill="x", pady=5)
        self.c_sheet_sel.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # 2. 規格與 K90
        f_k = tk.LabelFrame(container, text="2. 規格檢索", padx=10, pady=5)
        f_k.pack(fill="x", pady=5)
        
        tk.Label(f_k, text="項目/材質:").grid(row=0, column=0, sticky="w")
        self.c_row = ttk.Combobox(f_k, state="readonly")
        self.c_row.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        tk.Label(f_k, text="板厚 (T):").grid(row=1, column=0, sticky="w")
        self.c_col = ttk.Combobox(f_k, state="readonly")
        self.c_col.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        tk.Label(f_k, text="90° 補償值:").grid(row=2, column=0, sticky="w")
        self.e_k90 = tk.Entry(f_k, font=("Arial", 10, "bold"), fg="blue", bg="#E3F2FD")
        self.e_k90.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        self.c_row.bind("<<ComboboxSelected>>", self.update_k90_val)
        self.c_col.bind("<<ComboboxSelected>>", self.update_k90_val)

        # 3. 尺寸輸入
        f_input = tk.LabelFrame(container, text="3. 輸入尺寸與角度 (內邊法)", padx=10, pady=5)
        f_input.pack(fill="both", expand=True, pady=5)
        
        tk.Label(f_input, text="折彎次數 (n):").pack(side="top", anchor="w")
        self.s_n = tk.Spinbox(f_input, from_=1, to=15, command=self.refresh_bend_ui)
        self.s_n.pack(fill="x", pady=5)
        self.s_n.bind("<KeyRelease>", lambda e: self.refresh_bend_ui())

        self.bend_area = tk.Frame(f_input)
        self.bend_area.pack(fill="both", expand=True)
        self.refresh_bend_ui()

        # 4. 計算
        tk.Button(container, text="執行展開計算", bg="#1976D2", fg="white", font=("Arial", 12, "bold"),
                  command=self.calc_final).pack(fill="x", pady=10)
        
        self.res_box = tk.LabelFrame(container, text="計算報告", padx=10, pady=10)
        self.res_box.pack(fill="x")
        self.l_res = tk.Label(self.res_box, text="等待輸入...", font=("Consolas", 10), justify="left")
        self.l_res.pack(anchor="w")

    def on_sheet_change(self, event):
        sheet = self.c_sheet_sel.get()
        df = self.all_sheets[sheet]
        self.c_row.config(values=df.index.tolist())
        self.c_col.config(values=df.columns.tolist())
        self.c_row.set(""); self.c_col.set(""); self.e_k90.delete(0, tk.END)

    def update_k90_val(self, event):
        s, r, c = self.c_sheet_sel.get(), self.c_row.get(), self.c_col.get()
        if s and r and c:
            val = self.all_sheets[s].loc[r, c]
            self.e_k90.delete(0, tk.END)
            self.e_k90.insert(0, str(val))

    def refresh_bend_ui(self):
        for w in self.bend_area.winfo_children(): w.destroy()
        self.side_entries = []; self.angle_entries = []
        try: n = int(self.s_n.get())
        except: n = 1
        for i in range(n + 1):
            tk.Label(self.bend_area, text=f"內邊長 L{i+1}:", font=("Arial", 9)).pack(anchor="w")
            e_s = tk.Entry(self.bend_area); e_s.pack(fill="x")
            self.side_entries.append(e_s)
            if i < n:
                tk.Label(self.bend_area, text=f"  ↳ 折彎角度 {i+1} (°):", fg="gray", font=("Arial", 9)).pack(anchor="w")
                e_a = tk.Entry(self.bend_area); e_a.insert(0, "90"); e_a.pack(fill="x")
                self.angle_entries.append(e_a)

    def calc_final(self):
        try:
            k90 = float(self.e_k90.get())
            sum_l = sum([float(e.get() or 0) for e in self.side_entries])
            sum_k = 0
            detail = ""
            for i, e in enumerate(self.angle_entries):
                ang = float(e.get() or 90)
                k_adj = (k90/90)*(180-ang)
                sum_k += k_adj
                shape = "鈍角" if ang < 90 else ("直角" if ang == 90 else "銳角")
                detail += f"折{i+1}: {ang}°({shape}) K'={k_adj:.3f}\n"
            
            report = f"{detail}----------\n內邊總和: {sum_l:.2f}\n補償總和: {sum_k:.3f}\n展開總長: {sum_l + sum_k:.3f} mm"
            self.l_res.config(text=report)
        except: messagebox.showerror("錯誤", "請檢查數值輸入")

    # --- 分頁 2: 特殊折彎 ---
    def setup_special_tab(self):
        container = tk.Frame(self.tab_special, padx=30, pady=20)
        container.pack(fill="both", expand=True)
        tk.Label(container, text="特殊工法係數表", font=("Arial", 14, "bold")).pack(pady=10)
        
        df = self.all_sheets.get("Special", pd.DataFrame())
        self.c_sp_type = ttk.Combobox(container, values=df.index.tolist(), state="readonly")
        self.c_sp_type.pack(fill="x", pady=10)
        self.c_sp_thick = ttk.Combobox(container, values=df.columns.tolist(), state="readonly")
        self.c_sp_thick.pack(fill="x", pady=10)
        
        self.l_sp_res = tk.Label(container, text="補償值: --", font=("Arial", 20, "bold"), fg="#D32F2F")
        self.l_sp_res.pack(pady=40)
        
        self.c_sp_type.bind("<<ComboboxSelected>>", self.update_sp)
        self.c_sp_thick.bind("<<ComboboxSelected>>", self.update_sp)

    def update_sp(self, event):
        t, th = self.c_sp_type.get(), self.c_sp_thick.get()
        if t and th:
            val = self.all_sheets["Special"].loc[t, th]
            self.l_sp_res.config(text=f"補償值: {val} mm")

    # --- 分頁 3: 鉚合零件 ---
    def setup_hw_tab(self):
        container = tk.Frame(self.tab_hw, padx=30, pady=20)
        container.pack(fill="both", expand=True)
        tk.Label(container, text="零件開孔規格查詢", font=("Arial", 14, "bold")).pack(pady=10)
        
        df = self.all_sheets.get("Hardware", pd.DataFrame())
        self.c_hw_t = ttk.Combobox(container, values=df.index.tolist(), state="readonly")
        self.c_hw_t.pack(fill="x", pady=10)
        self.c_hw_s = ttk.Combobox(container, values=df.columns.tolist(), state="readonly")
        self.c_hw_s.pack(fill="x", pady=10)
        
        self.l_hw_res = tk.Label(container, text="Ø --", font=("Arial", 36, "bold"), fg="#1976D2")
        self.l_hw_res.pack(pady=40)
        
        self.c_hw_t.bind("<<ComboboxSelected>>", self.update_hw)
        self.c_hw_s.bind("<<ComboboxSelected>>", self.update_hw)

    def update_hw(self, event):
        t, s = self.c_hw_t.get(), self.c_hw_s.get()
        if t and s:
            val = self.all_sheets["Hardware"].loc[t, s]
            self.l_hw_res.config(text=f"Ø {val}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MetalApp(root)
    root.mainloop()
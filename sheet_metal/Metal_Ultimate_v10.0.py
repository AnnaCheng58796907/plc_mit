import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os
import math

class MetalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("板金設計工具箱 v10.0 - 圖形示意版")
        self.root.geometry("1000x900") # 加寬以容納畫布
        
        self.excel_file = "bend_parameters.xlsx"
        self.all_sheets = {}
        self.side_entries = []
        self.angle_entries = []
        
        self.load_all_excel_sheets()
        
        # 建立左右分割視窗 (左側輸入，右側繪圖)
        self.main_paned = tk.PanedWindow(self.root, orient="horizontal")
        self.main_paned.pack(expand=True, fill="both")
        
        self.left_frame = tk.Frame(self.main_paned)
        self.right_frame = tk.Frame(self.main_paned, bg="#E0E0E0")
        self.main_paned.add(self.left_frame, width=450)
        self.main_paned.add(self.right_frame)
        
        # 分頁系統 (放在左側)
        self.notebook = ttk.Notebook(self.left_frame)
        self.notebook.pack(expand=True, fill="both")
        
        self.tab_bend = tk.Frame(self.notebook)
        self.tab_special = tk.Frame(self.notebook)
        self.tab_hw = tk.Frame(self.notebook)
        
        self.notebook.add(self.tab_bend, text=" 📐 展開計算 ")
        self.notebook.add(self.tab_special, text=" 🌀 特殊折彎 ")
        self.notebook.add(self.tab_hw, text=" 🔩 零件查詢 ")
        
        self.setup_bend_tab()
        self.setup_special_tab()
        self.setup_hw_tab()
        
        # 右側畫布設定
        tk.Label(self.right_frame, text="側視圖示意 (非精密比例)", bg="#E0E0E0", font=("Arial", 10)).pack(pady=5)
        self.canvas = tk.Canvas(self.right_frame, bg="white", highlightthickness=2, highlightbackground="#999")
        self.canvas.pack(expand=True, fill="both", padx=10, pady=10)

    def load_all_excel_sheets(self):
        if os.path.exists(self.excel_file):
            try:
                xls = pd.ExcelFile(self.excel_file)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name, index_col=0)
                    df.columns = [str(col) for col in df.columns]
                    self.all_sheets[sheet_name] = df
            except Exception as e:
                messagebox.showerror("Excel 錯誤", str(e))
        else:
            self.all_sheets = {"範例表": pd.DataFrame()}

    def setup_bend_tab(self):
        container = tk.Frame(self.tab_bend, padx=10, pady=10)
        container.pack(fill="both", expand=True)

        # 1. 選擇工作表
        f_sheet = tk.LabelFrame(container, text="1. 選擇係數來源")
        f_sheet.pack(fill="x", pady=5)
        bend_sheets = [s for s in self.all_sheets.keys() if s not in ["Hardware", "Special"]]
        self.c_sheet_sel = ttk.Combobox(f_sheet, values=bend_sheets, state="readonly")
        self.c_sheet_sel.pack(fill="x", padx=5, pady=5)
        self.c_sheet_sel.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # 2. 規格
        f_k = tk.LabelFrame(container, text="2. K90 檢索")
        f_k.pack(fill="x", pady=5)
        self.c_row = ttk.Combobox(f_k, state="readonly")
        self.c_row.pack(fill="x", padx=5)
        self.c_col = ttk.Combobox(f_k, state="readonly")
        self.c_col.pack(fill="x", padx=5, pady=5)
        self.e_k90 = tk.Entry(f_k, font=("Arial", 10, "bold"), fg="blue")
        self.e_k90.pack(fill="x", padx=5, pady=5)
        self.c_row.bind("<<ComboboxSelected>>", self.update_k90_val)
        self.c_col.bind("<<ComboboxSelected>>", self.update_k90_val)

        # 3. 尺寸輸入
        f_input = tk.LabelFrame(container, text="3. 輸入內邊長度與角度")
        f_input.pack(fill="both", expand=True, pady=5)
        self.s_n = tk.Spinbox(f_input, from_=1, to=10, command=self.refresh_bend_ui)
        self.s_n.pack(fill="x", pady=5)
        
        self.bend_area = tk.Frame(f_input)
        self.bend_area.pack(fill="both", expand=True)
        self.refresh_bend_ui()

        # 按鈕
        tk.Button(container, text="執行計算並繪圖", bg="#1976D2", fg="white", font=("Arial", 12, "bold"),
                  command=self.calculate_and_draw).pack(fill="x", pady=5)
        
        self.l_res = tk.Label(container, text="展開長度: --", font=("Arial", 12, "bold"), fg="red")
        self.l_res.pack(pady=5)

    def refresh_bend_ui(self):
        for w in self.bend_area.winfo_children(): w.destroy()
        self.side_entries = []; self.angle_entries = []
        try: n = int(self.s_n.get())
        except: n = 1
        for i in range(n + 1):
            tk.Label(self.bend_area, text=f"L{i+1}:").pack(side="top", anchor="w")
            e_s = tk.Entry(self.bend_area); e_s.insert(0, "50"); e_s.pack(fill="x")
            self.side_entries.append(e_s)
            if i < n:
                tk.Label(self.bend_area, text=f"Angle {i+1}:", fg="gray").pack(side="top", anchor="w")
                e_a = tk.Entry(self.bend_area); e_a.insert(0, "90"); e_a.pack(fill="x")
                self.angle_entries.append(e_a)

    def calculate_and_draw(self):
        self.canvas.delete("all")
        try:
            # 1. 計算邏輯
            k90 = float(self.e_k90.get())
            sum_l = sum([float(e.get() or 0) for e in self.side_entries])
            sum_k = sum([(k90/90)*(180-float(e.get() or 90)) for e in self.angle_entries])
            self.l_res.config(text=f"展開總長: {sum_l + sum_k:.3f} mm")

            # 2. 繪圖邏輯 (側視圖)
            x, y = 100, 400 # 起始座標
            current_angle = 0 # 初始方向 (向右)
            scale = 2 # 縮放比例 (1mm = 2px)

            for i, e_s in enumerate(self.side_entries):
                length = float(e_s.get() or 0) * scale
                # 計算終點
                rad = math.radians(current_angle)
                nx = x + length * math.cos(rad)
                ny = y - length * math.sin(rad) # 畫布 Y 軸向下增加，所以要用減
                
                # 畫線
                self.canvas.create_line(x, y, nx, ny, width=3, fill="#333")
                self.canvas.create_text(x + (nx-x)/2, y + (ny-y)/2 - 10, text=f"L{i+1}")
                
                x, y = nx, ny
                
                if i < len(self.angle_entries):
                    # 更新角度 (折彎角度是指板材轉折的角度)
                    # 這裡是簡化的示意：90度代表轉向垂直
                    bend_angle = float(self.angle_entries[i].get() or 90)
                    current_angle += (180 - bend_angle) 

        except Exception as e:
            messagebox.showerror("繪圖錯誤", "請檢查數值輸入")

    # (其他分頁功能省略，保持與上一版一致)
    def on_sheet_change(self, e): 
        s=self.c_sheet_sel.get(); df=self.all_sheets[s]
        self.c_row.config(values=df.index.tolist()); self.c_col.config(values=df.columns.tolist())
    def update_k90_val(self, e):
        s,r,c=self.c_sheet_sel.get(),self.c_row.get(),self.c_col.get()
        if s and r and c: self.e_k90.delete(0,tk.END); self.e_k90.insert(0,str(self.all_sheets[s].loc[r,c]))
    def setup_special_tab(self): tk.Label(self.tab_special, text="特殊折彎分頁").pack()
    def setup_hw_tab(self): tk.Label(self.tab_hw, text="零件查詢分頁").pack()

if __name__ == "__main__":
    root = tk.Tk(); app = MetalApp(root); root.mainloop()
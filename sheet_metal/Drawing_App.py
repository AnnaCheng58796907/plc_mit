import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

class DrawingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("圖號編碼與製程管理系統")
        self.root.geometry("800x500")
        
        self.csv_file = 'process_codes.csv'
        self.load_data()
        
        # 建立分頁
        notebook = ttk.Notebook(root)
        notebook.pack(expand=True, fill='both')
        
        self.frame_main = ttk.Frame(notebook)
        self.frame_admin = ttk.Frame(notebook)
        
        notebook.add(self.frame_main, text="圖號產生器")
        notebook.add(self.frame_admin, text="製程代碼管理")
        
        self.setup_main_ui()
        self.setup_admin_ui()

    def load_data(self):
        """讀取或初始化 CSV 檔案"""
        if not os.path.exists(self.csv_file):
            self.df = pd.DataFrame({'代碼': ['AS', 'PC', 'LC', 'MA', 'QC', 'SA'], 
                                    '名稱': ['組裝', '粉烤', '液烤', '素材', '品檢', '半成品']})
            self.save_to_csv()
        else:
            self.df = pd.read_csv(self.csv_file)

    def save_to_csv(self):
        self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')

    # --- 介面 1: 圖號產生器 ---
    def setup_main_ui(self):
        padding = {'padx': 10, 'pady': 5}
        
        # 輸入區
        input_group = ttk.LabelFrame(self.frame_main, text="編碼輸入")
        input_group.pack(fill="x", **padding)

        # 排版使用 Grid
        ttk.Label(input_group, text="客戶(5碼):").grid(row=0, column=0, sticky='e')
        self.ent_cust = ttk.Entry(input_group, width=10)
        self.ent_cust.grid(row=0, column=1, sticky='w')

        ttk.Label(input_group, text="來源:").grid(row=0, column=2, sticky='e')
        self.var_source = tk.StringVar(value="A")
        ttk.Radiobutton(input_group, text="A廠內", variable=self.var_source, value="A").grid(row=0, column=3)
        ttk.Radiobutton(input_group, text="B委外", variable=self.var_source, value="B").grid(row=0, column=4)

        ttk.Label(input_group, text="成品碼(4碼):").grid(row=1, column=0, sticky='e')
        self.ent_prod = ttk.Entry(input_group, width=10)
        self.ent_prod.grid(row=1, column=1, sticky='w')

        ttk.Label(input_group, text="製程(2碼):").grid(row=1, column=2, sticky='e')
        self.ent_proc = ttk.Entry(input_group, width=10)
        self.ent_proc.grid(row=1, column=3, sticky='w')

        # 階層區
        lvl_group = ttk.LabelFrame(self.frame_main, text="階層流水號 (L1-L2-L3)")
        lvl_group.pack(fill="x", **padding)

        self.var_is_fg = tk.BooleanVar()
        ttk.Checkbutton(lvl_group, text="業務成品階 (FG0000)", variable=self.var_is_fg, command=self.toggle_fg).grid(row=0, column=0, columnspan=2)

        self.ent_l1 = ttk.Entry(lvl_group, width=5); self.ent_l1.grid(row=0, column=2)
        self.ent_l2 = ttk.Entry(lvl_group, width=5); self.ent_l2.grid(row=0, column=3)
        self.ent_l3 = ttk.Entry(lvl_group, width=5); self.ent_l3.grid(row=0, column=4)
        
        ttk.Label(lvl_group, text="版本:").grid(row=0, column=5)
        self.ent_ver = ttk.Entry(lvl_group, width=3); self.ent_ver.insert(0, "1"); self.ent_ver.grid(row=0, column=6)

        # 結果按鈕
        ttk.Button(self.frame_main, text="生成圖號並複製", command=self.generate).pack(pady=10)
        self.lbl_result = ttk.Label(self.frame_main, text="等待生成...", font=("Consolas", 14, "bold"), foreground="blue")
        self.lbl_result.pack()

    def toggle_fg(self):
        state = 'disabled' if self.var_is_fg.get() else 'normal'
        self.ent_l1.config(state=state)
        self.ent_l2.config(state=state)
        self.ent_l3.config(state=state)

    def generate(self):
        cust = self.ent_cust.get().upper().zfill(5)[:5]
        src = self.var_source.get()
        prod = self.ent_prod.get().zfill(4)[:4]
        
        # 製程代碼判斷：不在清單就用 SA
        proc_input = self.ent_proc.get().upper().strip()
        final_proc = proc_input if proc_input in self.df['代碼'].values else "SA"
        
        if self.var_is_fg.get():
            mid = "FG0000"
        else:
            l1 = self.ent_l1.get().zfill(2)[:2]
            l2 = self.ent_l2.get().zfill(2)[:2]
            l3 = self.ent_l3.get().zfill(2)[:2]
            mid = f"{l1}{l2}{l3}"
            
        ver = self.ent_ver.get().upper()[:1]
        
        res = f"{cust}{src}{prod}{final_proc}{mid}{ver}"
        self.lbl_result.config(text=res)
        self.root.clipboard_clear()
        self.root.clipboard_append(res)
        
        if final_proc == "SA" and proc_input != "SA":
            messagebox.showinfo("提醒", f"代碼 '{proc_input}' 不在清單中，已自動改為 SA")

    # --- 介面 2: 管理區 ---
    def setup_admin_ui(self):
        # 表格顯示
        self.tree = ttk.Treeview(self.frame_admin, columns=("ID", "Name"), show='headings')
        self.tree.heading("ID", text="製程代碼")
        self.tree.heading("Name", text="製程名稱")
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.refresh_tree()

        # 編輯區
        edit_frame = ttk.Frame(self.frame_admin)
        edit_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(edit_frame, text="代碼:").grid(row=0, column=0)
        self.new_code = ttk.Entry(edit_frame, width=10)
        self.new_code.grid(row=0, column=1)
        
        ttk.Label(edit_frame, text="名稱:").grid(row=0, column=2)
        self.new_name = ttk.Entry(edit_frame, width=15)
        self.new_name.grid(row=0, column=3)
        
        ttk.Button(edit_frame, text="新增/更新", command=self.admin_add).grid(row=0, column=4, padx=5)
        ttk.Button(edit_frame, text="刪除選中", command=self.admin_del).grid(row=0, column=5)

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=(row['代碼'], row['名稱']))

    def admin_add(self):
        code = self.new_code.get().upper().strip()
        name = self.new_name.get().strip()
        if not code or not name: return
        
        # 若代碼存在則更新，不存在則新增
        self.df = self.df[self.df['代碼'] != code]
        new_row = pd.DataFrame({'代碼': [code], '名稱': [name]})
        self.df = pd.concat([self.df, new_row], ignore_index=True)
        self.save_to_csv()
        self.refresh_tree()
        
    def admin_del(self):
        selected = self.tree.selection()
        if not selected: return
        code = self.tree.item(selected[0])['values'][0]
        self.df = self.df[self.df['代碼'] != code]
        self.save_to_csv()
        self.refresh_tree()

if __name__ == "__main__":
    root = tk.Tk()
    app = DrawingApp(root)
    root.mainloop()
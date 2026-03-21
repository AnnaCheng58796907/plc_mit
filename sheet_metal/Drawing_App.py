import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

class DrawingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("圖號編碼與製程管理系統 v2.0")
        self.root.geometry("850x550")
        
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
            # 初始預設值
            data = {
                '代碼': ['AS', 'PC', 'LC', 'MA', 'QC', 'SA'], 
                '名稱': ['組裝(Assembly)', '粉烤(Powder)', '液烤(Liquid)', '素材(Material)', '品檢(QC)', '半成品(Sub-Assy)']
            }
            self.df = pd.DataFrame(data)
            self.save_to_csv()
        else:
            self.df = pd.read_csv(self.csv_file)

    def save_to_csv(self):
        self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')

    # --- 介面 1: 圖號產生器 ---
    def setup_main_ui(self):
        padding = {'padx': 10, 'pady': 8}
        
        # 1. 基礎資訊區
        base_group = ttk.LabelFrame(self.frame_main, text="基礎資訊")
        base_group.pack(fill="x", **padding)

        ttk.Label(base_group, text="客戶(5碼):").grid(row=0, column=0, sticky='e', padx=5)
        self.ent_cust = ttk.Entry(base_group, width=12)
        self.ent_cust.grid(row=0, column=1, sticky='w')

        ttk.Label(base_group, text="來源:").grid(row=0, column=2, sticky='e', padx=5)
        self.var_source = tk.StringVar(value="A")
        ttk.Radiobutton(base_group, text="A 廠內", variable=self.var_source, value="A").grid(row=0, column=3)
        ttk.Radiobutton(base_group, text="B 委外", variable=self.var_source, value="B").grid(row=0, column=4)

        ttk.Label(base_group, text="成品碼(4碼):").grid(row=1, column=0, sticky='e', padx=5)
        self.ent_prod = ttk.Entry(base_group, width=12)
        self.ent_prod.grid(row=1, column=1, sticky='w')

        # 重點：製程下拉選單
        ttk.Label(base_group, text="選擇製程:").grid(row=1, column=2, sticky='e', padx=5)
        self.cb_proc = ttk.Combobox(base_group, width=25, state="readonly")
        self.cb_proc.grid(row=1, column=3, columnspan=2, sticky='w')
        self.update_combobox() # 初始化下拉清單內容

        # 2. 階層流水號區
        lvl_group = ttk.LabelFrame(self.frame_main, text="階層與版本控制")
        lvl_group.pack(fill="x", **padding)

        self.var_is_fg = tk.BooleanVar()
        ttk.Checkbutton(lvl_group, text="業務成品 (自動 FG0000)", variable=self.var_is_fg, command=self.toggle_fg).grid(row=0, column=0, columnspan=2, pady=5)

        ttk.Label(lvl_group, text="L1(零件):").grid(row=1, column=0)
        self.ent_l1 = ttk.Entry(lvl_group, width=5); self.ent_l1.grid(row=1, column=1)
        ttk.Label(lvl_group, text="L2(次階):").grid(row=1, column=2)
        self.ent_l2 = ttk.Entry(lvl_group, width=5); self.ent_l2.grid(row=1, column=3)
        ttk.Label(lvl_group, text="L3(末階):").grid(row=1, column=4)
        self.ent_l3 = ttk.Entry(lvl_group, width=5); self.ent_l3.grid(row=1, column=5)
        
        ttk.Label(lvl_group, text="版本:").grid(row=1, column=6, padx=10)
        self.ent_ver = ttk.Entry(lvl_group, width=5); self.ent_ver.insert(0, "1"); self.ent_ver.grid(row=1, column=7)

        # 3. 結果區
        ttk.Button(self.frame_main, text="生成圖號並自動複製", command=self.generate).pack(pady=20)
        self.lbl_result = ttk.Label(self.frame_main, text="請輸入資料後點擊生成", font=("Consolas", 16, "bold"), foreground="#2C3E50")
        self.lbl_result.pack()

    def update_combobox(self):
        """同步 CSV 的代碼到下拉選單"""
        options = [f"{row['代碼']} - {row['名稱']}" for _, row in self.df.iterrows()]
        self.cb_proc['values'] = options
        if options: self.cb_proc.current(0)

    def toggle_fg(self):
        state = 'disabled' if self.var_is_fg.get() else 'normal'
        for entry in [self.ent_l1, self.ent_l2, self.ent_l3]:
            entry.config(state=state)

    def generate(self):
        try:
            cust = self.ent_cust.get().upper().strip().zfill(5)[:5]
            src = self.var_source.get()
            prod = self.ent_prod.get().strip().zfill(4)[:4]
            
            # 取得下拉選單選中的代碼 (取前兩碼)
            selected_proc = self.cb_proc.get()
            proc_code = selected_proc.split(" - ")[0] if selected_proc else "SA"
            
            if self.var_is_fg.get():
                mid = "FG0000"
            else:
                l1 = self.ent_l1.get().strip().zfill(2)[:2]
                l2 = self.ent_l2.get().strip().zfill(2)[:2]
                l3 = self.ent_l3.get().strip().zfill(2)[:2]
                mid = f"{l1}{l2}{l3}"
                
            ver = self.ent_ver.get().upper().strip()[:1]
            
            final_res = f"{cust}{src}{prod}{proc_code}{mid}{ver}"
            
            # 長度檢查
            if len(final_res) != 19:
                messagebox.showwarning("長度提醒", f"生成的圖號長度為 {len(final_res)} 碼，請檢查輸入是否完整。")

            self.lbl_result.config(text=final_res)
            self.root.clipboard_clear()
            self.root.clipboard_append(final_res)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"生成失敗：{str(e)}")

    # --- 介面 2: 管理區 ---
    def setup_admin_ui(self):
        frame = ttk.Frame(self.frame_admin, padding=10)
        frame.pack(fill="both", expand=True)

        # 表格
        self.tree = ttk.Treeview(frame, columns=("ID", "Name"), show='headings', height=12)
        self.tree.heading("ID", text="代碼 (2碼)")
        self.tree.heading("Name", text="製程名稱/說明")
        self.tree.column("ID", width=100)
        self.tree.pack(fill="both", expand=True)
        self.refresh_tree()

        # 編輯工具列
        edit_frame = ttk.LabelFrame(frame, text="編輯清單", padding=10)
        edit_frame.pack(fill="x", pady=10)
        
        ttk.Label(edit_frame, text="代碼:").grid(row=0, column=0)
        self.new_code = ttk.Entry(edit_frame, width=8)
        self.new_code.grid(row=0, column=1, padx=5)
        
        ttk.Label(edit_frame, text="名稱說明:").grid(row=0, column=2)
        self.new_name = ttk.Entry(edit_frame, width=20)
        self.new_name.grid(row=0, column=3, padx=5)
        
        ttk.Button(edit_frame, text="新增/更新代碼", command=self.admin_add).grid(row=0, column=4, padx=2)
        ttk.Button(edit_frame, text="刪除選中項目", command=self.admin_del).grid(row=0, column=5, padx=2)

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        # 排序讓列表整齊
        self.df = self.df.sort_values(by='代碼')
        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=(row['代碼'], row['名稱']))
        self.update_combobox() # 同步更新第一頁的下拉選單

    def admin_add(self):
        code = self.new_code.get().upper().strip()[:2]
        name = self.new_name.get().strip()
        if not code or not name:
            messagebox.showwarning("輸入錯誤", "代碼與名稱不可為空")
            return
        
        # 覆蓋舊有的同代碼資料
        self.df = self.df[self.df['代碼'] != code]
        new_row = pd.DataFrame({'代碼': [code], '名稱': [name]})
        self.df = pd.concat([self.df, new_row], ignore_index=True)
        self.save_to_csv()
        self.refresh_tree()
        self.new_code.delete(0, tk.END)
        self.new_name.delete(0, tk.END)
        
    def admin_del(self):
        selected = self.tree.selection()
        if not selected: return
        if messagebox.askyesno("確認", "確定要刪除此製程代碼嗎？"):
            code = self.tree.item(selected[0])['values'][0]
            self.df = self.df[self.df['代碼'] != code]
            self.save_to_csv()
            self.refresh_tree()

if __name__ == "__main__":
    root = tk.Tk()
    # 設定整體風格
    style = ttk.Style()
    style.configure("TLabel", font=("Microsoft JhengHei", 10))
    style.configure("TButton", font=("Microsoft JhengHei", 10))
    
    app = DrawingApp(root)
    root.mainloop()
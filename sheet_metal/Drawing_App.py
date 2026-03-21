import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

class DrawingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("圖號編碼管理系統 v3.0 (含連字號格式)")
        self.root.geometry("900x600")
        
        self.csv_file = 'process_codes.csv'
        self.load_data()
        
        notebook = ttk.Notebook(root)
        notebook.pack(expand=True, fill='both')
        
        self.frame_main = ttk.Frame(notebook)
        self.frame_admin = ttk.Frame(notebook)
        
        notebook.add(self.frame_main, text="圖號產生器")
        notebook.add(self.frame_admin, text="製程代碼管理")
        
        self.setup_main_ui()
        self.setup_admin_ui()

    def load_data(self):
        if not os.path.exists(self.csv_file):
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

    def setup_main_ui(self):
        padding = {'padx': 15, 'pady': 10}
        
        # 1. 輸入區
        base_group = ttk.LabelFrame(self.frame_main, text="基礎編碼資訊")
        base_group.pack(fill="x", **padding)

        ttk.Label(base_group, text="客戶(5碼):").grid(row=0, column=0, sticky='e', padx=5)
        self.ent_cust = ttk.Entry(base_group, width=15)
        self.ent_cust.grid(row=0, column=1, sticky='w')

        ttk.Label(base_group, text="製作來源:").grid(row=0, column=2, sticky='e', padx=5)
        self.var_source = tk.StringVar(value="A")
        ttk.Radiobutton(base_group, text="A 廠內", variable=self.var_source, value="A").grid(row=0, column=3)
        ttk.Radiobutton(base_group, text="B 委外", variable=self.var_source, value="B").grid(row=0, column=4)

        ttk.Label(base_group, text="成品碼(4碼):").grid(row=1, column=0, sticky='e', padx=5)
        self.ent_prod = ttk.Entry(base_group, width=15)
        self.ent_prod.grid(row=1, column=1, sticky='w')

        ttk.Label(base_group, text="製程代碼:").grid(row=1, column=2, sticky='e', padx=5)
        self.cb_proc = ttk.Combobox(base_group, width=25, state="readonly")
        self.cb_proc.grid(row=1, column=3, columnspan=2, sticky='w')
        self.update_combobox()

        # 2. 階層區
        lvl_group = ttk.LabelFrame(self.frame_main, text="階層流水號 (L1-L2-L3)")
        lvl_group.pack(fill="x", **padding)

        self.var_is_fg = tk.BooleanVar()
        ttk.Checkbutton(lvl_group, text="業務成品 (13-18碼固定為 FG0000)",
                        variable=self.var_is_fg,
                        command=self.toggle_fg).grid(row=0, column=0, columnspan=2, pady=5)

        ttk.Label(lvl_group, text="L1:").grid(row=1, column=0); self.ent_l1 = ttk.Entry(lvl_group, width=6); self.ent_l1.grid(row=1, column=1)
        ttk.Label(lvl_group, text="L2:").grid(row=1, column=2); self.ent_l2 = ttk.Entry(lvl_group, width=6); self.ent_l2.grid(row=1, column=3)
        ttk.Label(lvl_group, text="L3:").grid(row=1, column=4); self.ent_l3 = ttk.Entry(lvl_group, width=6); self.ent_l3.grid(row=1, column=5)
        
        ttk.Label(lvl_group, text="版本:").grid(row=1, column=6, padx=10)
        self.ent_ver = ttk.Entry(lvl_group, width=6); self.ent_ver.insert(0, "0"); self.ent_ver.grid(row=1, column=7)

        # 3. 生成按鈕與大字顯示結果
        ttk.Button(self.frame_main, text="生成圖號並自動複製 (含連字號)", command=self.generate).pack(pady=20)
        
        result_frame = ttk.Frame(self.frame_main)
        result_frame.pack(fill="x", padx=20)
        
        self.lbl_result = tk.Label(result_frame, text="等待輸入...", font=("Consolas", 22, "bold"), fg="#E74C3C", bg="#FDFEFE", relief="sunken", pady=10)
        self.lbl_result.pack(fill="x")

    def update_combobox(self):
        options = [f"{row['代碼']} - {row['名稱']}" for _, row in self.df.iterrows()]
        self.cb_proc['values'] = options
        if options: self.cb_proc.current(0)

    def toggle_fg(self):
        state = 'disabled' if self.var_is_fg.get() else 'normal'
        for e in [self.ent_l1, self.ent_l2, self.ent_l3]: e.config(state=state)

    def generate(self):
        try:
            # 1-5碼 客戶
            cust = self.ent_cust.get().upper().strip().zfill(5)[:5]
            # 6碼 來源
            src = self.var_source.get()
            # 7-10碼 成品
            prod = self.ent_prod.get().strip().zfill(4)[:4]
            # 11-12碼 製程
            selected_proc = self.cb_proc.get()
            proc_code = selected_proc.split(" - ")[0] if selected_proc else "SA"
            
            # 13-18碼 零件階層
            if self.var_is_fg.get():
                mid = "FG0000"
            else:
                l1 = self.ent_l1.get().strip().zfill(2)[:2]
                l2 = self.ent_l2.get().strip().zfill(2)[:2]
                l3 = self.ent_l3.get().strip().zfill(2)[:2]
                mid = f"{l1}{l2}{l3}"
            
            # 19碼 版本
            ver = self.ent_ver.get().upper().strip()[:1] or "0"

            # --- 關鍵：格式化輸出 ---
            # 規則：客戶(5) - 來源(1)成品(4) - 製程(2)零件(6) - 版本(1)
            # 例如: A0000-A0000-FG000000-0
            formatted_res = f"{cust}-{src}{prod}-{proc_code}{mid}-{ver}"

            self.lbl_result.config(text=formatted_res)
            self.root.clipboard_clear()
            self.root.clipboard_append(formatted_res)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"生成失敗：{str(e)}")

    def setup_admin_ui(self):
        frame = ttk.Frame(self.frame_admin, padding=15)
        frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(frame, columns=("ID", "Name"), show='headings', height=15)
        self.tree.heading("ID", text="代碼 (2碼)"); self.tree.heading("Name", text="名稱說明")
        self.tree.column("ID", width=100); self.tree.pack(fill="both", expand=True)
        self.refresh_tree()
        edit_frame = ttk.LabelFrame(frame, text="管理製程清單", padding=10)
        edit_frame.pack(fill="x", pady=15)
        ttk.Label(edit_frame, text="代碼:").grid(row=0, column=0)
        self.new_code = ttk.Entry(edit_frame, width=10); self.new_code.grid(row=0, column=1, padx=5)
        ttk.Label(edit_frame, text="名稱:").grid(row=0, column=2)
        self.new_name = ttk.Entry(edit_frame, width=25); self.new_name.grid(row=0, column=3, padx=5)
        ttk.Button(edit_frame, text="儲存/更新", command=self.admin_add).grid(row=0, column=4, padx=5)
        ttk.Button(edit_frame, text="刪除", command=self.admin_del).grid(row=0, column=5)

    def refresh_tree(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        self.df = self.df.sort_values(by='代碼')
        for _, row in self.df.iterrows(): self.tree.insert("", "end", values=(row['代碼'], row['名稱']))
        self.update_combobox()

    def admin_add(self):
        code, name = self.new_code.get().upper().strip()[:2], self.new_name.get().strip()
        if not code or not name: return
        self.df = self.df[self.df['代碼'] != code]
        self.df = pd.concat([self.df, pd.DataFrame({'代碼': [code], '名稱': [name]})], ignore_index=True)
        self.save_to_csv(); self.refresh_tree()
        self.new_code.delete(0, tk.END); self.new_name.delete(0, tk.END)

    def admin_del(self):
        selected = self.tree.selection()
        if not selected: return
        if messagebox.askyesno("確認", "確定刪除？"):
            code = self.tree.item(selected[0])['values'][0]
            self.df = self.df[self.df['代碼'] != code]
            self.save_to_csv(); self.refresh_tree()

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(); style.configure("TLabel", font=("Microsoft JhengHei", 10))
    app = DrawingApp(root)
    root.mainloop()
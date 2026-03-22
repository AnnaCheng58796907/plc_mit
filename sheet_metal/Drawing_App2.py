import sys
import os
import csv
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QRadioButton, 
                             QCheckBox, QTabWidget, QTableWidget, QTableWidgetItem, 
                             QMessageBox, QGroupBox, QFileDialog, QHeaderView)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIntValidator

class DrawingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # 核心優化：使用 set 進行極速比對
        self.registry_set = set()
        self.temp_records = [] 
        self.csv_file = 'process_codes.csv'
        
        self.init_data() 
        self.init_ui()
        self.refresh_admin_table()

    def init_data(self):
        """讀取製程代碼 CSV"""
        if not os.path.exists(self.csv_file):
            self.proc_data = [
                ['AS', '組裝'], ['PC', '粉烤'], ['LC', '液烤'], 
                ['MA', '素材'], ['QC', '品檢'], ['SA', '半成品']
            ]
            self.save_proc_to_csv()
        else:
            self.proc_data = []
            with open(self.csv_file, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                next(reader, None) 
                for row in reader:
                    if row: self.proc_data.append(row)
        self.proc_data.sort()

    def init_ui(self):
        self.setWindowTitle("圖號編碼管理系統")
        self.setGeometry(100, 100, 1100, 850)
        self.setFont(QFont("Microsoft JhengHei", 12))

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.tab_gen = QWidget()
        self.tab_adm = QWidget()
        self.tabs.addTab(self.tab_gen, "圖號產生器")
        self.tabs.addTab(self.tab_adm, "製程代碼管理")

        self.setup_generator_tab()
        self.setup_admin_tab()
        self.update_combobox()

    def setup_generator_tab(self):
        layout = QVBoxLayout(self.tab_gen)
        num_v = QIntValidator()

        # --- 輸入區 ---
        box = QGroupBox("編碼輸入規範")
        grid = QVBoxLayout()
        
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("客戶碼: A"))
        self.input_cust = QLineEdit(); self.input_cust.setFixedWidth(100); self.input_cust.setValidator(num_v); self.input_cust.setMaxLength(4)
        row1.addWidget(self.input_cust); row1.addSpacing(30)
        self.rb_a = QRadioButton("廠內(A)"); self.rb_a.setChecked(True); self.rb_b = QRadioButton("委外(B)")
        row1.addWidget(self.rb_a); row1.addWidget(self.rb_b); row1.addStretch()
        grid.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("成品碼:  "))
        self.input_prod = QLineEdit(); self.input_prod.setFixedWidth(100); self.input_prod.setValidator(num_v); self.input_prod.setMaxLength(4)
        row2.addWidget(self.input_prod); row2.addSpacing(30); row2.addWidget(QLabel("製程:"))
        self.cb_proc = QComboBox(); self.cb_proc.setMinimumWidth(250)
        row2.addWidget(self.cb_proc); row2.addStretch()
        grid.addLayout(row2)

        row3 = QHBoxLayout()
        self.chk_fg = QCheckBox("業務成品(FG)"); self.chk_fg.toggled.connect(self.on_fg_toggled)
        row3.addWidget(self.chk_fg); row3.addSpacing(20)
        self.l1 = QLineEdit(); self.l2 = QLineEdit(); self.l3 = QLineEdit()
        for e in [self.l1, self.l2, self.l3]: e.setFixedWidth(45); e.setValidator(num_v); e.setMaxLength(2); row3.addWidget(e)
        row3.addSpacing(20); row3.addWidget(QLabel("版本:"))
        self.ver = QLineEdit("0"); self.ver.setFixedWidth(45); self.ver.setMaxLength(1)
        row3.addWidget(self.ver); row3.addStretch()
        grid.addLayout(row3)
        box.setLayout(grid)
        layout.addWidget(box)

        # --- 生成按鈕 ---
        self.btn_gen = QPushButton("確認生成圖號並加入清單")
        self.btn_gen.setFixedHeight(65)
        self.btn_gen.setStyleSheet("background-color: #27ae60; color: white; font-size: 22px; font-weight: bold; border-radius: 5px;")
        self.btn_gen.clicked.connect(self.handle_generate)
        layout.addWidget(self.btn_gen)

        # --- 表格區 ---
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["生成時間", "完整圖號結果", "自定義備註"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.itemChanged.connect(self.sync_note)
        layout.addWidget(self.table)

        # --- 功能按鈕區 ---
        btn_layout = QHBoxLayout()
        btn_del_rec = QPushButton("刪除選中紀錄"); btn_del_rec.setStyleSheet("background-color: #e67e22; color: white; padding: 10px;")
        btn_del_rec.clicked.connect(self.delete_record)
        
        # 新增：全部清空按鈕
        btn_clear_all = QPushButton("清空全部紀錄"); btn_clear_all.setStyleSheet("background-color: #c0392b; color: white; padding: 10px;")
        btn_clear_all.clicked.connect(self.clear_all_records)
        
        btn_csv = QPushButton("匯出 CSV"); btn_excel = QPushButton("匯出 Excel")
        btn_csv.clicked.connect(lambda: self.export_data('csv'))
        btn_excel.clicked.connect(lambda: self.export_data('xlsx'))
        
        btn_layout.addWidget(btn_del_rec); btn_layout.addWidget(btn_clear_all); btn_layout.addStretch()
        btn_layout.addWidget(btn_csv); btn_layout.addWidget(btn_excel)
        layout.addLayout(btn_layout)

    # --- 邏輯功能 ---
    def handle_generate(self):
        c = f"A{self.input_cust.text().zfill(4)}"
        s = "A" if self.rb_a.isChecked() else "B"
        p = self.input_prod.text().zfill(4)
        pr = self.cb_proc.currentText().split(" - ")[0]
        mid = "FG0000" if self.chk_fg.isChecked() else f"{self.l1.text().zfill(2)}{self.l2.text().zfill(2)}{self.l3.text().zfill(2)}"
        v = self.ver.text().upper() or "0"
        full_code = f"{c}-{s}{p}-{pr}{mid}-{v}"

        if full_code in self.registry_set:
            QMessageBox.critical(self, "重複鎖定", f"圖號已存在：\n{full_code}")
            return

        self.registry_set.add(full_code)
        time_str = datetime.now().strftime("%H:%M:%S")
        self.temp_records.append([time_str, full_code, ""])
        self.refresh_main_table()
        QApplication.clipboard().setText(full_code)

    def refresh_main_table(self):
        self.table.itemChanged.disconnect(self.sync_note)
        self.table.setRowCount(len(self.temp_records))
        for i, (t, c, n) in enumerate(self.temp_records):
            self.table.setItem(i, 0, QTableWidgetItem(t))
            self.table.setItem(i, 1, QTableWidgetItem(c))
            self.table.setItem(i, 2, QTableWidgetItem(n))
        self.table.itemChanged.connect(self.sync_note)

    def delete_record(self):
        row = self.table.currentRow()
        if row >= 0:
            code_to_remove = self.temp_records[row][1]
            self.registry_set.discard(code_to_remove)
            del self.temp_records[row]
            self.refresh_main_table()

    def clear_all_records(self):
        """新增：清除所有暫存圖號與鎖定狀態"""
        if not self.temp_records: return
        confirm = QMessageBox.question(self, "確認清空", "確定要清空清單中所有紀錄嗎？\n(清空後原圖號可重新生成)", 
                                       QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            self.temp_records = []
            self.registry_set.clear()
            self.refresh_main_table()

    def sync_note(self, item):
        if item.column() == 2:
            self.temp_records[item.row()][2] = item.text()

    def export_data(self, ext):
        if not self.temp_records: return
        path, _ = QFileDialog.getSaveFileName(self, f"儲存 {ext.upper()}", "", f"Files (*.{ext})")
        if path:
            if ext == 'csv':
                with open(path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["生成時間", "完整圖號", "備註"])
                    writer.writerows(self.temp_records)
            else:
                import pandas as pd
                pd.DataFrame(self.temp_records, columns=["生成時間", "完整圖號", "備註"]).to_excel(path, index=False)
            QMessageBox.information(self, "成功", f"{ext.upper()} 匯出完成！")

    def on_fg_toggled(self, checked):
        for e in [self.l1, self.l2, self.l3]: e.setDisabled(checked)

    # --- 管理頁面 ---
    def setup_admin_tab(self):
        layout = QVBoxLayout(self.tab_adm)
        self.admin_table = QTableWidget(0, 2)
        self.admin_table.setHorizontalHeaderLabels(["代碼", "名稱"])
        self.admin_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.admin_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.admin_table)

        edit_box = QGroupBox("代碼管理")
        edit_layout = QHBoxLayout()
        self.new_c = QLineEdit(); self.new_c.setPlaceholderText("代碼(2位)"); self.new_c.setMaxLength(2)
        self.new_n = QLineEdit(); self.new_n.setPlaceholderText("名稱")
        btn_add = QPushButton("新增"); btn_add.clicked.connect(self.admin_add)
        btn_del = QPushButton("刪除選中"); btn_del.clicked.connect(self.admin_del)
        
        edit_layout.addWidget(self.new_c); edit_layout.addWidget(self.new_n)
        edit_layout.addWidget(btn_add); edit_layout.addWidget(btn_del)
        edit_box.setLayout(edit_layout)
        layout.addWidget(edit_box)

    def admin_add(self):
        c, n = self.new_c.text().upper().strip(), self.new_n.text().strip()
        if len(c) == 2 and n:
            if any(c == item[0] for item in self.proc_data):
                QMessageBox.warning(self, "錯誤", "代碼重複！")
                return
            self.proc_data.append([c, n])
            self.proc_data.sort()
            self.save_proc_to_csv()
            self.refresh_admin_table()
            self.update_combobox()
            self.new_c.clear(); self.new_n.clear()

    def admin_del(self):
        row = self.admin_table.currentRow()
        if row >= 0:
            del self.proc_data[row]
            self.save_proc_to_csv()
            self.refresh_admin_table()
            self.update_combobox()

    def refresh_admin_table(self):
        self.admin_table.setRowCount(len(self.proc_data))
        for i, (code, name) in enumerate(self.proc_data):
            self.admin_table.setItem(i, 0, QTableWidgetItem(code))
            self.admin_table.setItem(i, 1, QTableWidgetItem(name))

    def update_combobox(self):
        self.cb_proc.clear()
        self.cb_proc.addItems([f"{c} - {n}" for c, n in self.proc_data])

    def save_proc_to_csv(self):
        with open(self.csv_file, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['代碼', '名稱'])
            writer.writerows(self.proc_data)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DrawingApp()
    window.show()
    sys.exit(app.exec_())
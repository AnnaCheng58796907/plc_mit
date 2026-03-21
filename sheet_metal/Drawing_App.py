import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QRadioButton, 
                             QButtonGroup, QCheckBox, QTabWidget, QTableWidget, 
                             QTableWidgetItem, QMessageBox, QGroupBox, QFileDialog, QHeaderView)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class DrawingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("圖號編碼管理系統 v8.0 (18px 專業版)")
        self.setGeometry(100, 100, 1100, 900)
        
        self.csv_file = 'process_codes.csv'
        self.temp_records = []
        self.load_data()
        
        # --- 全域字體設定：18px ---
        self.main_font = QFont("Microsoft JhengHei", 18)
        self.setFont(self.main_font)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        
        self.tabs = QTabWidget()
        self.tab_generator = QWidget()
        self.tab_admin = QWidget()
        self.tabs.addTab(self.tab_generator, "圖號產生器")
        self.tabs.addTab(self.tab_admin, "製程管理")
        self.main_layout.addWidget(self.tabs)
        
        self.init_generator_ui()
        self.init_admin_ui()
        self.update_combobox()

    def load_data(self):
        if not os.path.exists(self.csv_file):
            data = {'代碼': ['AS', 'PC', 'LC', 'MA', 'QC', 'SA'], 
                    '名稱': ['組裝', '粉烤', '液烤', '素材', '品檢', '半成品']}
            self.df = pd.DataFrame(data)
            self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')
        else:
            self.df = pd.read_csv(self.csv_file)

    def init_generator_ui(self):
        layout = QVBoxLayout(self.tab_generator)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # --- 區塊 1: 輸入區 (加大高度與控制寬度) ---
        input_box = QGroupBox("編碼輸入")
        grid = QVBoxLayout()
        grid.setSpacing(15)
        
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("客戶:"))
        self.input_cust = QLineEdit()
        self.input_cust.setFixedWidth(150) # 精簡寬度
        self.input_cust.setFixedHeight(45)
        row1.addWidget(self.input_cust)
        
        row1.addSpacing(30)
        row1.addWidget(QLabel("來源:"))
        self.rb_a = QRadioButton("A廠內"); self.rb_a.setChecked(True)
        self.rb_b = QRadioButton("B委外")
        row1.addWidget(self.rb_a); row1.addWidget(self.rb_b)
        row1.addStretch()
        grid.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("成品:"))
        self.input_prod = QLineEdit()
        self.input_prod.setFixedWidth(150) # 精簡寬度
        self.input_prod.setFixedHeight(45)
        row2.addWidget(self.input_prod)
        
        row2.addSpacing(30)
        row2.addWidget(QLabel("製程:"))
        self.cb_proc = QComboBox()
        self.cb_proc.setMinimumWidth(350)
        self.cb_proc.setFixedHeight(45)
        row2.addWidget(self.cb_proc)
        row2.addStretch()
        grid.addLayout(row2)

        row3 = QHBoxLayout()
        self.check_fg = QCheckBox("成品(FG)")
        self.check_fg.stateChanged.connect(self.toggle_fg)
        row3.addWidget(self.check_fg)
        
        row3.addSpacing(20)
        row3.addWidget(QLabel("L1-2-3:"))
        self.input_l1 = QLineEdit(); self.input_l1.setFixedWidth(60)
        self.input_l2 = QLineEdit(); self.input_l2.setFixedWidth(60)
        self.input_l3 = QLineEdit(); self.input_l3.setFixedWidth(60)
        for edit in [self.input_l1, self.input_l2, self.input_l3]: edit.setFixedHeight(45)
        row3.addWidget(self.input_l1); row3.addWidget(self.input_l2); row3.addWidget(self.input_l3)
        
        row3.addSpacing(20)
        row3.addWidget(QLabel("版本:"))
        self.input_ver = QLineEdit("0"); self.input_ver.setFixedWidth(60); self.input_ver.setFixedHeight(45)
        row3.addWidget(self.input_ver)
        row3.addStretch()
        grid.addLayout(row3)
        input_box.setLayout(grid)
        layout.addWidget(input_box)

        # --- 區塊 2: 生成按鈕 (醒目) ---
        self.btn_gen = QPushButton("★ 生成並存入紀錄 ★")
        self.btn_gen.setFixedHeight(70)
        self.btn_gen.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; border-radius: 10px;")
        self.btn_gen.clicked.connect(self.generate_and_save)
        layout.addWidget(self.btn_gen)

        # --- 區塊 3: 紀錄區 (寬度重新分配) ---
        record_box = QGroupBox("暫存紀錄管理 (18px 表格)")
        record_layout = QVBoxLayout()
        
        self.table_record = QTableWidget()
        self.table_record.setColumnCount(3)
        self.table_record.setHorizontalHeaderLabels(["時間", "完整圖號", "備註"])
        
        # 重要：調整欄位比例
        header = self.table_record.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)       # 時間欄位固定
        self.table_record.setColumnWidth(0, 120)               # 縮小時間
        header.setSectionResizeMode(1, QHeaderView.Stretch)     # 圖號欄位展開 (重點)
        header.setSectionResizeMode(2, QHeaderView.Fixed)       # 備註欄位固定
        self.table_record.setColumnWidth(2, 150)               # 縮小備註
        
        self.table_record.setStyleSheet("font-size: 18px;")
        record_layout.addWidget(self.table_record)
        
        # 功能按鈕
        action_layout = QHBoxLayout()
        btn_del = QPushButton("刪除單項"); btn_del.setStyleSheet("background-color: #f39c12; color: white;")
        btn_clr = QPushButton("全數清空"); btn_clr.setStyleSheet("background-color: #c0392b; color: white;")
        btn_del.clicked.connect(self.delete_selected_record)
        btn_clr.clicked.connect(self.clear_all_records)
        
        btn_csv = QPushButton("匯出 CSV"); btn_excel = QPushButton("匯出 Excel")
        btn_csv.clicked.connect(lambda: self.export_data('csv'))
        btn_excel.clicked.connect(lambda: self.export_data('xlsx'))
        
        for b in [btn_del, btn_clr, btn_csv, btn_excel]: b.setFixedHeight(45); b.setFixedWidth(160)
        
        action_layout.addWidget(btn_del); action_layout.addWidget(btn_clr)
        action_layout.addStretch()
        action_layout.addWidget(btn_csv); action_layout.addWidget(btn_excel)
        record_layout.addLayout(action_layout)
        record_box.setLayout(record_layout)
        layout.addWidget(record_box)

    def toggle_fg(self, state):
        e = state != Qt.Checked
        for i in [self.input_l1, self.input_l2, self.input_l3]: i.setEnabled(e)

    def update_combobox(self):
        self.cb_proc.clear()
        self.cb_proc.addItems([f"{r['代碼']} - {r['名稱']}" for _, r in self.df.iterrows()])

    def generate_and_save(self):
        cust = self.input_cust.text().upper().strip().zfill(5)[:5]
        src = "A" if self.rb_a.isChecked() else "B"
        prod = self.input_prod.text().strip().zfill(4)[:4]
        proc = self.cb_proc.currentText().split(" - ")[0] if self.cb_proc.currentText() else "SA"
        mid = "FG0000" if self.check_fg.isChecked() else f"{self.input_l1.text().strip().zfill(2)[:2]}{self.input_l2.text().strip().zfill(2)[:2]}{self.input_l3.text().strip().zfill(2)[:2]}"
        ver = self.input_ver.text().upper().strip()[:1] or "0"
        
        formatted = f"{cust}-{src}{prod}-{proc}{mid}-{ver}"
        now = datetime.now().strftime("%H:%M:%S")
        self.temp_records.append([now, formatted, ""])
        self.refresh_record_table()
        QApplication.clipboard().setText(formatted)

    def refresh_record_table(self):
        self.table_record.setRowCount(len(self.temp_records))
        for i, (t, c, n) in enumerate(self.temp_records):
            self.table_record.setItem(i, 0, QTableWidgetItem(t))
            item_code = QTableWidgetItem(c)
            item_code.setForeground(Qt.blue) # 讓圖號變藍色易於區分
            self.table_record.setItem(i, 1, item_code)
            self.table_record.setItem(i, 2, QTableWidgetItem(n))

    def delete_selected_record(self):
        row = self.table_record.currentRow()
        if row >= 0: del self.temp_records[row]; self.refresh_record_table()

    def clear_all_records(self):
        if self.temp_records and QMessageBox.question(self, '確認', '清空所有紀錄？', QMessageBox.Yes|QMessageBox.No) == QMessageBox.Yes:
            self.temp_records = []; self.refresh_record_table()

    def export_data(self, f_type):
        if not self.temp_records: return
        path, _ = QFileDialog.getSaveFileName(self, "儲存檔案", "", f"Files (*.{f_type})")
        if path:
            df_e = pd.DataFrame(self.temp_records, columns=["時間", "圖號", "備註"])
            if f_type == 'csv': df_e.to_csv(path, index=False, encoding='utf-8-sig')
            else: df_e.to_excel(path, index=False)
            QMessageBox.information(self, "成功", "匯出完成")

    def init_admin_ui(self):
        layout = QVBoxLayout(self.tab_admin)
        self.admin_table = QTableWidget(); self.admin_table.setColumnCount(2)
        self.admin_table.setHorizontalHeaderLabels(["代碼", "名稱"])
        self.admin_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.admin_table.setStyleSheet("font-size: 18px;")
        layout.addWidget(self.admin_table)
        self.refresh_admin_table()
        
        edit_layout = QHBoxLayout()
        self.new_code = QLineEdit(); self.new_code.setFixedWidth(100); self.new_code.setFixedHeight(45)
        self.new_name = QLineEdit(); self.new_name.setFixedHeight(45)
        btn_add = QPushButton("新增/更新"); btn_add.setFixedHeight(45)
        btn_add.clicked.connect(self.admin_add)
        edit_layout.addWidget(self.new_code); edit_layout.addWidget(self.new_name); edit_layout.addWidget(btn_add)
        layout.addLayout(edit_layout)

    def refresh_admin_table(self):
        self.df = self.df.sort_values(by='代碼')
        self.admin_table.setRowCount(len(self.df))
        for i, row in self.df.iterrows():
            self.admin_table.setItem(i, 0, QTableWidgetItem(str(row['代碼'])))
            self.admin_table.setItem(i, 1, QTableWidgetItem(str(row['名稱'])))
        self.update_combobox()

    def admin_add(self):
        c, n = self.new_code.text().upper()[:2], self.new_name.text().strip()
        if not c or not n: return
        self.df = self.df[self.df['代碼'] != c]
        self.df = pd.concat([self.df, pd.DataFrame({'代碼': [c], '名稱': [n]})], ignore_index=True)
        self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig'); self.refresh_admin_table()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DrawingApp()
    window.show()
    sys.exit(app.exec_())
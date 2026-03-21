import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QRadioButton, 
                             QButtonGroup, QCheckBox, QTabWidget, QTableWidget, 
                             QTableWidgetItem, QMessageBox, QGroupBox, QFileDialog, QHeaderView)
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtGui import QFont, QIntValidator, QRegExpValidator

class DrawingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("圖號編碼管理系統 v12.0")
        self.setGeometry(100, 100, 1100, 850)
        
        self.csv_file = 'process_codes.csv'
        self.temp_records = []
        self.load_data()
        
        # 全域字體 (12px)
        self.general_font = QFont("Microsoft JhengHei", 12)
        self.setFont(self.general_font)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        
        self.tabs = QTabWidget()
        self.tab_generator = QWidget()
        self.tab_admin = QWidget()
        self.tabs.addTab(self.tab_generator, "圖號產生器")
        self.tabs.addTab(self.tab_admin, "製程代碼管理")
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
        layout.setSpacing(10)
        
        num_validator = QIntValidator()
        alphanum_validator = QRegExpValidator(QRegExp("[a-zA-Z0-9]"))

        # --- 輸入區 ---
        input_box = QGroupBox("編碼輸入規範")
        grid = QVBoxLayout()
        
        # 客戶碼 (固定 A + 4 碼數字)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("客戶碼: A"))
        self.input_cust = QLineEdit()
        self.input_cust.setPlaceholderText("4位數字")
        self.input_cust.setFixedWidth(120)
        self.input_cust.setMaxLength(4)
        self.input_cust.setValidator(num_validator)
        row1.addWidget(self.input_cust)
        
        row1.addSpacing(30); row1.addWidget(QLabel("來源:"))
        self.rb_a = QRadioButton("廠內(A)"); self.rb_a.setChecked(True)
        self.rb_b = QRadioButton("委外(B)")
        row1.addWidget(self.rb_a); row1.addWidget(self.rb_b); row1.addStretch()
        grid.addLayout(row1)

        # 成品碼 (4 碼數字)
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("成品碼:  "))
        self.input_prod = QLineEdit()
        self.input_prod.setPlaceholderText("4位數字")
        self.input_prod.setFixedWidth(120)
        self.input_prod.setMaxLength(4)
        self.input_prod.setValidator(num_validator)
        row2.addWidget(self.input_prod)
        
        row2.addSpacing(30); row2.addWidget(QLabel("製程:"))
        self.cb_proc = QComboBox(); self.cb_proc.setMinimumWidth(300)
        row2.addWidget(self.cb_proc); row2.addStretch()
        grid.addLayout(row2)

        # 階層與版本
        row3 = QHBoxLayout()
        self.check_fg = QCheckBox("業務成品(FG)"); self.check_fg.stateChanged.connect(self.toggle_fg)
        row3.addWidget(self.check_fg)
        
        row3.addSpacing(20); row3.addWidget(QLabel("階層(L1-L3):"))
        self.input_l1 = QLineEdit(); self.input_l1.setFixedWidth(50); self.input_l1.setMaxLength(2); self.input_l1.setValidator(num_validator)
        self.input_l2 = QLineEdit(); self.input_l2.setFixedWidth(50); self.input_l2.setMaxLength(2); self.input_l2.setValidator(num_validator)
        self.input_l3 = QLineEdit(); self.input_l3.setFixedWidth(50); self.input_l3.setMaxLength(2); self.input_l3.setValidator(num_validator)
        row3.addWidget(self.input_l1); row3.addWidget(self.input_l2); row3.addWidget(self.input_l3)
        
        row3.addSpacing(20); row3.addWidget(QLabel("版本:"))
        self.input_ver = QLineEdit("0"); self.input_ver.setFixedWidth(50); self.input_ver.setMaxLength(1); self.input_ver.setValidator(alphanum_validator)
        row3.addWidget(self.input_ver); row3.addStretch()
        grid.addLayout(row3)
        
        input_box.setLayout(grid)
        layout.addWidget(input_box)

        # --- 生成按鈕 (加大至 22px) ---
        self.btn_gen = QPushButton("確認生成圖號並加入清單")
        self.btn_gen.setFixedHeight(65)
        self.btn_gen.setStyleSheet("""
            QPushButton {
                background-color: #27ae60; 
                color: white; 
                font-weight: bold; 
                font-size: 22px; 
                border-radius: 8px;
            }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        self.btn_gen.clicked.connect(self.generate_and_save)
        layout.addWidget(self.btn_gen)

        # --- 紀錄區 (18px) ---
        record_box = QGroupBox("暫存紀錄管理 (備註欄位可點擊兩下輸入)")
        record_layout = QVBoxLayout()
        self.table_record = QTableWidget()
        self.table_record.setColumnCount(3)
        self.table_record.setHorizontalHeaderLabels(["生成時間", "完整圖號結果", "自定義備註"])
        
        table_font = QFont("Microsoft JhengHei", 12)
        self.table_record.setFont(table_font)
        self.table_record.horizontalHeader().setFont(table_font)
        self.table_record.verticalHeader().setDefaultSectionSize(45)
        
        header = self.table_record.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed); self.table_record.setColumnWidth(0, 130)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Fixed); self.table_record.setColumnWidth(2, 220)
        
        self.table_record.itemChanged.connect(self.sync_note_to_list)
        record_layout.addWidget(self.table_record)
        
        action_layout = QHBoxLayout()
        btn_del = QPushButton("刪除單筆項目"); btn_clr = QPushButton("清空全部紀錄")
        btn_del.setStyleSheet("background-color: #f39c12; color: white; height: 35px;")
        btn_clr.setStyleSheet("background-color: #c0392b; color: white; height: 35px;")
        btn_del.clicked.connect(self.delete_selected_record)
        btn_clr.clicked.connect(self.clear_all_records)
        
        btn_csv = QPushButton("匯出 CSV 檔"); btn_excel = QPushButton("匯出 Excel 檔")
        btn_csv.clicked.connect(lambda: self.export_data('csv'))
        btn_excel.clicked.connect(lambda: self.export_data('xlsx'))
        
        action_layout.addWidget(btn_del); action_layout.addWidget(btn_clr); action_layout.addStretch()
        action_layout.addWidget(btn_csv); action_layout.addWidget(btn_excel)
        record_layout.addLayout(action_layout)
        record_box.setLayout(record_layout)
        layout.addWidget(record_box)

    def toggle_fg(self, state):
        enabled = state != Qt.Checked
        for e in [self.input_l1, self.input_l2, self.input_l3]: e.setEnabled(enabled)

    def update_combobox(self):
        self.cb_proc.clear()
        self.cb_proc.addItems([f"{row['代碼']} - {row['名稱']}" for _, row in self.df.iterrows()])

    def generate_and_save(self):
        # 防呆與補位邏輯
        cust = f"A{self.input_cust.text().strip().zfill(4)}"
        src = "A" if self.rb_a.isChecked() else "B"
        prod = self.input_prod.text().strip().zfill(4)
        proc = self.cb_proc.currentText().split(" - ")[0] if self.cb_proc.currentText() else "SA"
        
        if self.check_fg.isChecked():
            mid = "FG0000"
        else:
            mid = f"{self.input_l1.text().strip().zfill(2)}{self.input_l2.text().strip().zfill(2)}{self.input_l3.text().strip().zfill(2)}"
        
        ver = self.input_ver.text().upper().strip() or "0"
        formatted = f"{cust}-{src}{prod}-{proc}{mid}-{ver}"
        now = datetime.now().strftime("%H:%M:%S")
        
        self.table_record.itemChanged.disconnect(self.sync_note_to_list)
        self.temp_records.append([now, formatted, ""])
        self.refresh_record_table()
        self.table_record.itemChanged.connect(self.sync_note_to_list)
        QApplication.clipboard().setText(formatted)

    def refresh_record_table(self):
        self.table_record.setRowCount(len(self.temp_records))
        for i, (t, c, n) in enumerate(self.temp_records):
            it_time = QTableWidgetItem(t); it_time.setFlags(it_time.flags() ^ Qt.ItemIsEditable)
            it_code = QTableWidgetItem(c); it_code.setFlags(it_code.flags() ^ Qt.ItemIsEditable)
            it_code.setForeground(Qt.blue)
            it_note = QTableWidgetItem(n)
            self.table_record.setItem(i, 0, it_time)
            self.table_record.setItem(i, 1, it_code)
            self.table_record.setItem(i, 2, it_note)

    def sync_note_to_list(self, item):
        if item.column() == 2:
            row = item.row()
            if row < len(self.temp_records): self.temp_records[row][2] = item.text()

    def delete_selected_record(self):
        row = self.table_record.currentRow()
        if row >= 0: del self.temp_records[row]; self.refresh_record_table()

    def clear_all_records(self):
        if self.temp_records and QMessageBox.question(self, '確認', '確定要清空此次暫存紀錄嗎？', QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
            self.temp_records = []; self.refresh_record_table()

    def export_data(self, f_type):
        if not self.temp_records: return
        df_e = pd.DataFrame(self.temp_records, columns=["時間", "圖號", "備註"])
        f_p, _ = QFileDialog.getSaveFileName(self, "儲存檔案", "", f"檔案 (*.{f_type})")
        if f_p:
            if f_type == 'csv': df_e.to_csv(f_p, index=False, encoding='utf-8-sig')
            else: df_e.to_excel(f_p, index=False)
            QMessageBox.information(self, "成功", "檔案匯出成功！")

    def init_admin_ui(self):
        layout = QVBoxLayout(self.tab_admin)
        self.admin_table = QTableWidget(); self.admin_table.setColumnCount(2)
        self.admin_table.setHorizontalHeaderLabels(["代碼", "名稱說明"])
        self.admin_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.admin_table)
        self.refresh_admin_table()
        edit_layout = QHBoxLayout()
        self.new_code = QLineEdit(); self.new_code.setFixedWidth(80); self.new_code.setPlaceholderText("代碼")
        self.new_name = QLineEdit(); self.new_name.setPlaceholderText("製程名稱")
        btn_add = QPushButton("新增或更新代碼"); btn_add.clicked.connect(self.admin_add)
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
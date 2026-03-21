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
        self.setWindowTitle("圖號編碼管理系統 v17.0 (圖號唯一性鎖定)")
        self.setGeometry(100, 100, 1100, 850)
        
        self.csv_file = 'process_codes.csv'
        self.temp_records = [] # 儲存格式: [時間, 圖號, 備註]
        self.load_data()
        
        self.general_font = QFont("Microsoft JhengHei", 12)
        self.setFont(self.general_font)
        
        self.tabs = QTabWidget()
        self.tab_generator = QWidget()
        self.tab_admin = QWidget()
        self.tabs.addTab(self.tab_generator, "圖號產生器")
        self.tabs.addTab(self.tab_admin, "製程代碼管理")
        
        self.main_layout = QVBoxLayout()
        self.main_layout.addWidget(self.tabs)
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.main_layout)
        self.setCentralWidget(self.central_widget)
        
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
        self.df = self.df.sort_values(by='代碼').reset_index(drop=True)

    def init_generator_ui(self):
        layout = QVBoxLayout(self.tab_generator)
        num_v = QIntValidator()
        alpha_v = QRegExpValidator(QRegExp("[a-zA-Z0-9]"))

        # --- 輸入區 ---
        input_box = QGroupBox("編碼輸入規範")
        grid = QVBoxLayout()
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("客戶碼: A"))
        self.input_cust = QLineEdit(); self.input_cust.setFixedWidth(100); self.input_cust.setMaxLength(4); self.input_cust.setValidator(num_v)
        row1.addWidget(self.input_cust); row1.addSpacing(30); row1.addWidget(QLabel("來源:"))
        self.rb_a = QRadioButton("廠內(A)"); self.rb_a.setChecked(True); self.rb_b = QRadioButton("委外(B)")
        row1.addWidget(self.rb_a); row1.addWidget(self.rb_b); row1.addStretch()
        grid.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("成品碼:  "))
        self.input_prod = QLineEdit(); self.input_prod.setFixedWidth(100); self.input_prod.setMaxLength(4); self.input_prod.setValidator(num_v)
        row2.addWidget(self.input_prod); row2.addSpacing(30); row2.addWidget(QLabel("製程:"))
        self.cb_proc = QComboBox(); self.cb_proc.setMinimumWidth(300)
        row2.addWidget(self.cb_proc); row2.addStretch()
        grid.addLayout(row2)

        row3 = QHBoxLayout()
        self.check_fg = QCheckBox("業務成品(FG)"); self.check_fg.stateChanged.connect(self.toggle_fg)
        row3.addWidget(self.check_fg); row3.addSpacing(20); row3.addWidget(QLabel("階層(L1-L3):"))
        self.input_l1 = QLineEdit(); self.input_l1.setFixedWidth(40); self.input_l1.setMaxLength(2); self.input_l1.setValidator(num_v)
        self.input_l2 = QLineEdit(); self.input_l2.setFixedWidth(40); self.input_l2.setMaxLength(2); self.input_l2.setValidator(num_v)
        self.input_l3 = QLineEdit(); self.input_l3.setFixedWidth(40); self.input_l3.setMaxLength(2); self.input_l3.setValidator(num_v)
        row3.addWidget(self.input_l1); row3.addWidget(self.input_l2); row3.addWidget(self.input_l3)
        row3.addSpacing(20); row3.addWidget(QLabel("版本:"))
        self.input_ver = QLineEdit("0"); self.input_ver.setFixedWidth(40); self.input_ver.setMaxLength(1); self.input_ver.setValidator(alpha_v)
        row3.addWidget(self.input_ver); row3.addStretch()
        grid.addLayout(row3)
        input_box.setLayout(grid)
        layout.addWidget(input_box)

        # --- 生成按鈕 (22px) ---
        self.btn_gen = QPushButton("確認生成圖號並加入清單")
        self.btn_gen.setFixedHeight(65)
        self.btn_gen.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; font-size: 22px; border-radius: 8px;")
        self.btn_gen.clicked.connect(self.generate_and_save)
        layout.addWidget(self.btn_gen)

        # --- 紀錄表格 (12px) ---
        record_box = QGroupBox("暫存紀錄管理 (已啟用圖號重複鎖定)")
        record_layout = QVBoxLayout()
        self.table_record = QTableWidget(); self.table_record.setColumnCount(3)
        self.table_record.setHorizontalHeaderLabels(["生成時間", "完整圖號結果", "自定義備註"])
        self.table_record.setFont(QFont("Microsoft JhengHei", 12))
        self.table_record.horizontalHeader().setFont(QFont("Microsoft JhengHei", 12))
        self.table_record.verticalHeader().setDefaultSectionSize(30)
        
        header = self.table_record.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed); self.table_record.setColumnWidth(0, 100)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Fixed); self.table_record.setColumnWidth(2, 200)
        self.table_record.itemChanged.connect(self.sync_note_to_list)
        record_layout.addWidget(self.table_record)
        
        action_layout = QHBoxLayout()
        btn_del = QPushButton("刪除單筆項目"); btn_clr = QPushButton("清空全部紀錄")
        btn_del.setStyleSheet("background-color: #f39c12; color: white; height: 35px;")
        btn_clr.setStyleSheet("background-color: #c0392b; color: white; height: 35px;")
        btn_del.clicked.connect(self.delete_selected_record)
        btn_clr.clicked.connect(self.clear_all_records)
        
        btn_csv = QPushButton("匯出 CSV 檔"); btn_excel = QPushButton("匯出 Excel 檔")
        btn_csv.setStyleSheet("height: 35px;"); btn_excel.setStyleSheet("height: 35px;")
        btn_csv.clicked.connect(lambda: self.export_data('csv'))
        btn_excel.clicked.connect(lambda: self.export_data('xlsx'))
        
        action_layout.addWidget(btn_del); action_layout.addWidget(btn_clr); action_layout.addStretch()
        action_layout.addWidget(btn_csv); action_layout.addWidget(btn_excel)
        record_layout.addLayout(action_layout)
        record_box.setLayout(record_layout)
        layout.addWidget(record_box)

    def generate_and_save(self):
        # 1. 組合圖號
        cust = f"A{self.input_cust.text().strip().zfill(4)}"
        src = "A" if self.rb_a.isChecked() else "B"
        prod = self.input_prod.text().strip().zfill(4)
        proc = self.cb_proc.currentText().split(" - ")[0] if self.cb_proc.currentText() else "SA"
        mid = "FG0000" if self.check_fg.isChecked() else f"{self.input_l1.text().strip().zfill(2)}{self.input_l2.text().strip().zfill(2)}{self.input_l3.text().strip().zfill(2)}"
        ver = self.input_ver.text().upper().strip() or "0"
        
        new_formatted_code = f"{cust}-{src}{prod}-{proc}{mid}-{ver}"
        
        # 2. 重複性檢核 (關鍵改動)
        # 檢查當前暫存清單 (self.temp_records) 中是否已存在該圖號
        existing_codes = [record[1] for record in self.temp_records]
        
        if new_formatted_code in existing_codes:
            QMessageBox.critical(self, "生成失敗", f"圖號重複！\n\n此圖號已存在於清單中：\n{new_formatted_code}\n\n請檢查規格或修改版本碼。")
            return
        
        # 3. 若通過檢查，加入清單
        now = datetime.now().strftime("%H:%M:%S")
        self.table_record.itemChanged.disconnect(self.sync_note_to_list)
        self.temp_records.append([now, new_formatted_code, ""])
        self.refresh_record_table()
        self.table_record.itemChanged.connect(self.sync_note_to_list)
        
        # 自動複製到剪貼簿
        QApplication.clipboard().setText(new_formatted_code)

    def refresh_record_table(self):
        self.table_record.setRowCount(len(self.temp_records))
        for i, (t, c, n) in enumerate(self.temp_records):
            it0 = QTableWidgetItem(t); it0.setFlags(it0.flags() ^ Qt.ItemIsEditable)
            it1 = QTableWidgetItem(c); it1.setFlags(it1.flags() ^ Qt.ItemIsEditable); it1.setForeground(Qt.blue)
            it2 = QTableWidgetItem(n)
            self.table_record.setItem(i, 0, it0); self.table_record.setItem(i, 1, it1); self.table_record.setItem(i, 2, it2)

    # --- 以下為其餘功能保持不變 ---
    def load_data(self):
        if not os.path.exists(self.csv_file):
            data = {'代碼': ['AS', 'PC', 'LC', 'MA', 'QC', 'SA'], '名稱': ['組裝', '粉烤', '液烤', '素材', '品檢', '半成品']}
            self.df = pd.DataFrame(data); self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')
        else: self.df = pd.read_csv(self.csv_file)
        self.df = self.df.sort_values(by='代碼').reset_index(drop=True)

    def init_admin_ui(self):
        layout = QVBoxLayout(self.tab_admin)
        admin_box = QGroupBox("現有製程代碼清單")
        admin_layout = QVBoxLayout()
        self.admin_table = QTableWidget(); self.admin_table.setColumnCount(2)
        self.admin_table.setHorizontalHeaderLabels(["英文代碼", "中文說明"])
        self.admin_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.admin_table.setSelectionBehavior(QTableWidget.SelectRows); self.admin_table.setEditTriggers(QTableWidget.NoEditTriggers)
        admin_layout.addWidget(self.admin_table); admin_box.setLayout(admin_layout); layout.addWidget(admin_box)
        edit_box = QGroupBox("新增或刪除製程代碼"); edit_layout = QVBoxLayout()
        input_row = QHBoxLayout(); self.new_code = QLineEdit(); self.new_code.setPlaceholderText("2位英文"); self.new_code.setMaxLength(2)
        self.new_code.setValidator(QRegExpValidator(QRegExp("[a-zA-Z]{2}"))); self.new_name = QLineEdit(); self.new_name.setPlaceholderText("中文說明")
        input_row.addWidget(QLabel("代碼:")); input_row.addWidget(self.new_code); input_row.addWidget(QLabel("說明:")); input_row.addWidget(self.new_name); edit_layout.addLayout(input_row)
        btn_row = QHBoxLayout(); self.btn_admin_add = QPushButton("【新增】製程代碼")
        self.btn_admin_add.setStyleSheet("background-color: #27ae60; color: white; height: 40px; font-weight: bold;")
        self.btn_admin_add.clicked.connect(self.admin_add); self.btn_admin_del = QPushButton("【刪除】選中代碼")
        self.btn_admin_del.setStyleSheet("background-color: #c0392b; color: white; height: 40px; font-weight: bold;")
        self.btn_admin_del.clicked.connect(self.admin_del); btn_row.addWidget(self.btn_admin_add); btn_row.addWidget(self.btn_admin_del); edit_layout.addLayout(btn_row); edit_box.setLayout(edit_layout); layout.addWidget(edit_box); self.refresh_admin_table()

    def refresh_admin_table(self):
        self.df = self.df.sort_values(by='代碼').reset_index(drop=True)
        self.admin_table.setRowCount(len(self.df))
        for i, row in self.df.iterrows():
            it_code = QTableWidgetItem(str(row['代碼'])); it_name = QTableWidgetItem(str(row['名稱']))
            it_code.setTextAlignment(Qt.AlignCenter); self.admin_table.setItem(i, 0, it_code); self.admin_table.setItem(i, 1, it_name)
        self.update_combobox()

    def admin_add(self):
        code = self.new_code.text().upper().strip(); name = self.new_name.text().strip()
        if len(code) != 2 or not name: return
        if code in self.df['代碼'].values or name in self.df['名稱'].values:
            QMessageBox.critical(self, "失敗", "代碼或名稱重複！"); return
        new_row = pd.DataFrame({'代碼': [code], '名稱': [name]})
        self.df = pd.concat([self.df, new_row], ignore_index=True); self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig'); self.new_code.clear(); self.new_name.clear(); self.refresh_admin_table()

    def admin_del(self):
        row = self.admin_table.currentRow()
        if row < 0: return
        target_code = self.admin_table.item(row, 0).text()
        if QMessageBox.question(self, '確認', f"確定刪除：{target_code}？", QMessageBox.Yes|QMessageBox.No) == QMessageBox.Yes:
            self.df = self.df[self.df['代碼'] != target_code]; self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig'); self.refresh_admin_table()

    def export_data(self, f_type):
        if not self.temp_records: return
        df_e = pd.DataFrame(self.temp_records, columns=["生成時間", "圖號", "備註"])
        f_p, _ = QFileDialog.getSaveFileName(self, "儲存檔案", "", f"檔案 (*.{f_type})")
        if f_p:
            if f_type == 'csv': df_e.to_csv(f_p, index=False, encoding='utf-8-sig')
            else: df_e.to_excel(f_p, index=False)
            QMessageBox.information(self, "成功", "匯出成功")

    def toggle_fg(self, state):
        enabled = state != Qt.Checked
        for e in [self.input_l1, self.input_l2, self.input_l3]: e.setEnabled(enabled)

    def update_combobox(self):
        self.cb_proc.clear(); self.cb_proc.addItems([f"{r['代碼']} - {r['名稱']}" for _, r in self.df.iterrows()])

    def sync_note_to_list(self, item):
        if item.column() == 2:
            row = item.row()
            if row < len(self.temp_records): self.temp_records[row][2] = item.text()

    def delete_selected_record(self):
        row = self.table_record.currentRow()
        if row >= 0: del self.temp_records[row]; self.refresh_record_table()

    def clear_all_records(self):
        if self.temp_records and QMessageBox.question(self, '確認', '確定清空？', QMessageBox.Yes|QMessageBox.No) == QMessageBox.Yes:
            self.temp_records = []; self.refresh_record_table()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DrawingApp()
    window.show()
    sys.exit(app.exec_())
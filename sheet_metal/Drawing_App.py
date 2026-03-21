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
        self.setWindowTitle("圖號編碼管理系統 v7.0 (大字體版)")
        self.setGeometry(100, 100, 1050, 850)
        
        self.csv_file = 'process_codes.csv'
        self.temp_records = []
        self.load_data()
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        
        # 設定全域大字體 (12級)
        self.general_font = QFont("Microsoft JhengHei", 12)
        self.setFont(self.general_font)
        
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
                    '名稱': ['組裝(Assembly)', '粉烤(Powder)', '液烤(Liquid)', '素材(Material)', '品檢(QC)', '半成品(Sub-Assy)']}
            self.df = pd.DataFrame(data)
            self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')
        else:
            self.df = pd.read_csv(self.csv_file)

    def init_generator_ui(self):
        layout = QVBoxLayout(self.tab_generator)
        layout.setSpacing(15)
        
        # --- 區塊 1: 輸入區 ---
        input_box = QGroupBox("編碼輸入")
        grid = QVBoxLayout()
        grid.setSpacing(10)
        
        # 第一排：客戶(縮短) 與 來源
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("客戶(5碼):"))
        self.input_cust = QLineEdit()
        self.input_cust.setFixedWidth(120) # 縮短一半寬度
        row1.addWidget(self.input_cust)
        
        row1.addSpacing(40) # 增加間距
        row1.addWidget(QLabel("來源:"))
        self.rb_a = QRadioButton("A 廠內"); self.rb_a.setChecked(True)
        self.rb_b = QRadioButton("B 委外")
        row1.addWidget(self.rb_a); row1.addWidget(self.rb_b)
        row1.addStretch() # 推向左側
        grid.addLayout(row1)

        # 第二排：成品碼(縮短) 與 製程
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("成品碼(4碼):"))
        self.input_prod = QLineEdit()
        self.input_prod.setFixedWidth(100) # 縮短一半寬度
        row2.addWidget(self.input_prod)
        
        row2.addSpacing(40)
        row2.addWidget(QLabel("選擇製程:"))
        self.cb_proc = QComboBox()
        self.cb_proc.setMinimumWidth(250)
        row2.addWidget(self.cb_proc)
        row2.addStretch()
        grid.addLayout(row2)

        # 第三排：階層與版本
        row3 = QHBoxLayout()
        self.check_fg = QCheckBox("業務成品 (FG0000)")
        self.check_fg.stateChanged.connect(self.toggle_fg)
        row3.addWidget(self.check_fg)
        
        row3.addSpacing(20)
        row3.addWidget(QLabel("L1-L2-L3:"))
        self.input_l1 = QLineEdit(); self.input_l1.setFixedWidth(50)
        self.input_l2 = QLineEdit(); self.input_l2.setFixedWidth(50)
        self.input_l3 = QLineEdit(); self.input_l3.setFixedWidth(50)
        row3.addWidget(self.input_l1); row3.addWidget(self.input_l2); row3.addWidget(self.input_l3)
        
        row3.addSpacing(20)
        row3.addWidget(QLabel("版本:"))
        self.input_ver = QLineEdit("0")
        self.input_ver.setFixedWidth(50)
        row3.addWidget(self.input_ver)
        row3.addStretch()
        grid.addLayout(row3)
        
        input_box.setLayout(grid)
        layout.addWidget(input_box)

        # --- 區塊 2: 生成按鈕 (加大) ---
        self.btn_gen = QPushButton("★ 生成圖號並儲存紀錄 ★")
        self.btn_gen.setStyleSheet("""
            QPushButton {
                background-color: #27ae60; 
                color: white; 
                height: 60px; 
                font-weight: bold; 
                font-size: 18px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        self.btn_gen.clicked.connect(self.generate_and_save)
        layout.addWidget(self.btn_gen)

        # 結果顯示區 (特大字體)
        self.lbl_result = QLabel("請輸入資料後點擊生成")
        self.lbl_result.setAlignment(Qt.AlignCenter)
        self.lbl_result.setStyleSheet("""
            font-size: 32px; 
            color: #c0392b; 
            font-family: 'Consolas', 'Courier New'; 
            font-weight: bold; 
            background: #ecf0f1; 
            border: 2px dashed #7f8c8d; 
            padding: 15px;
        """)
        layout.addWidget(self.lbl_result)

        # --- 區塊 3: 紀錄清單 ---
        record_box = QGroupBox("暫存紀錄管理")
        record_layout = QVBoxLayout()
        
        self.table_record = QTableWidget()
        self.table_record.setColumnCount(3)
        self.table_record.setHorizontalHeaderLabels(["時間", "完整圖號 (雙擊可複製)", "備註"])
        self.table_record.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_record.setStyleSheet("font-size:18px;") # 表格內容稍微小一點點以容納長圖號
        record_layout.addWidget(self.table_record)
        
        action_layout = QHBoxLayout()
        btn_del_selected = QPushButton("刪除選中紀錄")
        btn_del_selected.setStyleSheet("background-color: #f39c12; color: white; height: 40px;")
        btn_del_selected.clicked.connect(self.delete_selected_record)
        
        btn_clear_all = QPushButton("取消/清空所有")
        btn_clear_all.setStyleSheet("background-color: #c0392b; color: white; height: 40px;")
        btn_clear_all.clicked.connect(self.clear_all_records)
        
        action_layout.addWidget(btn_del_selected)
        action_layout.addWidget(btn_clear_all)
        action_layout.addStretch()
        
        btn_exp_csv = QPushButton("匯出 CSV")
        btn_exp_excel = QPushButton("匯出 Excel")
        for btn in [btn_exp_csv, btn_exp_excel]: btn.setFixedWidth(120); btn.setFixedHeight(35)
        btn_exp_csv.clicked.connect(lambda: self.export_data('csv'))
        btn_exp_excel.clicked.connect(lambda: self.export_data('xlsx'))
        
        action_layout.addWidget(btn_exp_csv)
        action_layout.addWidget(btn_exp_excel)
        record_layout.addLayout(action_layout)
        record_box.setLayout(record_layout)
        layout.addWidget(record_box)

    def toggle_fg(self, state):
        enabled = state != Qt.Checked
        for e in [self.input_l1, self.input_l2, self.input_l3]: e.setEnabled(enabled)

    def update_combobox(self):
        self.cb_proc.clear()
        options = [f"{row['代碼']} - {row['名稱']}" for _, row in self.df.iterrows()]
        self.cb_proc.addItems(options)

    def generate_and_save(self):
        cust = self.input_cust.text().upper().strip().zfill(5)[:5]
        src = "A" if self.rb_a.isChecked() else "B"
        prod = self.input_prod.text().strip().zfill(4)[:4]
        proc = self.cb_proc.currentText().split(" - ")[0] if self.cb_proc.currentText() else "SA"
        
        if self.check_fg.isChecked():
            mid = "FG0000"
        else:
            l1 = self.input_l1.text().strip().zfill(2)[:2]; l2 = self.input_l2.text().strip().zfill(2)[:2]; l3 = self.input_l3.text().strip().zfill(2)[:2]
            mid = f"{l1}{l2}{l3}"
        
        ver = self.input_ver.text().upper().strip()[:1] or "0"
        formatted = f"{cust}-{src}{prod}-{proc}{mid}-{ver}"
        
        self.lbl_result.setText(formatted)
        now = datetime.now().strftime("%H:%M:%S")
        self.temp_records.append([now, formatted, ""])
        self.refresh_record_table()
        QApplication.clipboard().setText(formatted)

    def refresh_record_table(self):
        self.table_record.setRowCount(len(self.temp_records))
        for i, (time, code, note) in enumerate(self.temp_records):
            self.table_record.setItem(i, 0, QTableWidgetItem(time))
            self.table_record.setItem(i, 1, QTableWidgetItem(code))
            self.table_record.setItem(i, 2, QTableWidgetItem(note))

    def delete_selected_record(self):
        row = self.table_record.currentRow()
        if row >= 0:
            del self.temp_records[row]
            self.refresh_record_table()

    def clear_all_records(self):
        if not self.temp_records: return
        reply = QMessageBox.question(self, '確認', '確定要清空此次暫存紀錄嗎？', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.temp_records = []; self.refresh_record_table()

    def export_data(self, f_type):
        if not self.temp_records: return
        df_e = pd.DataFrame(self.temp_records, columns=["時間", "圖號", "備註"])
        f_p, _ = QFileDialog.getSaveFileName(self, "儲存檔案", "", f"Files (*.{f_type})")
        if f_p:
            if f_type == 'csv': df_e.to_csv(f_p, index=False, encoding='utf-8-sig')
            else: df_e.to_excel(f_p, index=False)
            QMessageBox.information(self, "成功", "匯出完成")

    # --- 管理介面 (也同步加大字體) ---
    def init_admin_ui(self):
        layout = QVBoxLayout(self.tab_admin)
        self.admin_table = QTableWidget(); self.admin_table.setColumnCount(2)
        self.admin_table.setHorizontalHeaderLabels(["代碼", "名稱"])
        self.admin_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.admin_table)
        self.refresh_admin_table()
        
        edit_layout = QHBoxLayout()
        self.new_code = QLineEdit(); self.new_code.setPlaceholderText("代碼"); self.new_code.setFixedWidth(80)
        self.new_name = QLineEdit(); self.new_name.setPlaceholderText("名稱說明")
        btn_add = QPushButton("新增/更新"); btn_add.clicked.connect(self.admin_add)
        btn_del = QPushButton("刪除代碼"); btn_del.clicked.connect(self.admin_del)
        edit_layout.addWidget(self.new_code); edit_layout.addWidget(self.new_name)
        edit_layout.addWidget(btn_add); edit_layout.addWidget(btn_del)
        layout.addLayout(edit_layout)

    def refresh_admin_table(self):
        self.df = self.df.sort_values(by='代碼')
        self.admin_table.setRowCount(len(self.df))
        for i, row in self.df.iterrows():
            self.admin_table.setItem(i, 0, QTableWidgetItem(str(row['代碼'])))
            self.admin_table.setItem(i, 1, QTableWidgetItem(str(row['名稱'])))
        self.update_combobox()

    def admin_add(self):
        c, n = self.new_code.text().upper()[:2], self.new_name.text()
        if not c or not n: return
        self.df = self.df[self.df['代碼'] != c]
        self.df = pd.concat([self.df, pd.DataFrame({'代碼': [c], '名稱': [n]})], ignore_index=True)
        self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')
        self.refresh_admin_table()

    def admin_del(self):
        r = self.admin_table.currentRow()
        if r >= 0:
            c = self.admin_table.item(r, 0).text()
            self.df = self.df[self.df['代碼'] != c]; self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig'); self.refresh_admin_table()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DrawingApp()
    window.show()
    sys.exit(app.exec_())
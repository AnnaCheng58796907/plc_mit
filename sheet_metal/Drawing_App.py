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
        self.setWindowTitle("圖號編碼管理系統 v5.0 (含暫存與多格式匯出)")
        self.setGeometry(100, 100, 1000, 800)
        
        self.csv_file = 'process_codes.csv'
        self.temp_records = [] # 用於存放本次作業的編碼紀錄
        self.load_data()
        
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
                    '名稱': ['組裝(Assembly)', '粉烤(Powder)', '液烤(Liquid)', '素材(Material)', '品檢(QC)', '半成品(Sub-Assy)']}
            self.df = pd.DataFrame(data)
            self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')
        else:
            self.df = pd.read_csv(self.csv_file)

    # --- 介面 1: 圖號產生器 ---
    def init_generator_ui(self):
        layout = QVBoxLayout(self.tab_generator)
        
        # 上半部：輸入區
        input_box = QGroupBox("編碼輸入")
        grid = QVBoxLayout()
        
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("客戶(5碼):"))
        self.input_cust = QLineEdit()
        row1.addWidget(self.input_cust)
        row1.addWidget(QLabel("來源:"))
        self.rb_a = QRadioButton("A 廠內"); self.rb_a.setChecked(True)
        self.rb_b = QRadioButton("B 委外")
        row1.addWidget(self.rb_a); row1.addWidget(self.rb_b)
        grid.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("成品碼(4碼):"))
        self.input_prod = QLineEdit()
        row2.addWidget(self.input_prod)
        row2.addWidget(QLabel("製程:"))
        self.cb_proc = QComboBox()
        row2.addWidget(self.cb_proc)
        grid.addLayout(row2)

        row3 = QHBoxLayout()
        self.check_fg = QCheckBox("業務成品 (FG0000)")
        self.check_fg.stateChanged.connect(self.toggle_fg)
        row3.addWidget(self.check_fg)
        row3.addWidget(QLabel("L1-L2-L3:"))
        self.input_l1 = QLineEdit(); self.input_l1.setFixedWidth(40)
        self.input_l2 = QLineEdit(); self.input_l2.setFixedWidth(40)
        self.input_l3 = QLineEdit(); self.input_l3.setFixedWidth(40)
        row3.addWidget(self.input_l1); row3.addWidget(self.input_l2); row3.addWidget(self.input_l3)
        row3.addWidget(QLabel("版本:")); self.input_ver = QLineEdit("0"); self.input_ver.setFixedWidth(40)
        row3.addWidget(self.input_ver)
        grid.addLayout(row3)
        input_box.setLayout(grid)
        layout.addWidget(input_box)

        # 中間：功能按鈕
        btn_layout = QHBoxLayout()
        self.btn_gen = QPushButton("生成並儲存至紀錄")
        self.btn_gen.setStyleSheet("background-color: #27ae60; color: white; height: 40px; font-weight: bold;")
        self.btn_gen.clicked.connect(self.generate_and_save)
        btn_layout.addWidget(self.btn_gen)
        layout.addLayout(btn_layout)

        # 下半部：紀錄顯示區
        record_box = QGroupBox("本次編碼紀錄 (選取列後可按右鍵或下方按鈕刪除)")
        record_layout = QVBoxLayout()
        
        self.table_record = QTableWidget()
        self.table_record.setColumnCount(3)
        self.table_record.setHorizontalHeaderLabels(["生成時間", "完整圖號", "備註"])
        self.table_record.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        record_layout.addWidget(self.table_record)
        
        # 匯出與刪除按鈕
        export_layout = QHBoxLayout()
        btn_del_rec = QPushButton("刪除選中紀錄")
        btn_del_rec.clicked.connect(self.delete_record)
        btn_exp_csv = QPushButton("轉出 CSV")
        btn_exp_csv.clicked.connect(lambda: self.export_data('csv'))
        btn_exp_excel = QPushButton("轉出 EXCEL")
        btn_exp_excel.clicked.connect(lambda: self.export_data('xlsx'))
        
        export_layout.addWidget(btn_del_rec)
        export_layout.addStretch()
        export_layout.addWidget(btn_exp_csv)
        export_layout.addWidget(btn_exp_excel)
        record_layout.addLayout(export_layout)
        
        record_box.setLayout(record_layout)
        layout.addWidget(record_box)

    def toggle_fg(self, state):
        enabled = state != Qt.Checked
        self.input_l1.setEnabled(enabled); self.input_l2.setEnabled(enabled); self.input_l3.setEnabled(enabled)

    def update_combobox(self):
        self.cb_proc.clear()
        options = [f"{row['代碼']} - {row['名稱']}" for _, row in self.df.iterrows()]
        self.cb_proc.addItems(options)

    # --- 核心邏輯：生成與存檔 ---
    def generate_and_save(self):
        cust = self.input_cust.text().upper().strip().zfill(5)[:5]
        src = "A" if self.rb_a.isChecked() else "B"
        prod = self.input_prod.text().strip().zfill(4)[:4]
        proc = self.cb_proc.currentText().split(" - ")[0] if self.cb_proc.currentText() else "SA"
        
        if self.check_fg.isChecked():
            mid = "FG0000"
        else:
            l1 = self.input_l1.text().strip().zfill(2)[:2]
            l2 = self.input_l2.text().strip().zfill(2)[:2]
            l3 = self.input_l3.text().strip().zfill(2)[:2]
            mid = f"{l1}{l2}{l3}"
        
        ver = self.input_ver.text().upper().strip()[:1] or "0"
        formatted = f"{cust}-{src}{prod}-{proc}{mid}-{ver}"
        
        # 存入記憶體紀錄
        now = datetime.now().strftime("%H:%M:%S")
        self.temp_records.append([now, formatted, ""])
        self.refresh_record_table()
        
        # 自動複製最後一筆
        QApplication.clipboard().setText(formatted)

    def refresh_record_table(self):
        self.table_record.setRowCount(len(self.temp_records))
        for i, (time, code, note) in enumerate(self.temp_records):
            self.table_record.setItem(i, 0, QTableWidgetItem(time))
            self.table_record.setItem(i, 1, QTableWidgetItem(code))
            self.table_record.setItem(i, 2, QTableWidgetItem(note))

    def delete_record(self):
        current_row = self.table_record.currentRow()
        if current_row >= 0:
            del self.temp_records[current_row]
            self.refresh_record_table()

    def export_data(self, file_type):
        if not self.temp_records:
            QMessageBox.warning(self, "警告", "目前沒有任何編碼紀錄可匯出！")
            return
        
        df_export = pd.DataFrame(self.temp_records, columns=["時間", "圖號", "備註"])
        
        file_path, _ = QFileDialog.getSaveFileName(self, "儲存檔案", "", f"Files (*.{file_type})")
        if file_path:
            try:
                if file_type == 'csv':
                    df_export.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    df_export.to_excel(file_path, index=False)
                QMessageBox.information(self, "成功", f"檔案已儲存至：\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"匯出失敗：{str(e)}")

    # --- 介面 2: 管理區 (略，同前版本) ---
    def init_admin_ui(self):
        layout = QVBoxLayout(self.tab_admin)
        self.admin_table = QTableWidget(); self.admin_table.setColumnCount(2)
        self.admin_table.setHorizontalHeaderLabels(["代碼", "名稱"])
        layout.addWidget(self.admin_table)
        self.refresh_admin_table()
        
        edit_layout = QHBoxLayout()
        self.new_code = QLineEdit(); self.new_name = QLineEdit()
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
        code, name = self.new_code.text().upper()[:2], self.new_name.text()
        if not code or not name: return
        self.df = self.df[self.df['代碼'] != code]
        self.df = pd.concat([self.df, pd.DataFrame({'代碼': [code], '名稱': [name]})], ignore_index=True)
        self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')
        self.refresh_admin_table()

    def admin_del(self):
        row = self.admin_table.currentRow()
        if row >= 0:
            code = self.admin_table.item(row, 0).text()
            self.df = self.df[self.df['代碼'] != code]
            self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')
            self.refresh_admin_table()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft JhengHei", 10))
    window = DrawingApp()
    window.show()
    sys.exit(app.exec_())
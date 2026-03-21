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
        self.setWindowTitle("圖號編碼管理系統 v6.0 (含全數清空功能)")
        self.setGeometry(100, 100, 1000, 800)
        
        self.csv_file = 'process_codes.csv'
        self.temp_records = [] # 存放本次作業的編碼紀錄
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

    def init_generator_ui(self):
        layout = QVBoxLayout(self.tab_generator)
        
        # --- 區塊 1: 輸入區 ---
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

        # --- 區塊 2: 生成按鈕 ---
        self.btn_gen = QPushButton("生成並儲存至紀錄")
        self.btn_gen.setStyleSheet("background-color: #27ae60; color: white; height: 45px; font-weight: bold; font-size: 14px;")
        self.btn_gen.clicked.connect(self.generate_and_save)
        layout.addWidget(self.btn_gen)

        # --- 區塊 3: 紀錄顯示與功能區 ---
        record_box = QGroupBox("本次編碼紀錄清單")
        record_layout = QVBoxLayout()
        
        self.table_record = QTableWidget()
        self.table_record.setColumnCount(3)
        self.table_record.setHorizontalHeaderLabels(["生成時間", "完整圖號", "備註"])
        self.table_record.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        record_layout.addWidget(self.table_record)
        
        # 功能按鈕列
        action_layout = QHBoxLayout()
        
        # 刪除與清空 (左側)
        btn_del_selected = QPushButton("刪除選中項目")
        btn_del_selected.setStyleSheet("background-color: #f39c12; color: white;")
        btn_del_selected.clicked.connect(self.delete_selected_record)
        
        btn_clear_all = QPushButton("清空所有紀錄")
        btn_clear_all.setStyleSheet("background-color: #c0392b; color: white;")
        btn_clear_all.clicked.connect(self.clear_all_records)
        
        action_layout.addWidget(btn_del_selected)
        action_layout.addWidget(btn_clear_all)
        action_layout.addStretch() # 彈性空間
        
        # 匯出按鈕 (右側)
        btn_exp_csv = QPushButton("匯出 CSV")
        btn_exp_csv.setFixedHeight(35)
        btn_exp_csv.clicked.connect(lambda: self.export_data('csv'))
        
        btn_exp_excel = QPushButton("匯出 Excel")
        btn_exp_excel.setFixedHeight(35)
        btn_exp_excel.clicked.connect(lambda: self.export_data('xlsx'))
        
        action_layout.addWidget(btn_exp_csv)
        action_layout.addWidget(btn_exp_excel)
        
        record_layout.addLayout(action_layout)
        record_box.setLayout(record_layout)
        layout.addWidget(record_box)

    def toggle_fg(self, state):
        enabled = state != Qt.Checked
        self.input_l1.setEnabled(enabled); self.input_l2.setEnabled(enabled); self.input_l3.setEnabled(enabled)

    def update_combobox(self):
        self.cb_proc.clear()
        options = [f"{row['代碼']} - {row['名稱']}" for _, row in self.df.iterrows()]
        self.cb_proc.addItems(options)

    def generate_and_save(self):
        # 欄位補足邏輯
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
        
        # 暫存
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
        current_row = self.table_record.currentRow()
        if current_row >= 0:
            del self.temp_records[current_row]
            self.refresh_record_table()
        else:
            QMessageBox.information(self, "提示", "請先點選表格中的一列再進行刪除。")

    def clear_all_records(self):
        """取消此次所有編碼暫存"""
        if not self.temp_records:
            return
            
        reply = QMessageBox.question(self, '確認清空', '確定要取消並清空「此次」所有的暫存編碼紀錄嗎？',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.temp_records = []
            self.refresh_record_table()
            QMessageBox.information(self, "完成", "所有暫存紀錄已清空。")

    def export_data(self, file_type):
        if not self.temp_records:
            QMessageBox.warning(self, "警告", "目前沒有紀錄可以匯出！")
            return
        
        df_export = pd.DataFrame(self.temp_records, columns=["時間", "圖號", "備註"])
        file_path, _ = QFileDialog.getSaveFileName(self, "儲存檔案", "", f"Files (*.{file_type})")
        
        if file_path:
            try:
                if file_type == 'csv':
                    df_export.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    df_export.to_excel(file_path, index=False)
                QMessageBox.information(self, "成功", f"檔案匯出完成！")
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"匯出失敗：{str(e)}")

    # --- 製程管理頁面 (維持不變) ---
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
        new_row = pd.DataFrame({'代碼': [code], '名稱': [name]})
        self.df = pd.concat([self.df, new_row], ignore_index=True)
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
import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QRadioButton, 
                             QButtonGroup, QCheckBox, QTabWidget, QTableWidget, 
                             QTableWidgetItem, QMessageBox, QGroupBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class DrawingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("圖號編碼管理系統 v4.0 (PyQt5)")
        self.setGeometry(100, 100, 900, 650)
        
        self.csv_file = 'process_codes.csv'
        self.load_data()
        
        # 主元件
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        
        # 分頁系統
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
            self.save_to_csv()
        else:
            self.df = pd.read_csv(self.csv_file)

    def save_to_csv(self):
        self.df.to_csv(self.csv_file, index=False, encoding='utf-8-sig')

    # --- 介面 1: 圖號產生器 ---
    def init_generator_ui(self):
        layout = QVBoxLayout(self.tab_generator)
        
        # 基礎資訊群組
        base_group = QGroupBox("基礎資訊")
        base_grid = QVBoxLayout()
        
        # 第一排：客戶與來源
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("客戶(5碼):"))
        self.input_cust = QLineEdit(); self.input_cust.setPlaceholderText("例如: A0001")
        row1.addWidget(self.input_cust)
        
        row1.addWidget(QLabel("來源:"))
        self.btn_group_src = QButtonGroup(self)
        self.rb_a = QRadioButton("A 廠內"); self.rb_a.setChecked(True)
        self.rb_b = QRadioButton("B 委外")
        self.btn_group_src.addButton(self.rb_a); self.btn_group_src.addButton(self.rb_b)
        row1.addWidget(self.rb_a); row1.addWidget(self.rb_b)
        base_grid.addLayout(row1)

        # 第二排：成品碼與製程
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("成品碼(4碼):"))
        self.input_prod = QLineEdit(); self.input_prod.setPlaceholderText("例如: 8888")
        row2.addWidget(self.input_prod)
        
        row2.addWidget(QLabel("選擇製程:"))
        self.cb_proc = QComboBox()
        row2.addWidget(self.cb_proc)
        base_grid.addLayout(row2)
        base_group.setLayout(base_grid)
        layout.addWidget(base_group)

        # 階層區
        lvl_group = QGroupBox("階層流水號 (L1-L2-L3)")
        lvl_layout = QVBoxLayout()
        self.check_fg = QCheckBox("業務成品 (13-18碼固定為 FG0000)")
        self.check_fg.stateChanged.connect(self.toggle_fg)
        lvl_layout.addWidget(self.check_fg)
        
        row3 = QHBoxLayout()
        self.input_l1 = QLineEdit(); self.input_l1.setPlaceholderText("L1")
        self.input_l2 = QLineEdit(); self.input_l2.setPlaceholderText("L2")
        self.input_l3 = QLineEdit(); self.input_l3.setPlaceholderText("L3")
        row3.addWidget(QLabel("L1:")); row3.addWidget(self.input_l1)
        row3.addWidget(QLabel("L2:")); row3.addWidget(self.input_l2)
        row3.addWidget(QLabel("L3:")); row3.addWidget(self.input_l3)
        row3.addWidget(QLabel("版本:")); self.input_ver = QLineEdit("0")
        row3.addWidget(self.input_ver)
        lvl_layout.addLayout(row3)
        lvl_group.setLayout(lvl_layout)
        layout.addWidget(lvl_group)

        # 按鈕與結果
        self.btn_gen = QPushButton("生成圖號並自動複製")
        self.btn_gen.setFixedHeight(50)
        self.btn_gen.setStyleSheet("background-color: #3498db; color: white; font-weight: bold;")
        self.btn_gen.clicked.connect(self.generate_code)
        layout.addWidget(self.btn_gen)

        self.lbl_result = QLabel("等待輸入...")
        self.lbl_result.setAlignment(Qt.AlignCenter)
        self.lbl_result.setStyleSheet("font-size: 28px; color: #e74c3c; font-family: Consolas; font-weight: bold; background: white; border: 1px solid #bdc3c7; padding: 10px;")
        layout.addWidget(self.lbl_result)

    def toggle_fg(self, state):
        enabled = state != Qt.Checked
        self.input_l1.setEnabled(enabled)
        self.input_l2.setEnabled(enabled)
        self.input_l3.setEnabled(enabled)

    def update_combobox(self):
        self.cb_proc.clear()
        options = [f"{row['代碼']} - {row['名稱']}" for _, row in self.df.iterrows()]
        self.cb_proc.addItems(options)

    def generate_code(self):
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
        
        # 格式化輸出：A0000-A0000-FG000000-0
        formatted = f"{cust}-{src}{prod}-{proc}{mid}-{ver}"
        self.lbl_result.setText(formatted)
        
        # 複製到剪貼簿
        clipboard = QApplication.clipboard()
        clipboard.setText(formatted)

    # --- 介面 2: 管理區 ---
    def init_admin_ui(self):
        layout = QVBoxLayout(self.tab_admin)
        
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["代碼", "製程名稱"])
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)
        self.refresh_table()

        edit_layout = QHBoxLayout()
        self.new_code = QLineEdit(); self.new_code.setPlaceholderText("代碼(2碼)")
        self.new_name = QLineEdit(); self.new_name.setPlaceholderText("製程名稱")
        btn_add = QPushButton("新增/更新")
        btn_del = QPushButton("刪除選中")
        
        btn_add.clicked.connect(self.admin_add)
        btn_del.clicked.connect(self.admin_del)
        
        edit_layout.addWidget(self.new_code)
        edit_layout.addWidget(self.new_name)
        edit_layout.addWidget(btn_add)
        edit_layout.addWidget(btn_del)
        layout.addLayout(edit_layout)

    def refresh_table(self):
        self.df = self.df.sort_values(by='代碼')
        self.table.setRowCount(len(self.df))
        for i, row in self.df.iterrows():
            self.table.setItem(i, 0, QTableWidgetItem(str(row['代碼'])))
            self.table.setItem(i, 1, QTableWidgetItem(str(row['名稱'])))
        self.update_combobox()

    def admin_add(self):
        code = self.new_code.text().upper().strip()[:2]
        name = self.new_name.text().strip()
        if not code or not name: return
        self.df = self.df[self.df['代碼'] != code]
        new_row = pd.DataFrame({'代碼': [code], '名稱': [name]})
        self.df = pd.concat([self.df, new_row], ignore_index=True)
        self.save_to_csv()
        self.refresh_table()
        self.new_code.clear(); self.new_name.clear()

    def admin_del(self):
        current_row = self.table.currentRow()
        if current_row < 0: return
        code = self.table.item(current_row, 0).text()
        self.df = self.df[self.df['代碼'] != code]
        self.save_to_csv()
        self.refresh_table()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 設定字體
    font = QFont("Microsoft JhengHei", 10)
    app.setFont(font)
    window = DrawingApp()
    window.show()
    sys.exit(app.exec_())
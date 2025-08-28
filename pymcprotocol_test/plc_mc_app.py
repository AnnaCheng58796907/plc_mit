import sys
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QPlainTextEdit,
                             QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
                             QMenuBar, QAction, QGroupBox, QGridLayout, QCheckBox)
from PyQt5.QtCore import QThread, pyqtSignal

# 使用 pymcprotocol
from pymcprotocol import Type3E

# ----------------------------------------------------
# 背景工作執行緒，負責連線和讀取PLC資料
# ----------------------------------------------------
class PlcReaderThread(QThread):
    data_ready = pyqtSignal(dict)
    
    def __init__(self, ip, port, m_config, d_config, parent=None):
        super(PlcReaderThread, self).__init__(parent)
        self.ip = ip
        self.port = port
        self.m_config = m_config
        self.d_config = d_config

    def run(self):
        try:
            # 建立 pymcprotocol 類別物件
            client = Type3E()
            client.connect(self.ip, self.port)

            self.data_ready.emit({"status": "success", "message": "連線成功，開始讀取資料。"})

            m_values, d_values = None, None

            # 根據m_config設定決定是否讀取M值
            if self.m_config['read']:
                # 在 0.3.0 版本中，M值和D值都使用 batchread_word 讀取
                # 程式庫會根據 headdevice 判斷設備類型
                m_values = client.batchread_wordunits(headdevice=f"M{self.m_config['start']}", readsize=self.m_config['count'])
            
            # 根據d_config設定決定是否讀取D值
            if self.d_config['read']:
                d_values = client.batchread_wordunits(headdevice=f"D{self.d_config['start']}", readsize=self.d_config['count'])
            
            client.close()

            self.data_ready.emit({"status": "data", "m_values": m_values, "d_values": d_values})

        except Exception as e:
            self.data_ready.emit({"status": "error", "message": f"發生意外錯誤: {e}"})

# ----------------------------------------------------
# 主要的應用程式視窗 (GUI)
# ----------------------------------------------------
class PlcDataViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("三菱FX3U資料讀取器 (by pymcprotocol v0.3.0)")
        
        self.data_log = []
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        self.create_menu_bar()

        conn_group = QGroupBox("通訊與讀取設定")
        conn_layout = QGridLayout()
        
        conn_layout.addWidget(QLabel("IP 位址:"), 0, 0)
        self.ip_input = QLineEdit("192.168.1.100")
        conn_layout.addWidget(self.ip_input, 0, 1)
        
        conn_layout.addWidget(QLabel("埠號:"), 0, 2)
        self.port_input = QLineEdit("5007")
        conn_layout.addWidget(self.port_input, 0, 3)

        self.m_checkbox = QCheckBox("讀取 M 值")
        self.m_checkbox.setChecked(True)
        conn_layout.addWidget(self.m_checkbox, 1, 0)
        conn_layout.addWidget(QLabel("起始位址:"), 1, 1)
        self.m_start_input = QLineEdit("0")
        conn_layout.addWidget(self.m_start_input, 1, 2)
        conn_layout.addWidget(QLabel("數量:"), 1, 3)
        self.m_count_input = QLineEdit("10")
        conn_layout.addWidget(self.m_count_input, 1, 4)

        self.d_checkbox = QCheckBox("讀取 D 值")
        self.d_checkbox.setChecked(True)
        conn_layout.addWidget(self.d_checkbox, 2, 0)
        conn_layout.addWidget(QLabel("起始位址:"), 2, 1)
        self.d_start_input = QLineEdit("1000")
        conn_layout.addWidget(self.d_start_input, 2, 2)
        conn_layout.addWidget(QLabel("數量:"), 2, 3)
        self.d_count_input = QLineEdit("10")
        conn_layout.addWidget(self.d_count_input, 2, 4)
        
        conn_group.setLayout(conn_layout)
        main_layout.addWidget(conn_group)

        self.connect_btn = QPushButton("讀取資料")
        self.connect_btn.clicked.connect(self.start_reading)
        main_layout.addWidget(self.connect_btn)

        self.data_table = QTableWidget()
        self.data_table.setColumnCount(3)
        self.data_table.setHorizontalHeaderLabels(["時間戳記", "M值", "D值"])
        header = self.data_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.data_table)

        self.status_box = QPlainTextEdit()
        self.status_box.setReadOnly(True)
        self.status_box.setPlaceholderText("連線狀態和讀取訊息會顯示在這裡...")
        self.status_box.setMaximumHeight(100)
        main_layout.addWidget(self.status_box)

    def create_menu_bar(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("檔案(&F)")
        
        save_action = QAction("儲存為CSV", self)
        save_action.setShortcut("Ctrl+S")
        save_action.setStatusTip("將讀取到的資料儲存為CSV檔案")
        save_action.triggered.connect(self.save_to_csv)
        file_menu.addAction(save_action)

    def start_reading(self):
        ip = self.ip_input.text()
        port = int(self.port_input.text())
        
        if not self.m_checkbox.isChecked() and not self.d_checkbox.isChecked():
            QMessageBox.warning(self, "警告", "請至少選擇讀取M值或D值其中一項。")
            return

        m_config = {
            'read': self.m_checkbox.isChecked(),
            'start': int(self.m_start_input.text()),
            'count': int(self.m_count_input.text())
        }

        d_config = {
            'read': self.d_checkbox.isChecked(),
            'start': int(self.d_start_input.text()),
            'count': int(self.d_count_input.text())
        }

        self.thread = PlcReaderThread(ip, port, m_config, d_config)
        self.thread.data_ready.connect(self.update_data)
        self.thread.start()
        
        self.log_message("開始連線並讀取資料...")
        self.connect_btn.setEnabled(False)

    def update_data(self, data):
        self.connect_btn.setEnabled(True)
        status = data.get("status")
        
        if status == "data":
            m_values = data.get("m_values")
            d_values = data.get("d_values")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            m_str = "未讀取" if m_values is None else f"M{self.m_start_input.text()}-M{int(self.m_start_input.text()) + len(m_values) - 1}: {m_values}"
            d_str = "未讀取" if d_values is None else f"D{self.d_start_input.text()}-D{int(self.d_start_input.text()) + len(d_values) - 1}: {d_values}"

            self.log_message("成功讀取到新資料。")
            self.add_row_to_table(timestamp, m_str, d_str)

            self.data_log.append({
                "timestamp": timestamp,
                "m_values": m_values,
                "d_values": d_values
            })

        elif status == "success":
            self.log_message(data.get("message"))
        elif status == "error":
            self.log_message(f"錯誤：{data.get('message')}")
            QMessageBox.critical(self, "連線錯誤", data.get("message"))
            
    def add_row_to_table(self, timestamp, m_str, d_str):
        row_count = self.data_table.rowCount()
        self.data_table.insertRow(row_count)
        self.data_table.setItem(row_count, 0, QTableWidgetItem(timestamp))
        self.data_table.setItem(row_count, 1, QTableWidgetItem(m_str))
        self.data_table.setItem(row_count, 2, QTableWidgetItem(d_str))
        self.data_table.resizeRowsToContents()
        self.data_table.scrollToBottom()

    def save_to_csv(self):
        if not self.data_log:
            QMessageBox.warning(self, "警告", "沒有資料可以儲存。")
            return
            
        try:
            df = pd.DataFrame(self.data_log)
            
            df['m_values'] = df['m_values'].apply(lambda x: str(x) if x is not None else "未讀取")
            df['d_values'] = df['d_values'].apply(lambda x: str(x) if x is not None else "未讀取")

            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"plc_data_{timestamp}.csv"
            
            df.to_csv(filename, index=False, encoding='utf-8-sig')
            
            self.log_message(f"資料已成功儲存至：{filename}")
            QMessageBox.information(self, "儲存成功", f"資料已成功儲存為 {filename}")
            
        except Exception as e:
            self.log_message(f"儲存CSV檔案時發生錯誤：{e}")
            QMessageBox.critical(self, "儲存失敗", f"儲存CSV檔案時發生錯誤：{e}")

    def log_message(self, message):
        self.status_box.appendPlainText(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = PlcDataViewer()
    viewer.show()
    sys.exit(app.exec_())
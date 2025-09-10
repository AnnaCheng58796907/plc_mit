import sys
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QPlainTextEdit,
                             QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
                             QMenuBar, QAction, QGroupBox, QGridLayout, QCheckBox, QComboBox)
from PyQt5.QtCore import QThread, pyqtSignal
import minimalmodbus
import serial
import serial.tools.list_ports

# ----------------------------------------------------
# 背景工作執行緒，負責連線和讀取PLC資料
# ----------------------------------------------------
class PlcReaderThread(QThread):
    data_ready = pyqtSignal(dict)
    
    def __init__(self, port_name, slave_id, baudrate, parity, timeout, m_config, d_config, parent=None):
        super(PlcReaderThread, self).__init__(parent)
        self.port_name = port_name
        self.slave_id = slave_id
        self.baudrate = baudrate
        self.parity = parity
        self.timeout = timeout
        self.m_config = m_config
        self.d_config = d_config

    def run(self):
        try:
            # 建立 Modbus RTU 儀器物件
            client = minimalmodbus.Instrument(self.port_name, self.slave_id)
            client.serial.baudrate = self.baudrate
            client.serial.bytesize = 8
            client.serial.stopbits = 1
            client.serial.parity = self.parity
            client.serial.timeout = self.timeout
            client.mode = minimalmodbus.MODE_RTU
            
            self.data_ready.emit({"status": "success", "message": "連線成功，開始讀取資料。"})

            m_values, d_values = None, None

            # 讀取M值 (Coils)
            if self.m_config['read']:
                m_values = client.read_bits(self.m_config['start'], self.m_config['count'])
            
            # 讀取D值 (Holding Registers)
            if self.d_config['read']:
                d_values = client.read_registers(self.d_config['start'], self.d_config['count'], functioncode=3)
            
            client.serial.close()

            self.data_ready.emit({"status": "data", "m_values": m_values, "d_values": d_values})

        except IOError as e:
            self.data_ready.emit({"status": "error", "message": f"I/O 錯誤，請檢查通訊設定：{e}"})
        except Exception as e:
            self.data_ready.emit({"status": "error", "message": f"發生意外錯誤：{e}"})

# ----------------------------------------------------
# 主要的應用程式視窗 (GUI)
# ----------------------------------------------------
class PlcDataViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("三菱Modbus 485資料讀取器")
        self.resize(800, 600)  # 設定初始視窗大小
        self.data_log = []
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        self.create_menu_bar()

        conn_group = QGroupBox("通訊與讀取設定")
        conn_layout = QGridLayout()
        
        # COM 埠選擇
        conn_layout.addWidget(QLabel("COM 埠:"), 0, 0)
        self.com_port_combo = QComboBox()
        self.refresh_com_ports()
        conn_layout.addWidget(self.com_port_combo, 0, 1)
        refresh_btn = QPushButton("刷新")
        refresh_btn.clicked.connect(self.refresh_com_ports)
        conn_layout.addWidget(refresh_btn, 0, 2)

        # 其他通訊參數
        conn_layout.addWidget(QLabel("站號:"), 1, 0)
        self.slave_id_input = QLineEdit("1")
        conn_layout.addWidget(self.slave_id_input, 1, 1)
        
        conn_layout.addWidget(QLabel("波特率:"), 1, 2)
        self.baudrate_combo = QComboBox()
        self.baudrate_combo.addItems(["9600", "19200", "38400"])
        self.baudrate_combo.setCurrentText("9600")
        conn_layout.addWidget(self.baudrate_combo, 1, 3)
        
        conn_layout.addWidget(QLabel("同位元:"), 1, 4)
        self.parity_combo = QComboBox()
        self.parity_combo.addItems(["EVEN", "ODD", "NONE"])
        self.parity_combo.setCurrentText("EVEN")
        conn_layout.addWidget(self.parity_combo, 1, 5)

        # 讀取 M 值設定
        self.m_checkbox = QCheckBox("讀取 M 值")
        self.m_checkbox.setChecked(True)
        conn_layout.addWidget(self.m_checkbox, 2, 0)
        conn_layout.addWidget(QLabel("起始位址:"), 2, 1)
        self.m_start_input = QLineEdit("0")
        conn_layout.addWidget(self.m_start_input, 2, 2)
        conn_layout.addWidget(QLabel("數量:"), 2, 3)
        self.m_count_input = QLineEdit("10")
        conn_layout.addWidget(self.m_count_input, 2, 4)

        # 讀取 D 值設定
        self.d_checkbox = QCheckBox("讀取 D 值")
        self.d_checkbox.setChecked(True)
        conn_layout.addWidget(self.d_checkbox, 3, 0)
        conn_layout.addWidget(QLabel("起始位址:"), 3, 1)
        self.d_start_input = QLineEdit("40960") # D1000 對應的 Modbus 地址
        conn_layout.addWidget(self.d_start_input, 3, 2)
        conn_layout.addWidget(QLabel("數量:"), 3, 3)
        self.d_count_input = QLineEdit("10")
        conn_layout.addWidget(self.d_count_input, 3, 4)
        
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

    def refresh_com_ports(self):
        """刷新並列出可用的 COM 埠"""
        self.com_port_combo.clear()
        ports = serial.tools.list_ports.comports()
        if not ports:
            self.com_port_combo.addItem("無可用COM埠")
        else:
            for port in ports:
                self.com_port_combo.addItem(port.device)

    def create_menu_bar(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("檔案(&F)")
        
        save_action = QAction("儲存為CSV", self)
        save_action.setShortcut("Ctrl+S")
        save_action.setStatusTip("將讀取到的資料儲存為CSV檔案")
        save_action.triggered.connect(self.save_to_csv)
        file_menu.addAction(save_action)

    def start_reading(self):
        port_name = self.com_port_combo.currentText()
        if port_name == "無可用COM埠":
            QMessageBox.warning(self, "警告", "未選擇COM埠。")
            return
            
        try:
            slave_id = int(self.slave_id_input.text())
            baudrate = int(self.baudrate_combo.currentText())
            parity_str = self.parity_combo.currentText()
            parity = serial.PARITY_EVEN if parity_str == "EVEN" else serial.PARITY_ODD if parity_str == "ODD" else serial.PARITY_NONE
            timeout = 1  # 暫時固定為 1 秒

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

            self.thread = PlcReaderThread(port_name, slave_id, baudrate, parity, timeout, m_config, d_config)
            self.thread.data_ready.connect(self.update_data)
            self.thread.start()
            
            self.log_message(f"開始連線並讀取資料... COM埠: {port_name}, 站號: {slave_id}")
            self.connect_btn.setEnabled(False)

        except ValueError:
            QMessageBox.warning(self, "警告", "請輸入有效的數字。")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"無法啟動讀取：{e}")

    def update_data(self, data):
        self.connect_btn.setEnabled(True)
        status = data.get("status")
        
        if status == "data":
            m_values = data.get("m_values")
            d_values = data.get("d_values")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            m_str = "未讀取" if m_values is None else f"M{self.m_start_input.text()}: {m_values}"
            d_str = "未讀取" if d_values is None else f"D{self.d_start_input.text()}: {d_values}"

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
            filename = f"modbus_485_data_{timestamp}.csv"
            
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
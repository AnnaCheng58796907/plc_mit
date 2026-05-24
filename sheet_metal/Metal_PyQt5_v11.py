import sys
import os
import math
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QTabWidget,
                             QLabel, QComboBox, QLineEdit, QSpinBox, QPushButton,
                             QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
                             QMessageBox, QSplitter, QScrollArea)
from PyQt5.QtGui import QPainter, QPen, QColor, QFont
from PyQt5.QtCore import Qt, QRectF

class DrawWidget(QWidget):
    """專用的繪圖畫布組件"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.mode = "none" # "bend", "cylinder", "cone", "none"
        self.draw_data = {}
        setBackgroundRole = self.setAutoFillBackground(True)
        
        # 設定畫布背景為白色
        p = self.palette()
        p.setColor(self.backgroundRole(), QColor("white"))
        self.setPalette(p)

    def set_bend_data(self, side_lengths, angles):
        self.mode = "bend"
        self.draw_data = {"sides": side_lengths, "angles": angles}
        self.update()

    def set_cylinder_data(self, L, H):
        self.mode = "cylinder"
        self.draw_data = {"L": L, "H": H}
        self.update()

    def set_cone_data(self, R, r, theta):
        self.mode = "cone"
        self.draw_data = {"R": R, "r": r, "theta": theta}
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing) # 開啟反鋸齒，線條超平滑
        
        if self.mode == "bend":
            self.draw_bend(painter)
        elif self.mode == "cylinder":
            self.draw_cylinder(painter)
        elif self.mode == "cone":
            self.draw_cone(painter)

    def draw_bend(self, painter):
        sides = self.draw_data.get("sides", [])
        angles = self.draw_data.get("angles", [])
        if not sides: return

        pen = QPen(QColor("#333333"), 3, Qt.SolidLine)
        painter.setPen(pen)
        
        x, y = 80.0, 400.0
        current_angle = 0.0
        scale = 2.0  # 1mm = 2px

        for i, length_mm in enumerate(sides):
            length = length_mm * scale
            rad = math.radians(current_angle)
            nx = x + length * math.cos(rad)
            ny = y - length * math.sin(rad) # Qt Y軸向下

            painter.drawLine(int(x), int(y), int(nx), int(ny))
            
            # 標註文字
            painter.setFont(QFont("Arial", 9))
            painter.drawText(int(x + (nx-x)/2), int(y + (ny-y)/2 - 8), f"L{i+1}")

            x, y = nx, ny
            if i < len(angles):
                current_angle += (180.0 - angles[i])

    def draw_cylinder(self, painter):
        L = self.draw_data.get("L", 0)
        H = self.draw_data.get("H", 0)
        if L <= 0 or H <= 0: return

        scale = 0.6
        painter.setPen(QPen(QColor("#1976D2"), 2, Qt.SolidLine))
        painter.setBrush(QColor("#E3F2FD"))
        
        rect = QRectF(50, 150, L * scale, H * scale)
        painter.drawRect(rect)
        
        painter.setPen(QPen(QColor("black")))
        painter.drawText(int(50 + (L*scale)/2 - 30), 140, f"展開長: {L:.2f}")
        painter.drawText(20, int(150 + (H*scale)/2), f"H: {H:.1f}")

    def draw_cone(self, painter):
        R = self.draw_data.get("R", 0)
        r = self.draw_data.get("r", 0)
        theta = self.draw_data.get("theta", 0)
        if R <= 0: return

        scale = 1.0
        if R > 350: scale = 350.0 / R

        cx, cy = 250.0, 80.0 # 扇形圓心
        start_ang = 270.0 - (theta / 2.0)

        # Qt 角度單位為 1/16 度
        qt_start = int(start_ang * 16)
        qt_span = int(theta * 16)

        # 畫外弧與內弧
        painter.setPen(QPen(QColor("red"), 2, Qt.SolidLine))
        rect_R = QRectF(cx - R*scale, cy - R*scale, R*2*scale, R*2*scale)
        painter.drawArc(rect_R, qt_start, qt_span)

        painter.setPen(QPen(QColor("blue"), 2, Qt.SolidLine))
        rect_r = QRectF(cx - r*scale, cy - r*scale, r*2*scale, r*2*scale)
        painter.drawArc(rect_r, qt_start, qt_span)

        # 畫兩端連接線
        painter.setPen(QPen(QColor("black"), 2, Qt.SolidLine))
        for ang in [start_ang, start_ang + theta]:
            rad = math.radians(ang)
            x1 = cx + r * scale * math.cos(rad)
            y1 = cy - r * scale * math.sin(rad)
            x2 = cx + R * scale * math.cos(rad)
            y2 = cy - R * scale * math.sin(rad)
            painter.drawLine(int(x1), int(y1), int(x2), int(y2))


class MetalApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("板金設計全功能工具箱 v11.0 (PyQt5)")
        self.setGeometry(100, 50, 1100, 850)

        self.excel_file = "bend_parameters.xlsx"
        self.all_sheets = {}
        self.side_entries = []
        self.angle_entries = []
        self.layout_inputs = {}

        self.load_all_excel_sheets()
        self.init_ui()

    def load_all_excel_sheets(self):
        if os.path.exists(self.excel_file):
            try:
                xls = pd.ExcelFile(self.excel_file)
                for name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=name, index_col=0)
                    df.columns = [str(col) for col in df.columns]
                    self.all_sheets[name] = df
            except Exception as e:
                QMessageBox.critical(self, "Excel 錯誤", f"讀取失敗: {str(e)}")
        else:
            # 建立空測試資料
            self.all_sheets = {"範例黑鐵_V10": pd.DataFrame(index=["SPCC"], columns=["1.0", "2.0"])}

    def init_ui(self):
        # 使用 QSplitter 建立可左右拉動的左右分割面板
        splitter = QSplitter(Qt.Horizontal)
        self.setCentralWidget(splitter)

        # 左側控制區
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        self.tabs = QTabWidget()
        left_layout.addWidget(self.tabs)
        splitter.addWidget(left_widget)

        # 右側繪圖區
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(QLabel("📐 即時圖形放樣示意 (非精密加工比例)"))
        
        self.canvas = DrawWidget()
        right_layout.addWidget(self.canvas)
        splitter.addWidget(right_widget)

        # 調整左右初始比例
        splitter.setSizes([450, 650])

        # 初始化分頁
        self.init_bend_tab()
        self.init_special_tab()
        self.init_hw_tab()
        self.init_layout_tab()

    # --- 分頁 1: 動態角度折彎 ---
    def init_bend_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 1. 係數工作表選擇
        g1 = QGroupBox("1. 選擇係數來源工作表")
        f1 = QFormLayout(g1)
        self.c_sheet_sel = QComboBox()
        bend_sheets = [s for s in self.all_sheets.keys() if s not in ["Hardware", "Special"]]
        self.c_sheet_sel.addItems(bend_sheets)
        self.c_sheet_sel.currentIndexChanged.connect(self.on_sheet_change)
        f1.addRow("Excel 分頁:", self.c_sheet_sel)
        layout.addWidget(g1)

        # 2. 規格與 K90
        g2 = QGroupBox("2. 規格檢索")
        f2 = QFormLayout(g2)
        self.c_row = QComboBox()
        self.c_col = QComboBox()
        self.e_k90 = QLineEdit()
        self.c_row.currentIndexChanged.connect(self.update_k90_val)
        self.c_col.currentIndexChanged.connect(self.update_k90_val)
        f2.addRow("項目 / 材質:", self.c_row)
        f2.addRow("板厚 (T):", self.c_col)
        f2.addRow("90° 補償值:", self.e_k90)
        layout.addWidget(g2)

        # 3. 尺寸輸入區 (加滾動條防溢出)
        g3 = QGroupBox("3. 輸入尺寸與角度 (內邊相加法)")
        v3 = QVBoxLayout(g3)
        
        self.s_n = QSpinBox()
        self.s_n.setRange(1, 15)
        self.s_n.setValue(2)
        self.s_n.valueChanged.connect(self.refresh_bend_ui)
        v3.addWidget(QLabel("折彎次數 (n):"))
        v3.addWidget(self.s_n)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.bend_area_container = QWidget()
        self.bend_form = QFormLayout(self.bend_area_container)
        self.scroll_area.setWidget(self.bend_area_container)
        v3.addWidget(self.scroll_area)
        layout.addWidget(g3)

        # 按鈕與結果
        self.btn_calc = QPushButton("執行計算並繪圖")
        self.btn_calc.setStyleSheet("background-color: #1976D2; color: white; font-weight: bold; font-size: 14px; padding: 6px;")
        self.btn_calc.clicked.connect(self.calculate_bend)
        layout.addWidget(self.btn_calc)

        self.l_res = QLabel("展開長度: --")
        self.l_res.setStyleSheet("color: red; font-size: 16px; font-weight: bold;")
        self.l_res.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.l_res)

        self.tabs.addTab(tab, "📐 動態折彎")
        self.on_sheet_change() # 觸發首次初始化

    def on_sheet_change(self):
        sheet = self.c_sheet_sel.currentText()
        if not sheet or sheet not in self.all_sheets: return
        df = self.all_sheets[sheet]
        
        self.c_row.blockSignals(True)
        self.c_col.blockSignals(True)
        self.c_row.clear()
        self.c_col.clear()
        self.c_row.addItems(df.index.astype(str).tolist())
        self.c_col.addItems(df.columns.astype(str).tolist())
        self.c_row.blockSignals(False)
        self.c_col.blockSignals(False)
        self.update_k90_val()
        self.refresh_bend_ui()

    def update_k90_val(self):
        sheet = self.c_sheet_sel.currentText()
        r = self.c_row.currentText()
        c = self.c_col.currentText()
        if sheet and r and c:
            try:
                val = self.all_sheets[sheet].loc[r, c]
                self.e_k90.setText(str(val))
            except:
                self.e_k90.clear()

    def refresh_bend_ui(self):
        # 清除舊輸入框
        while self.bend_form.count():
            child = self.bend_form.takeAt(0)
            if child.widget(): child.widget().deleteLater()
            
        self.side_entries = []
        self.angle_entries = []
        n = self.s_n.value()

        for i in range(n + 1):
            le = QLineEdit("50")
            self.bend_form.addRow(f"內邊長 L{i+1} (mm):", le)
            self.side_entries.append(le)
            
            if i < n:
                ae = QLineEdit("90")
                self.bend_form.addRow(f"  ↳ 折彎角度 Angle {i+1} (°):", ae)
                self.angle_entries.append(ae)

    def calculate_bend(self):
        try:
            k90 = float(self.e_k90.text())
            sides = [float(e.text() or 0) for e in self.side_entries]
            angles = [float(e.text() or 90) for e in self.angle_entries]
            
            sum_l = sum(sides)
            sum_k = sum([(k90 / 90.0) * (180.0 - a) for a in angles])
            total = sum_l + sum_k
            
            self.l_res.setText(f"展開總長: {total:.3f} mm")
            self.canvas.set_bend_data(sides, angles) # 送資料去畫布畫圖
        except Exception as e:
            QMessageBox.warning(self, "輸入錯誤", "請確認所有尺寸與 K0 值皆為有效數字。")

    # --- 分頁 2: 特殊固定折彎 ---
    def init_special_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        f = QFormLayout()
        
        self.c_sp_type = QComboBox()
        self.c_sp_thick = QComboBox()
        df = self.all_sheets.get("Special", pd.DataFrame())
        self.c_sp_type.addItems(df.index.astype(str).tolist())
        self.c_sp_thick.addItems(df.columns.astype(str).tolist())
        
        self.c_sp_type.currentIndexChanged.connect(self.update_special_res)
        self.c_sp_thick.currentIndexChanged.connect(self.update_special_res)
        
        f.addRow("特殊折彎工法:", self.c_sp_type)
        f.addRow("選用板厚:", self.c_sp_thick)
        layout.addLayout(f)

        self.l_sp_res = QLabel("補償值: --")
        self.l_sp_res.setStyleSheet("font-size: 20px; color: #D32F2F; font-weight: bold;")
        self.l_sp_res.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.l_sp_res)
        layout.addStretch()
        self.tabs.addTab(tab, "🌀 特殊折彎")

    def update_special_res(self):
        t = self.c_sp_type.currentText()
        th = self.c_sp_thick.currentText()
        if t and th and "Special" in self.all_sheets:
            try:
                val = self.all_sheets["Special"].loc[t, th]
                self.l_sp_res.setText(f"固定補償值: {val} mm")
            except:
                self.l_sp_res.setText("查無資料")

    # --- 分頁 3: 螺紋與鉚合零件 ---
    def init_hw_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        f = QFormLayout()
        
        self.c_hw_type = QComboBox()
        self.c_hw_spec = QComboBox()
        df = self.all_sheets.get("Hardware", pd.DataFrame())
        self.c_hw_type.addItems(df.index.astype(str).tolist())
        self.c_hw_spec.addItems(df.columns.astype(str).tolist())
        
        self.c_hw_type.currentIndexChanged.connect(self.update_hw_res)
        self.c_hw_spec.currentIndexChanged.connect(self.update_hw_res)
        
        f.addRow("零件種類:", self.c_hw_type)
        f.addRow("規格螺牙:", self.c_hw_spec)
        layout.addLayout(f)

        self.l_hw_res = QLabel("Ø --")
        self.l_hw_res.setStyleSheet("font-size: 32px; color: #1976D2; font-weight: bold;")
        self.l_hw_res.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.l_hw_res)
        layout.addStretch()
        self.tabs.addTab(tab, "🔩 零件查詢")

    def update_hw_res(self):
        t = self.c_hw_type.currentText()
        s = self.c_hw_spec.currentText()
        if t and s and "Hardware" in self.all_sheets:
            try:
                val = self.all_sheets["Hardware"].loc[t, s]
                self.l_hw_res.setText(f"建議開孔: Ø {val}")
            except:
                self.l_hw_res.setText("Ø --")

    # --- 分頁 4: 圓柱/圓錐放樣 (幾何展開) ---
    def init_layout_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        f = QFormLayout()
        self.c_lay_type = QComboBox()
        self.c_lay_type.addItems(["圓柱體 (Cylinder)", "圓錐體 (Cone)"])
        self.c_lay_type.currentIndexChanged.connect(self.refresh_layout_fields)
        f.addRow("放樣類型:", self.c_lay_type)
        layout.addLayout(f)

        self.lay_group = QGroupBox("輸入放樣參數")
        self.lay_form = QFormLayout(self.lay_group)
        layout.addWidget(self.lay_group)

        self.btn_lay = QPushButton("執行放樣計算與繪圖")
        self.btn_lay.clicked.connect(self.calculate_layout)
        self.btn_lay.setStyleSheet("background-color: #2E7D32; color: white; font-weight: bold; padding: 5px;")
        layout.addWidget(self.btn_lay)

        self.l_lay_res = QLabel("放樣結果將在此顯示")
        self.l_lay_res.setStyleSheet("font-weight: bold; color: green;")
        layout.addWidget(self.l_lay_res)
        layout.addStretch()
        
        self.tabs.addTab(tab, "🌀 幾何放樣")
        self.refresh_layout_fields()

    def refresh_layout_fields(self):
        while self.lay_form.count():
            child = self.lay_form.takeAt(0)
            if child.widget(): child.widget().deleteLater()
            
        self.layout_inputs = {}
        t = self.c_lay_type.currentText()
        
        if "圓柱" in t:
            self.layout_inputs["d"] = QLineEdit("100")
            self.layout_inputs["T"] = QLineEdit("2.0")
            self.layout_inputs["H"] = QLineEdit("150")
            self.lay_form.addRow("內直徑 (d):", self.layout_inputs["d"])
            self.lay_form.addRow("板厚 (T):", self.layout_inputs["T"])
            self.lay_form.addRow("高度 (H):", self.layout_inputs["H"])
        else:
            self.layout_inputs["D"] = QLineEdit("150")
            self.layout_inputs["d"] = QLineEdit("80")
            self.layout_inputs["H"] = QLineEdit("100")
            self.layout_inputs["T"] = QLineEdit("2.0")
            self.lay_form.addRow("大端內徑 (D):", self.layout_inputs["D"])
            self.lay_form.addRow("小端內徑 (d):", self.layout_inputs["d"])
            self.lay_form.addRow("垂直高度 (H):", self.layout_inputs["H"])
            self.lay_form.addRow("板厚 (T):", self.layout_inputs["T"])

    def calculate_layout(self):
        t = self.c_lay_type.currentText()
        try:
            if "圓柱" in t:
                d = float(self.layout_inputs["d"].text())
                T = float(self.layout_inputs["T"].text())
                H = float(self.layout_inputs["H"].text())
                
                L = math.pi * (d + T) # 中軸層周長
                self.l_lay_res.setText(f"展開長方形尺寸:\n長度 (L): {L:.2f} mm\n寬度 (H): {H:.1f} mm")
                self.canvas.set_cylinder_data(L, H)
                
            elif "圓錐" in t:
                D_in = float(self.layout_inputs["D"].text())
                d_in = float(self.layout_inputs["d"].text())
                H = float(self.layout_inputs["H"].text())
                T = float(self.layout_inputs["T"].text())
                
                if D_in == d_in:
                    QMessageBox.warning(self, "資料錯誤", "大小端內徑不能相同，請使用圓柱體模式。")
                    return
                    
                D = D_in + T
                d = d_in + T
                
                # 斜邊高
                slant = math.sqrt(H**2 + ((D - d) / 2.0)**2)
                R = (D * slant) / (D - d)
                r = R - slant
                theta = (D * 180.0) / R
                
                self.l_lay_res.setText(f"扇形放樣數據:\n外半徑 (R): {R:.2f} mm\n內半徑 (r): {r:.2f} mm\n展開角度 (θ): {theta:.2f}°")
                self.canvas.set_cone_data(R, r, theta)
                
        except Exception as e:
            QMessageBox.warning(self, "輸入錯誤", "請檢查所有參數欄位是否為正確數字。")

if __name__ == "__main__":
    # --- 1. 關鍵核心：強制 PyQt5 配合系統螢幕與文字縮放設定 ---
    import os
    from PyQt5.QtCore import Qt
    
    # 啟用高 DPI 自動縮放
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    # 讓圖示與控制元件也配合高 DPI 縮放
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    # --- 2. 啟動應用程式 ---
    app = QApplication(sys.argv)
    window = MetalApp()
    window.show()
    sys.exit(app.exec_())

# main.py
import sys
import os
import math
from pem_viewer import AdvancedPemPdfViewer  # 👈 匯入我們寫好的自動搜尋 PDF 元件
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QTabWidget,
                             QLabel, QComboBox, QLineEdit, QSpinBox, QPushButton,
                             QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox, 
                             QMessageBox, QScrollArea, QGridLayout, QSplitter)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

# 導入原本的副程式
from calculations import SheetMetalCalc
from drawing import DrawWidget
from special_features import SpecialTabsManager

class SheetMetalApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("板金展開與五金查詢系統 v14.0")
        self.setGeometry(100, 50, 1320, 950) # 寬度與高度微調，提供極佳操作視野

        # 統一調大全域介面字型
        global_font = QFont("Microsoft JhengHei", 12)
        QApplication.setFont(global_font)

        try:
            self.calc = SheetMetalCalc()
        except Exception as e:
            QMessageBox.critical(self, "核心加載失敗", str(e))
            sys.exit(1)

        # 💡 修正：移除了原本在這裡會導致報錯的 self.r_entries 綁定，改到元件建立時動態綁定
        self.side_entries = []
        self.angle_entries = []
        self.r_entries = [] # 確保初始化變數存在
        self.layout_inputs = {}
        self.init_ui()

    def init_ui(self):
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("QTabBar::tab { min-width: 150px; padding: 8px; font-weight: bold; }")
        self.setCentralWidget(self.tabs)

        self.init_bend_tab()
        
        # 🚀 實體化 PDF 元件（請確保已從 pem_viewer 匯入 AdvancedPemPdfViewer）
        from pem_viewer import AdvancedPemPdfViewer
        self.pem_pdf_widget = AdvancedPemPdfViewer()
        
        # 🚀 將 PDF 元件當作引數傳入
        self.special_mgr = SpecialTabsManager(self.tabs, self.calc, self.pem_pdf_widget)
        
        self.init_layout_tab()

    def init_bend_tab(self):
        """1. 動態折彎分頁"""
        tab = QWidget()
        main_layout = QHBoxLayout(tab) 
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(12)

        # ─── 左側：參數、條件與結果控制台 ───
        left_panel = QVBoxLayout()
        
        g_param = QGroupBox("1. 參數與製程選擇")
        f_param = QFormLayout(g_param)
        self.c_sheet = QComboBox()
        self.c_sheet.addItems(self.calc.get_bend_sheets())
        self.s_bends = QSpinBox()
        self.s_bends.setRange(1, 15) # 擴充支援到 15 折
        self.s_bends.setValue(2)
        self.s_bends.valueChanged.connect(self.update_bend_inputs)
        f_param.addRow("選擇板金參數表:", self.c_sheet)
        f_param.addRow("折彎次數 (N):", self.s_bends)
        left_panel.addWidget(g_param)

        g_rc = QGroupBox("2. 選擇條件")
        f_rc = QFormLayout(g_rc)
        self.c_thick = QComboBox()
        self.c_r = QComboBox()
        self.l_current_k = QLabel("當前對應係數: --")
        self.l_current_k.setStyleSheet("color: #1565C0; font-weight: bold;")
        
        self.c_sheet.currentTextChanged.connect(self.update_rc_combos)
        self.c_thick.currentTextChanged.connect(self.update_current_k_display)
        self.c_r.currentTextChanged.connect(self.update_current_k_display)
        
        # 當板厚選擇改變時，一併同步刷新畫布與 R1 標籤
        self.c_thick.currentTextChanged.connect(self.update_r1_label_and_canvas)
        
        f_rc.addRow("選用板厚 (T):", self.c_thick)
        f_rc.addRow("內折半徑 (R):", self.c_r)
        f_rc.addRow("📄 查表狀態:", self.l_current_k)
        left_panel.addWidget(g_rc)
        
        g_res = QGroupBox("📊 運算與計算結果")
        v_res = QVBoxLayout(g_res)
        v_res.setSpacing(10)
        
        self.btn_calc = QPushButton("🚀 執行計算並同步繪圖")
        self.btn_calc.setStyleSheet("""
            font-size: 16px; font-weight: bold; color: white; 
            background-color: #2E7D32; padding: 12px; border-radius: 5px;
        """)
        self.btn_calc.clicked.connect(self.process_bend_calculation)
        
        self.l_total_l = QLabel("展開總長: 0.00 mm")
        self.l_total_l.setStyleSheet("""
            font-size: 26px; color: #E65100; font-weight: bold; 
            background-color: #FFF3E0; border: 1px solid #FFE0B2;
            border-radius: 4px; padding: 12px;
        """)
        
        v_res.addWidget(self.btn_calc)
        v_res.addWidget(self.l_total_l)
        left_panel.addWidget(g_res)
        
        left_panel.addStretch()
        main_layout.addLayout(left_panel, stretch=1)

        # ─── 右側：💡 改造為可自由拖曳高度的 Splitter 佈局 ───
        g_combined = QGroupBox("3. 輸入尺寸與角度(內邊相加法) ＆ 圖形即時放樣")
        combined_outer_layout = QVBoxLayout(g_combined)
        combined_outer_layout.setContentsMargins(10, 20, 10, 10)

        # 💡 建立上下分割器 (讓使用者可以用滑鼠任意拖拉改變上下高度範圍)
        self.right_splitter = QSplitter(Qt.Vertical)
        self.right_splitter.setHandleWidth(7) # 讓中間那條分隔線明顯一點，方便滑鼠抓取
        self.right_splitter.setStyleSheet("QSplitter::handle { background-color: #D3D3D3; border-radius: 2px; }")

        # 1️⃣ 上方組件：智慧並排滾動區
        input_scroll = QScrollArea()
        input_scroll.setWidgetResizable(True)
        input_scroll.setFrameShape(QScrollArea.StyledPanel)
        
        self.scroll_widget = QWidget()
        # 💡 將原本的單列 FormLayout 改為強大的 QGridLayout 達成動態並排
        self.input_grid = QGridLayout(self.scroll_widget)
        self.input_grid.setSpacing(12)
        self.input_grid.setContentsMargins(5, 5, 5, 5)
        input_scroll.setWidget(self.scroll_widget)
        
        # 2️⃣ 下方組件：完全頂滿的放樣畫布區
        canvas_container = QWidget()
        canvas_grid = QGridLayout(canvas_container)
        canvas_grid.setContentsMargins(0, 0, 0, 0)
        canvas_grid.setSpacing(0)

        self.canvas = DrawWidget()
        canvas_grid.addWidget(self.canvas, 0, 0)

        title_label = QLabel("📐 即時放樣示意 (可拖曳調整高度)")
        title_label.setStyleSheet("""
            font-size: 13px; font-weight: bold; color: #666666; 
            background-color: rgba(255, 255, 255, 200); 
            padding: 4px 8px; border-bottom-right-radius: 5px;
        """)
        canvas_grid.addWidget(title_label, 0, 0, Qt.AlignLeft | Qt.AlignTop)

        # 將上下兩個大組件塞進分割器中
        self.right_splitter.addWidget(input_scroll)
        self.right_splitter.addWidget(canvas_container)
        
        # 💡 設定預設高度分配：上方輸入框佔 200 像素，其餘全部給下方畫布
        self.right_splitter.setSizes([200, 650])
        
        # 將分割器放入右側大框內
        combined_outer_layout.addWidget(self.right_splitter)
        main_layout.addWidget(g_combined, stretch=3)

        # 初始化連動
        self.update_rc_combos(self.c_sheet.currentText())
        self.update_bend_inputs()
        self.tabs.addTab(tab, "📐 動態折彎展開")

    def init_layout_tab(self):
        """2. 幾何特殊放樣分頁"""
        tab = QWidget()
        main_layout = QHBoxLayout(tab)
        
        left_panel = QVBoxLayout()
        g_mode = QGroupBox("1. 放樣類型")
        f_mode = QFormLayout(g_mode)
        self.c_mode = QComboBox()
        self.c_mode.addItems(["圓柱體 (管狀)", "圓錐體 (大小頭)"])
        self.c_mode.currentTextChanged.connect(self.switch_layout_mode)
        f_mode.addRow("選擇幾何形狀:", self.c_mode)
        left_panel.addWidget(g_mode)

        self.g_inputs = QGroupBox("2. 輸入加工參數 (mm)")
        self.f_inputs = QFormLayout(self.g_inputs)
        left_panel.addWidget(self.g_inputs)

        self.btn_layout_calc = QPushButton("⚙️ 產出展開數據與繪圖")
        self.btn_layout_calc.setStyleSheet("""
            font-size: 16px; font-weight: bold; color: white; 
            background-color: #1565C0; padding: 12px; border-radius: 5px;
        """)
        self.btn_layout_calc.clicked.connect(self.process_layout_calculation)
        left_panel.addWidget(self.btn_layout_calc)
        
        self.g_layout_res = QGroupBox("📊 展開幾何數據")
        self.v_layout_res = QVBoxLayout(self.g_layout_res)
        self.l_layout_out = QLabel("請輸入參數並點擊計算")
        self.l_layout_out.setStyleSheet("font-size: 15px; font-weight: bold; color: #1565C0;")
        self.v_layout_res.addWidget(self.l_layout_out)
        left_panel.addWidget(self.g_layout_res)
        
        left_panel.addStretch()
        main_layout.addLayout(left_panel, stretch=1)

        g_canvas = QGroupBox("特殊圖形展開放樣示意")
        canvas_layout = QVBoxLayout(g_canvas)
        canvas_layout.setContentsMargins(5, 15, 5, 5)
        self.layout_canvas = DrawWidget()
        canvas_layout.addWidget(self.layout_canvas)
        
        main_layout.addWidget(g_canvas, stretch=2)
        
        self.switch_layout_mode(self.c_mode.currentText())
        self.tabs.addTab(tab, "🌀 幾何特殊放樣")

    # ──── 邏輯與動態排版控制核心 ────
    def update_rc_combos(self, sheet_name):
        df = self.calc.get_sheet_data(sheet_name)
        self.c_thick.clear()
        self.c_r.clear()
        if not df.empty:
            self.c_thick.addItems(df.columns.tolist())
            self.c_r.addItems(df.index.tolist())
        self.update_current_k_display()

    def update_current_k_display(self):
        sheet = self.c_sheet.currentText()
        r = self.c_r.currentText()
        c = self.c_thick.currentText()
        if not r or not c: return
        
        k_val = self.calc.get_k90_value(sheet, r, c)
        if k_val == "":
            self.l_current_k.setText("⚠️ 參數表中無此規格數據")
            self.l_current_k.setStyleSheet("color: red; font-weight: bold;")
        else:
            self.l_current_k.setText(f"折彎內縮補償量: {k_val} mm")
            self.l_current_k.setStyleSheet("color: #2E7D32; font-weight: bold;")

    def update_bend_inputs(self):
        """動態格狀橫向並排，預設 R 角為 0，並綁定即時標籤與畫布刷新"""
        # 清除舊元件
        for w in self.side_entries + self.angle_entries + getattr(self, 'r_entries', []):
            w.deleteLater()
        self.side_entries.clear()
        self.angle_entries.clear()
        self.r_entries = []
        
        while self.input_grid.count() > 0:
            item = self.input_grid.takeAt(0)
            if item.widget(): 
                item.widget().deleteLater()

        n = self.s_bends.value()
        COLUMNS_PER_ROW = 3 
        
        self.lbl_sides_list = []
        self.lbl_r1 = None # 💡 用來儲存第一道折彎的 R1 標籤控制指標

        current_item_index = 0
        for i in range(1, n + 2):
            row = current_item_index // COLUMNS_PER_ROW
            col = (current_item_index % COLUMNS_PER_ROW) * 3 
            
            # --- 1. 建立邊長元件 (L) ---
            h_layout_l = QHBoxLayout()
            lbl_l = QLabel(f"L{i}:")
            lbl_l.setStyleSheet("font-weight: bold; color: #444444;")
            le = QLineEdit("50")
            le.setFixedWidth(85)
            le.textChanged.connect(self.process_bend_calculation)
            
            h_layout_l.addWidget(lbl_l)
            h_layout_l.addWidget(le)
            h_layout_l.addStretch()
            
            container_l = QWidget()
            container_l.setLayout(h_layout_l)
            self.input_grid.addWidget(container_l, row, col)
            
            self.side_entries.append(le)
            self.lbl_sides_list.append(lbl_l)
            
            current_item_index += 1

            # --- 2. 建立折彎元件 (A 角度 & R 角) ---
            if i <= n:
                row_a = current_item_index // COLUMNS_PER_ROW
                col_a = (current_item_index % COLUMNS_PER_ROW) * 3
                
                h_layout_b = QHBoxLayout()
                
                # 角度輸入框 (A)
                lbl_a = QLabel(f"A{i}(°):")
                lbl_a.setStyleSheet("color: #0D47A1; font-weight: bold;")
                sb = QSpinBox()
                sb.setRange(0, 179) 
                sb.setValue(90)
                sb.setFixedWidth(65)
                sb.setKeyboardTracking(False)
                sb.setLineEdit(QLineEdit())
                sb.valueChanged.connect(self.refresh_t_labels) # 角度變更刷新標籤
                sb.valueChanged.connect(self.process_bend_calculation)
                
                # 💡 針對 A1 的變化，即時觸發 R1 旁邊標籤與畫布繪圖重新計算
                if i == 1:
                    sb.valueChanged.connect(self.update_r_label_display)
                    sb.valueChanged.connect(self.canvas.update)

                self.angle_entries.append(sb)
                
                # R 角輸入框 (R) -> 💡 依照你的要求，預設一律為 0
                lbl_r = QLabel(f"R{i}:")
                lbl_r.setStyleSheet("color: #2E7D32; font-weight: bold;")
                le_r = QLineEdit("0") 
                le_r.setFixedWidth(45)
                
                # 💡 關鍵：當 R 角數字改變時，除了重新計算，也要立刻強迫刷新 0T/1T/2T 標籤！
                le_r.textChanged.connect(self.refresh_t_labels) 
                le_r.textChanged.connect(self.process_bend_calculation)
                
                # 💡 針對 R1 欄位：動態綁定剛才在 __init__ 報錯的連動更新與畫布 repaint
                if i == 1:
                    self.lbl_r1 = lbl_r # 將 R1 的 QLabel 指標留下來，方便後續更改文字
                    le_r.textChanged.connect(self.update_r_label_display)
                    le_r.textChanged.connect(self.canvas.update)

                self.r_entries.append(le_r)
                
                h_layout_b.addWidget(lbl_a)
                h_layout_b.addWidget(sb)
                h_layout_b.addWidget(lbl_r)
                h_layout_b.addWidget(le_r)
                h_layout_b.addStretch()
                
                container_b = QWidget()
                container_b.setLayout(h_layout_b)
                self.input_grid.addWidget(container_b, row_a, col_a)
                
                current_item_index += 1
                
        # 初始化刷新
        self.refresh_t_labels()
        self.update_r_label_display()

    def update_r_label_display(self):
        """💡 當 R1角、角度、板厚改變時，同步把弧長 L 顯示在 R1 輸入框旁邊的 Label"""
        # 安全防呆
        if not hasattr(self, 'r_entries') or not self.r_entries or not self.lbl_r1:
            return

        try:
            r_val = float(self.r_entries[0].text())
            angle_val = self.angle_entries[0].value()
            # 取得當前選用的板厚 T
            t_val = float(self.c_thick.currentText()) if self.c_thick.currentText() else 2.0
            # 使用固定中性層係數 0.5 (與計算公式一致)
            k_factor = 0.5 
        except (IndexError, ValueError):
            return

        # 計算中性線弧長 (注意：轉折外角 = 180 - 內折角度)
        if r_val > 0:
            outer_angle = 180.0 - angle_val
            neutral_radius = r_val + (k_factor * t_val)
            arc_length = (math.pi * neutral_radius * outer_angle) / 180.0
            
            # 更新 R1 輸入框左邊 Label 的文字與視覺外觀
            self.lbl_r1.setText(f"R1(L={arc_length:.2f}):")
            self.lbl_r1.setStyleSheet("color: #2E7D32; font-weight: bold;") # 綠色醒目提醒
        else:
            self.lbl_r1.setText("R1:")
            self.lbl_r1.setStyleSheet("color: #2E7D32; font-weight: bold;")

    def update_r1_label_and_canvas(self):
        """供板厚下拉選單變更時，一併驅動文字與畫布刷新"""
        self.update_r_label_display()
        self.canvas.update()

    def refresh_t_labels(self):
        """即時檢查角度與R角輸入，精確判定每條板邊相鄰的折彎狀態（已移除L1/L2標籤上的弧長贅字）"""
        n = self.s_bends.value()
        
        # 安全防呆
        if not hasattr(self, 'r_entries') or len(self.r_entries) < n:
            return

        for i in range(len(self.lbl_sides_list)):
            # --- 1. 檢查這條邊的「左側」有沒有緊鄰折彎 ---
            if i > 0:
                left_angle = self.angle_entries[i-1].value()
                try: left_r = float(self.r_entries[i-1].text())
                except ValueError: left_r = 0.0
                left_t = 1 if (left_angle == 90 and left_r == 0) else 0
            else:
                left_t = 0

            # --- 2. 檢查這條邊的「右側」有沒有緊鄰折彎 ---
            if i < n:
                right_angle = self.angle_entries[i].value()
                try: right_r = float(self.r_entries[i].text())
                except ValueError: right_r = 0.0
                right_t = 1 if (right_angle == 90 and right_r == 0) else 0
            else:
                right_t = 0
            
            # --- 3. 總計這條邊要扣幾個 T ---
            t_count = left_t + right_t
            
            # 檢查是否為死邊 (0度)
            is_dead_edge = False
            if i > 0 and self.angle_entries[i-1].value() == 0: is_dead_edge = True
            if i < n and self.angle_entries[i].value() == 0: is_dead_edge = True
            suffix = " [死邊]" if is_dead_edge else ""

            # --- 4. 刷新標籤顯示與顏色 (徹底移除 r_info 弧長，保持純淨) ---
            lbl = self.lbl_sides_list[i]
            if t_count == 2:
                lbl.setText(f"L{i+1} [2T]{suffix}:")
                lbl.setStyleSheet("font-weight: bold; color: #2E7D32;") 
            elif t_count == 1:
                lbl.setText(f"L{i+1} [1T]{suffix}:")
                lbl.setStyleSheet("font-weight: bold; color: #E65100;") 
            else:
                lbl.setText(f"L{i+1} [0T]{suffix}:")
                lbl.setStyleSheet("font-weight: bold; color: #C62828;")

    def process_bend_calculation(self):
        c_str = self.c_thick.currentText()
        k_factor = 0.5 # 中性層固定為板厚正中心 (T/2)
        
        try:
            t = float(c_str) # 板厚 T
        except:
            return

        try:
            sides = [float(le.text()) for le in self.side_entries]
            angles = [float(sb.value()) for sb in self.angle_entries]
            
            r_list = []
            for le_r in self.r_entries:
                try: r_list.append(float(le_r.text()))
                except ValueError: r_list.append(0.0) 
        except ValueError:
            return

        # 1️⃣ 第一步：平直段長度完全不需要扣板厚(0T)與R角，直接加總輸入的尺寸
        total_l = sum(sides)

        # 2️⃣ 第二步：逐一加上每道折彎的「中心線圓弧長度」
        for i in range(len(angles)):
            angle = angles[i]
            r_inner = r_list[i]        # 該道折彎的內 R 半徑
            outer_angle = 180.0 - angle # 計算轉折外角
            
            if r_inner > 0 or angle == 0:
                r_neutral = r_inner + (k_factor * t)
                arc_length = (math.pi * r_neutral * outer_angle) / 180.0
                total_l += arc_length

        # 3️⃣ 更新介面與畫布
        self.l_total_l.setText(f"展開總長: {total_l:.2f} mm")
        
        # 💡 將板厚與完整的 R 角清單傳給畫布，讓畫布有足夠的資料可以繪製「自然 R 角圓弧」
        if hasattr(self.canvas, 'set_bend_extended_data'):
            self.canvas.set_bend_extended_data(sides, angles, r_list, t)
        else:
            self.canvas.set_bend_data(sides, angles)

    def switch_layout_mode(self, mode):
        while self.f_inputs.count() > 0:
            item = self.f_inputs.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.layout_inputs.clear()

        if "圓柱體" in mode:
            self.add_layout_input("d_in", "圓柱內徑 (d):", "100")
            self.add_layout_input("T", "板厚 (T):", "2")
            self.add_layout_input("H", "圓柱高度 (H):", "200")
        else:
            self.add_layout_input("D_in", "大端內徑 (D_in):", "150")
            self.add_layout_input("d_in", "小端內徑 (d_in):", "80")
            self.add_layout_input("H", "錐體垂直高度 (H):", "120")
            self.add_layout_input("T", "板厚 (T):", "2")

    def add_layout_input(self, key, label, default_val):
        le = QLineEdit(default_val)
        le.setFixedWidth(120)
        self.f_inputs.addRow(label, le)
        self.layout_inputs[key] = le

    def process_layout_calculation(self):
        mode = self.c_mode.currentText()
        try:
            vals = {k: float(v.text()) for k, v in self.layout_inputs.items()}
        except ValueError:
            QMessageBox.warning(self, "輸入錯誤", "請檢查數值是否正確。")
            return

        if "圓柱體" in mode:
            res = self.calc.calculate_cylinder(vals["d_in"], vals["T"], vals["H"])
            self.l_layout_out.setText(f"★ 展開矩形尺寸:\n展開長度 (L): {res['L']:.2f} mm\n高度 (H): {res['H']:.1f} mm")
            self.layout_canvas.set_cylinder_data(res["L"], res["H"])
        else:
            try:
                res = self.calc.calculate_cone(vals["D_in"], vals["d_in"], vals["H"], vals["T"])
                self.l_layout_out.setText(
                    f"★ 扇形展開參數:\n外弧半徑 (R): {res['R']:.1f} mm\n內弧半徑 (r): {res['r']:.1f} mm\n展開夾角 (θ): {res['theta']:.1f}°"
                )
                self.layout_canvas.set_cone_data(res["R"], res["r"], res["theta"])
            except ValueError as ve:
                QMessageBox.warning(self, "計算錯誤", str(ve))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = SheetMetalApp()
    ex.show()
    sys.exit(app.exec_())
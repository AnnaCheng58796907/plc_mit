# special_features.py
from PyQt5.QtWidgets import QWidget, QFormLayout, QComboBox, QLabel, QVBoxLayout, QGroupBox
from PyQt5.QtCore import Qt

class SpecialTabsManager:
    """集中管理 鉚釘、抽孔等五金零件與特殊製程查詢的 UI 連動邏輯"""
    # 🚀 修改 __init__ 接收傳過來的 pem_pdf_widget
    def __init__(self, tab_widget, calc_model, pem_pdf_widget=None):
        self.tabs = tab_widget
        self.calc = calc_model
        self.pem_pdf_widget = pem_pdf_widget  # 存入類別屬性供下方使用
        
        # 依照您的 Excel 架構建立兩個專用查詢頁面
        self.init_hardware_tab(
            tab_title="🔩 鉚釘底孔查詢", 
            sheet_name="鉚釘預開孔徑", 
            type_label="鉚釘種類 / 規格:", 
            spec_label="對應板厚 / 參數:"
        )
        self.init_hardware_tab(
            tab_title="🕳️ 抽孔攻牙查詢", 
            sheet_name="抽孔攻牙預開孔徑", 
            type_label="螺牙規格 (M):", 
            spec_label="選用板厚 (T):"
        )

    def init_hardware_tab(self, tab_title, sheet_name, type_label, spec_label):
        """通用型五金表格動態生成器"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        df = self.calc.get_sheet_data(sheet_name)
        
        if df.empty:
            layout.addWidget(QLabel(f"⚠️ 在 Excel 中找不到 [{sheet_name}] 工作表"))
            self.tabs.addTab(tab, tab_title)
            return

        g_box = QGroupBox("規格篩選")
        f_layout = QFormLayout(g_box)
        
        c_type = QComboBox()
        c_spec = QComboBox()
        
        c_type.addItems(df.index.astype(str).tolist())
        c_spec.addItems(df.columns.astype(str).tolist())
        
        f_layout.addRow(type_label, c_type)
        f_layout.addRow(spec_label, c_spec)
        layout.addWidget(g_box)

        # 顯示查詢結果的大字級標籤
        l_res = QLabel("Ø --")
        l_res.setStyleSheet("font-size: 32px; color: #1976D2; font-weight: bold; margin: 20px;")
        l_res.setAlignment(Qt.AlignCenter)
        layout.addWidget(l_res)
        
        # 🚀 【安全防呆修正點】：只有鉚釘頁面才加載 PDF，且絕對不影響其他特殊圖形放樣頁面的 Stretch
        if sheet_name == "鉚釘預開孔徑" and self.pem_pdf_widget:
            # 先清除原本可能殘留的 Stretch 或干擾
            layout.addWidget(self.pem_pdf_widget)
            
            # 💡 精準指定鉚釘頁的拉伸比例：
            # 確保 0(篩選方塊)、1(大字標籤) 不被擠壓，2(PDF 區) 拿到剩餘的大空間
            layout.setStretch(0, 0) 
            layout.setStretch(1, 0) 
            layout.setStretch(2, 1) # 👈 改成 1 讓它自然填滿，不再用強烈的 3 去擠壓整體介面
        else:
            # 💡 關鍵：其他頁面（包含你的抽孔攻牙或特殊圖形展開放樣頁面）
            # 必須「維持你原本一模一樣的排版邏輯」，絕對不要去設定 layout.setStretch！
            layout.addStretch()
        
        # 綁定連動事件
        def update_result():
            t = c_type.currentText()
            s = c_spec.currentText()
            if t and s:
                try:
                    val = df.loc[t, s]
                    l_res.setText(f"建議預開孔: Ø {val}")
                except:
                    l_res.setText("查無對應資料")
            
            # 🚀 【關鍵修改點 2】：只有當處理「鉚釘」工作表且數據更新時，才通知 PDF 跑自動搜尋
            if sheet_name == "鉚釘預開孔徑" and self.pem_pdf_widget and t:
                # 這裡的 t 就會是 "SO螺柱"、"S NUT螺帽" 等等
                # 我們把這兩個字串丟給 pem_viewer.py 的 set_category_and_spec 去做前綴模糊匹配
                # 備註：因為板金通常用 t 或 s 作為牙規或牙長代號（例：M3），這裡直接把當前的種類名稱 t 傳過去
                self.pem_pdf_widget.set_category_and_spec(t, t)

        c_type.currentIndexChanged.connect(update_result)
        c_spec.currentIndexChanged.connect(update_result)
        
        self.tabs.addTab(tab, tab_title)
        update_result()  # 初始執行一次
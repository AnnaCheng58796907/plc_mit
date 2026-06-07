import sys
import os
import fitz  # PyMuPDF
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtCore import Qt

# 🗂️ 設定代碼對應的實際 PDF 檔名
PEM_CONFIG = {
    "SO": { "pdf_name": "SO.pdf", "specs": ["M3", "M4", "M5"] },
    "S": { "pdf_name": "S.pdf", "specs": ["M2", "M2.5", "M3", "M4", "M5", "M6", "M8"] },
    "FH": { "pdf_name": "FH.pdf", "specs": ["M3", "M4", "M5", "M6", "M8"] },
    "B": { "pdf_name": "B.pdf", "specs": ["M3", "M4", "M5", "M6"] }
}

class AdvancedPemPdfViewer(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_folder = "pdf_specs"  
        self.current_doc = None        # 暫存當前開啟的 PDF 物件
        self.current_page_idx = 0      # 當前顯示的頁碼索引
        self.current_pdf_path = ""     # 當前 PDF 路徑
        
        # 後台邏輯虛擬選單（維持與 special_features 的接口相容）
        self.combo_category = QComboBox()
        self.combo_category.addItems(PEM_CONFIG.keys())
        self.combo_spec = QComboBox()
        
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(5)

        # ─── 🚀 頂部綜合控制列：左邊放翻頁，右邊放打開檔案 ───
        top_control_layout = QHBoxLayout()
        top_control_layout.setContentsMargins(5, 5, 5, 2)
        
        # 【左側：手動翻頁按鈕群】
        self.btn_prev = QPushButton("⬅️ 上一頁")
        self.btn_prev.setStyleSheet("font-weight: bold; padding: 5px 12px;")
        self.btn_prev.clicked.connect(self.prev_page)
        
        self.lbl_page_num = QLabel("第 0 / 0 頁")
        self.lbl_page_num.setStyleSheet("font-weight: bold; color: #333; font-size: 13px; margin: 0 10px;")
        
        self.btn_next = QPushButton("下一頁 ➡️")
        self.btn_next.setStyleSheet("font-weight: bold; padding: 5px 12px;")
        self.btn_next.clicked.connect(self.next_page)
        
        top_control_layout.addWidget(self.btn_prev)
        top_control_layout.addWidget(self.lbl_page_num)
        top_control_layout.addWidget(self.btn_next)
        
        # 中間彈性伸縮，把打開檔案的按鈕推到最右邊
        top_control_layout.addStretch()
        
        # 【右側：外部打開檔案按鈕群】
        self.btn_open_current = QPushButton("📄 打開此種類 PDF 檔案")
        self.btn_open_current.setStyleSheet("""
            QPushButton {
                background-color: #E8F5E9; font-weight: bold; color: #2E7D32;
                border: 1px solid #C8E6C9; border-radius: 4px; padding: 6px 12px;
            }
            QPushButton:hover { background-color: #C8E6C9; }
        """)
        self.btn_open_current.clicked.connect(self.open_current_pdf_file)
        self.btn_open_current.setEnabled(False)  
        
        self.btn_open_full = QPushButton("📖 打開 PEM 中文總目錄")
        self.btn_open_full.setStyleSheet("""
            QPushButton {
                background-color: #E0F7FA; font-weight: bold; color: #006064;
                border: 1px solid #B2EBF2; border-radius: 4px; padding: 6px 12px;
            }
            QPushButton:hover { background-color: #B2EBF2; }
        """)
        self.btn_open_full.clicked.connect(self.open_full_catalog)
        
        top_control_layout.addWidget(self.btn_open_current)
        top_control_layout.addWidget(self.btn_open_full)
        
        # 先將這一整排控制列加到最上方
        main_layout.addLayout(top_control_layout)

        # ─── 中間（即現在的下半部）PDF 內嵌滾動區域 ───
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        
        self.lbl_pdf_render = QLabel("請由上方 Excel 選擇規格以載入 PDF...")
        self.lbl_pdf_render.setAlignment(Qt.AlignCenter)
        self.lbl_pdf_render.setStyleSheet("background-color: #616161; border: 1px solid #333; color: white;")
        
        self.scroll_area.setWidget(self.lbl_pdf_render)
        main_layout.addWidget(self.scroll_area)

        # 觸發初始連動
        self.update_spec_options(self.combo_category.currentText())

    def set_category_and_spec(self, excel_category, spec_text):
        """由 special_features.py 呼叫的接口"""
        cleaned_cat = str(excel_category).strip().upper()
        cleaned_spec = str(spec_text).strip().upper()

        target_code = None
        if cleaned_cat.startswith("SO"): target_code = "SO"
        elif cleaned_cat.startswith("S"): target_code = "S"
        elif cleaned_cat.startswith("FH"): target_code = "FH"
        elif cleaned_cat.startswith("B"): target_code = "B"

        if target_code and target_code in PEM_CONFIG:
            self.combo_category.setCurrentText(target_code)
            
            pdf_name = PEM_CONFIG[target_code]["pdf_name"]
            self.current_pdf_path = os.path.join(self.pdf_folder, pdf_name)
            
            self.btn_open_current.setText(f"📄 打開 {pdf_name} 檔案")
            self.btn_open_current.setEnabled(True)
            
            self.execute_auto_search(cleaned_spec)

    def execute_auto_search(self, spec_keyword):
        """核心：自動盲搜定位頁碼"""
        if not os.path.exists(self.current_pdf_path):
            self.lbl_pdf_render.setText(f"❌ 找不到規格書：{os.path.basename(self.current_pdf_path)}")
            self.lbl_page_num.setText("第 0 / 0 頁")
            self.btn_open_current.setEnabled(False)
            return

        try:
            if self.current_doc:
                self.current_doc.close()
            
            self.current_doc = fitz.open(self.current_pdf_path)
            total_pages = len(self.current_doc)
            
            target_idx = 0
            found = False

            candidates = []
            for page_num in range(total_pages):
                page = self.current_doc[page_num]
                page_text = page.get_text()
                
                if any(k in page_text for k in ["目錄", "INDEX", "Index", "索引", "前言"]):
                    continue
                
                instances = page.search_for(spec_keyword)
                if instances:
                    score = len(instances)
                    words = page_text.split()
                    if any(w.strip(":,.-()\"' ") == spec_keyword for w in words):
                        score += 50
                    candidates.append((page_num, score))

            if candidates:
                candidates.sort(key=lambda x: x[1], reverse=True)
                target_idx = candidates[0][0]
                found = True

            self.current_page_idx = target_idx
            self.render_current_page()

        except Exception as e:
            self.lbl_pdf_render.setText(f"💥 PDF 載入失敗:\n{str(e)}")

    def render_current_page(self):
        """根據 self.current_page_idx 渲染畫面"""
        if not self.current_doc:
            return

        try:
            total_pages = len(self.current_doc)
            page = self.current_doc[self.current_page_idx]
            
            self.lbl_page_num.setText(f"第 {self.current_page_idx + 1} / {total_pages} 頁")
            
            zoom = 2.2  
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)

            fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
            q_img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
            pixmap = QPixmap.fromImage(q_img)
            
            target_w = self.scroll_area.width() - 35
            if pixmap.width() > target_w and target_w > 100:
                scaled_pixmap = pixmap.scaledToWidth(target_w, Qt.SmoothTransformation)
                self.lbl_pdf_render.setPixmap(scaled_pixmap)
            else:
                self.lbl_pdf_render.setPixmap(pixmap)
                
            self.lbl_pdf_render.setStyleSheet("background-color: #FFFFFF;")

        except Exception as e:
            self.lbl_pdf_render.setText(f"💥 渲染分頁失敗:\n{str(e)}")

    def prev_page(self):
        if self.current_doc and self.current_page_idx > 0:
            self.current_page_idx -= 1
            self.render_current_page()

    def next_page(self):
        if self.current_doc and self.current_page_idx < (len(self.current_doc) - 1):
            self.current_page_idx += 1
            self.render_current_page()

    def open_current_pdf_file(self):
        if self.current_pdf_path and os.path.exists(self.current_pdf_path):
            os.startfile(self.current_pdf_path)
        else:
            QMessageBox.warning(self, "找不到檔案", "無法開啟當前的規格 PDF 檔案。")

    def open_full_catalog(self):
        catalog_path = os.path.join(self.pdf_folder, "pem的中文總目錄.pdf")
        if os.path.exists(catalog_path):
            os.startfile(catalog_path)
        else:
            QMessageBox.warning(self, "找不到檔案", "找不到【pem的中文總目錄.pdf】")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.current_doc:
            self.render_current_page()
            
    def update_spec_options(self, category_name):
        """維持虛擬相容邏輯"""
        if category_name in PEM_CONFIG:
            self.combo_spec.clear()
            self.combo_spec.addItems(PEM_CONFIG[category_name]["specs"])
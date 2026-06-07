# drawing.py
import math
from PyQt5.QtWidgets import QWidget
from PyQt5.QtGui import QPainter, QPen, QColor, QPainterPath, QFont
from PyQt5.QtCore import Qt, QRectF, QPointF

class DrawWidget(QWidget):
    """專用的繪圖畫布組件（完美緊湊置中 + 標註文字放大版）"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.mode = "none"
        self.draw_data = {}
        
        self.setAutoFillBackground(True)
        p = self.palette()
        p.setColor(self.backgroundRole(), QColor("white"))
        self.setPalette(p)
        self.bend_sides = []
        self.bend_angles = []
        self.bend_r_list = []
        self.thickness = 2.0

    def set_bend_extended_data(self, sides, angles, r_list, thickness):
        """接收主視窗傳來的完整板金製程資料，並引發重繪"""
        self.bend_sides = sides
        self.bend_angles = angles
        self.bend_r_list = r_list
        self.thickness = thickness
        self.update()

    def paintEvent(self, event):
        """🎨 修改版：根據 self.mode 智慧分流繪圖，不再被折彎數據卡死"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)  # 開啟抗鋸齒
        
        # 畫布背景色刷白
        painter.fillRect(self.rect(), QColor("#FFFFFF"))
        
        # 獲取高DPI縮放比例（相容你的 draw_cylinder, draw_cone 參數）
        dpr = self.devicePixelRatioF() if hasattr(self, 'devicePixelRatioF') else 1.0

        # ==========================================
        # 💡 根據當前模式 (self.mode) 進行繪圖分流
        # ==========================================
        
        # ── 模式一：幾何特殊放樣 - 圓柱體 ──
        if self.mode == "cylinder":
            self.draw_cylinder(painter, dpr)
            return
            
        # ── 模式二：幾何特殊放樣 - 圓錐體 ──
        elif self.mode == "cone":
            self.draw_cone(painter, dpr)
            return
            
        # ── 模式三：多道折彎展開圖 ──
        elif self.mode == "bend":
            self.draw_bend(painter, dpr)
            return

        # ── 模式四：原有的單道折彎示意圖（相容舊有預設邏輯） ──
        else:
            # 如果是傳統單道折彎，才去檢查 bend_sides 資料
            if not hasattr(self, 'bend_sides') or len(self.bend_sides) < 2:
                # 若完全沒資料，可以在畫布中央淡淡顯示提示，視覺上更專業
                painter.setPen(QColor("#999999"))
                painter.setFont(QFont("Microsoft JhengHei", 12))
                painter.drawText(self.rect(), Qt.AlignCenter, "等待輸入參數計算放樣圖面...")
                return

            try:
                l1_val = self.bend_sides[0]
                l2_val = self.bend_sides[1]
                angle_val = self.bend_angles[0]
                r_val = self.bend_r_list[0] if len(self.bend_r_list) > 0 else 0.0
                t_val = self.thickness
            except (IndexError, ValueError):
                return

            # [這裡保留你原本 paintEvent 內繪製單道 L1、L2、R1 線段與文字的全部程式碼...]
            canvas_w = self.width()
            canvas_h = self.height()
            cx = int(canvas_w / 2 - r_val + (l1_val / 4))
            cy = int(canvas_h / 2 + r_val - (l2_val / 4))
            
            pen_line = QPen(QColor("#1A237E"), 3)
            painter.setPen(pen_line)

            if r_val <= 0:
                painter.drawLine(int(cx - l1_val), int(cy), int(cx), int(cy))
                painter.drawLine(int(cx), int(cy), int(cx), int(cy - l2_val))
            else:
                path = QPainterPath()
                path.moveTo(int(cx - l1_val), int(cy)) 
                path.lineTo(int(cx - r_val), int(cy))  
                path.arcTo(int(cx - 2 * r_val), int(cy - 2 * r_val), int(2 * r_val), int(2 * r_val), 270, 90)
                path.lineTo(int(cx), int(cy - l2_val))  
                painter.drawPath(path)

            font = painter.font()
            font.setPointSize(11) 
            font.setBold(True)
            painter.setFont(font)

            painter.setPen(QColor("#333333"))
            l1_text = f"L1 = {l1_val}"
            painter.drawText(int(cx - (l1_val / 2) - 35), int(cy + 35), l1_text) 

            l2_text = f"L2 = {l2_val}"
            painter.drawText(int(cx + 30), int(cy - (l2_val / 2) + 5), l2_text)

            if r_val > 0:
                k_factor = 0.5
                outer_angle = 180.0 - angle_val
                neutral_radius = r_val + (k_factor * t_val)
                arc_length = (math.pi * neutral_radius * outer_angle) / 180.0
                
                painter.setPen(QColor("#2E7D32"))
                r_canvas_text = f"R1 = {r_val} (L = {arc_length:.2f})"
                painter.drawText(int(cx - r_val - 110), int(cy - r_val - 20), r_canvas_text)
            else:
                painter.setPen(QColor("#999999")) 
                painter.drawText(int(cx - 50), int(cy - 20), "R0")                

    def set_bend_data(self, side_lengths, angles):
        self.mode = "bend"
        self.draw_data = {"sides": side_lengths, "angles": angles}
        self.update()

    def set_cylinder_data(self, L, H):
        self.mode = "cylinder"
        self.draw_data = {"L": L, "H": H}
        self.update()

    def set_cone_data(self, R, r, theta, H=0.0):
        """接收圓錐展開數據，包含核心的垂直高度 H"""
        self.mode = "cone"
        self.draw_data = {"R": R, "r": r, "theta": theta, "H": H}
        self.update()

    def draw_bend_preview(self, painter):
        """處理畫布示意圖：R角圓弧化、L1在下方、L2靠右移開、R角標註中性線長度"""
        painter.setRenderHint(QPainter.Antialiasing) # 讓圓弧平滑不毛躁
        
        # 1. 讀取介面數值（安全防呆）
        try:
            l1_val = float(self.side_entries[0].text())
            l2_val = float(self.side_entries[1].text())
            angle_val = self.angle_entries[0].value()
            r_val = float(self.r_entries[0].text())
        except (IndexError, ValueError):
            return

        # 讀取板厚 T 與 K-Factor（計算中性線必備）
        try: t_val = float(self.thickness_entry.text())
        except ValueError: t_val = 2.0
        try: k_val = float(self.k_entry.text())
        except ValueError: k_val = 0.4

        # 2. 定義畫布上的基準中心點 (以折彎內側交點為原點)
        cx, cy = 130, 230 
        
        # 畫筆設定 (深藍色線條代表板金)
        pen_line = QPen(QColor("#1A237E"), 3)
        painter.setPen(pen_line)

        # ==========================================
        # 繪製板金線條 (R=0 畫直角，R>0 畫圓弧)
        # ==========================================
        if r_val <= 0:
            # 傳統直角：L1 水平向左，L2 垂直向上
            painter.drawLine(cx - l1_val, cy, cx, cy)
            painter.drawLine(cx, cy, cx, cy - l2_val)
        else:
            # 自然 R 角：使用 Path 接合直線與圓弧
            path = QPainterPath()
            path.moveTo(cx - l1_val, cy) # 起點：L1 左端
            path.lineTo(cx - r_val, cy)  # 水平線到 R 角起點
            # 畫圓弧 (參數：矩形左上X, 左上Y, 寬, 高, 起始角度, 跨越角度)
            path.arcTo(cx - 2 * r_val, cy - 2 * r_val, 2 * r_val, 2 * r_val, 270, 90)
            path.lineTo(cx, cy - l2_val)  # 續接 L2 垂直向上
            painter.drawPath(path)

        # ==========================================
        # 繪製標籤文字 (位置精密微調)
        # ==========================================
        font = painter.font()
        font.setPointSize(10)
        font.setBold(True)
        painter.setFont(font)

        # --- L1 標籤：強迫置於水平線「正下方」 ---
        painter.setPen(QColor("#333333"))
        l1_text = f"L1 = {l1_val}"
        # cy + 25 確保完全在線的下方，不與線重疊
        painter.drawText(cx - (l1_val / 2) - 20, cy + 25, l1_text) 

        # --- L2 標籤：向右移動 2px，並離垂直線開一點 ---
        l2_text = f"L2 = {l2_val}"
        # cx + 15 遠離垂直線，再加 2px 符合你的微調需求
        painter.drawText(cx + 15 + 2, cy - (l2_val / 2) + 5, l2_text)

        # --- R1 標籤：在畫布圓弧旁顯示中性線長度 L=xx.xx ---
        # 計算中性線弧長公式：L = (π * (R + K*T) * 角度) / 180
        neutral_radius = r_val + (k_val * t_val)
        arc_length = (3.14159265 * neutral_radius * angle_val) / 180.0

        if r_val > 0:
            painter.setPen(QColor("#2E7D32")) # 使用綠色突顯 R 角與弧長
            r_canvas_text = f"R1={r_val} (L={arc_length:.2f})"
            # 標註在圓弧的右下斜對角外側
            painter.drawText(cx + 10, cy + 20, r_canvas_text)
        else:
            painter.setPen(QColor("#999999"))
            painter.drawText(cx + 10, cy + 20, "R0")

    def draw_bend(self, painter, dpr):
        sides = self.draw_data.get("sides", [])
        angles = self.draw_data.get("angles", [])
        if not sides: return

        scale = 2.0 / dpr
        points = [QPointF(0.0, 0.0)]
        current_angle = 0.0
        
        for i, length_mm in enumerate(sides):
            length = length_mm * scale
            rad = math.radians(current_angle)
            last_pt = points[-1]
            nx = last_pt.x() + length * math.cos(rad)
            ny = last_pt.y() - length * math.sin(rad)
            points.append(QPointF(nx, ny))
            if i < len(angles):
                current_angle += (180.0 - angles[i])
                
        x_coords = [p.x() for p in points]
        y_coords = [p.y() for p in points]
        min_x, max_x = min(x_coords), max(x_coords)
        min_y, max_y = min(y_coords), max(y_coords)
        
        shape_w = max_x - min_x
        shape_h = max_y - min_y
        
        # 找到原本計算 offset_y 的地方，改成這樣：
        offset_x = (self.width() - shape_w) / 2.0 - min_x
        
        # 💡 原本是 / 2.0，我們把它稍微往上拉（比如改用扣除浮動文字高度後的中心）
        # 讓它整體往上頂，消滅視覺空隙
        offset_y = ((self.height() + 40) - shape_h) / 2.0 - min_y
        
        pen = QPen(QColor("#333333"), 3, Qt.SolidLine)
        painter.setPen(pen)
        
        # ── 💡 放大折彎線段標註文字 (改為 12pt) ──
        painter.setFont(QFont("Microsoft JhengHei", 12, QFont.Bold))
        
        final_points = [QPointF(p.x() + offset_x, p.y() + offset_y) for p in points]
        
        for i in range(len(final_points) - 1):
            p1 = final_points[i]
            p2 = final_points[i+1]
            painter.drawLine(p1, p2)
            
            mid_x = (p1.x() + p2.x()) / 2.0
            mid_y = (p1.y() + p2.y()) / 2.0 - 10 # 稍微拉開避免壓線
            painter.drawText(int(mid_x), int(mid_y), f"L{i+1}")

    def draw_cylinder(self, painter, dpr):
        L = self.draw_data.get("L", 0)
        H = self.draw_data.get("H", 0)
        if L <= 0 or H <= 0: return

        scale = 0.6 / dpr
        rect_w = L * scale
        rect_h = H * scale
        
        # 找到原本計算 y 的地方，改成這樣：
        x = (self.width() - rect_w) / 2.0
        y = ((self.height() + 40) - rect_h) / 2.0 # 💡 加上 40 讓幾何中心微調往上

        painter.setPen(QPen(QColor("#1976D2"), 2, Qt.SolidLine))
        painter.setBrush(QColor("#E3F2FD"))
        
        rect = QRectF(x, y, rect_w, rect_h)
        painter.drawRect(rect)
        
        # ── 💡 放大圓柱體展開文字 (改為 13pt) ──
        painter.setPen(QPen(QColor("black")))
        painter.setFont(QFont("Microsoft JhengHei", 13, QFont.Bold))
        painter.drawText(int(x + rect_w/2 - 60), int(y - 15), f"展開長: {L:.2f}")
        painter.drawText(int(x - 85), int(y + rect_h/2 + 5), f"H: {H:.1f}")

    def draw_cone(self, painter, dpr):
        """圓錐體緊湊置中算法：H 尺寸精準修正為「兩弧線間的垂直線長（斜面尺寸 R - r）」"""
        R = self.draw_data.get("R", 0)
        r = self.draw_data.get("r", 0)
        theta = self.draw_data.get("theta", 0)
        
        if R <= 0: return

        # 🚀 【核心修正】：加工現場要的平面垂直線長（斜面尺寸），直接由展開半徑相減獲得
        H_slant = R - r

        scale = 1.0 / dpr
        max_allowed = min(self.width(), self.height()) * 0.85 
        if R * scale > max_allowed:
            scale = (max_allowed / R) / dpr

        # ── 物理邊界置中計算 ──
        start_ang = 270.0 - (theta / 2.0)
        end_ang = start_ang + theta
        
        p_center = QPointF(0, 0)
        p_left = QPointF(R * scale * math.cos(math.radians(start_ang)), -R * scale * math.sin(math.radians(start_ang)))
        p_right = QPointF(R * scale * math.cos(math.radians(end_ang)), -R * scale * math.sin(math.radians(end_ang)))
        p_top = QPointF(0, -R * scale) if (start_ang <= 90 <= end_ang or start_ang <= 450 <= end_ang) else p_center
        
        xs = [p_center.x(), p_left.x(), p_right.x(), p_top.x()]
        ys = [p_center.y(), p_left.y(), p_right.y(), p_top.y()]
        
        cone_w = max(xs) - min(xs)
        cone_h = max(ys) - min(ys)
        
        cx = (self.width() - cone_w) / 2.0 - min(xs)
        cy = ((self.height() + 40) - cone_h) / 2.0 - min(ys)

        qt_start = int(start_ang * 16)
        qt_span = int(theta * 16)

        # 1. 畫高度輔助中軸線（綠色虛線，改為只畫在內外弧之間，凸顯兩弧垂直距離）
        pen_dash = QPen(QColor("#2E7D32"), 1, Qt.DashLine)
        painter.setPen(pen_dash)
        painter.drawLine(int(cx), int(cy + r * scale), int(cx), int(cy + R * scale))

        # 2. 畫外弧（紅色）
        painter.setPen(QPen(QColor("red"), 2, Qt.SolidLine))
        rect_R = QRectF(cx - R*scale, cy - R*scale, R*2*scale, R*2*scale)
        painter.drawArc(rect_R, qt_start, qt_span)

        # 3. 畫內弧（藍色）
        painter.setPen(QPen(QColor("blue"), 2, Qt.SolidLine))
        rect_r = QRectF(cx - r*scale, cy - r*scale, r*2*scale, r*2*scale)
        painter.drawArc(rect_r, qt_start, qt_span)

        # 4. 畫側邊線（黑色）
        painter.setPen(QPen(QColor("black"), 2, Qt.SolidLine))
        for ang in [start_ang, start_ang + theta]:
            rad = math.radians(ang)
            x1 = cx + r * scale * math.cos(rad)
            y1 = cy - r * scale * math.sin(rad)
            x2 = cx + R * scale * math.cos(rad)
            y2 = cy - R * scale * math.sin(rad)
            painter.drawLine(int(x1), int(y1), int(x2), int(y2))
            
        # 5. 繪製虛擬圓心處的角度標註弧線（紫色）
        ang_r = 30.0
        pen_angle = QPen(QColor("#7B1FA2"), 1, Qt.SolidLine)
        painter.setPen(pen_angle)
        rect_ang = QRectF(cx - ang_r, cy - ang_r, ang_r * 2, ang_r * 2)
        painter.drawArc(rect_ang, qt_start, qt_span)

        # ==========================================================
        # 🎨 全方位尺寸標註文字（文字圖層最上層）
        # ==========================================================
        painter.setFont(QFont("Microsoft JhengHei", 12, QFont.Bold))
        
        # ── 內弧尺寸 r 文字（深藍色，貼在藍色弧線下方） ──
        text_r = f"r = {r:.1f}"
        metrics_r = painter.fontMetrics()
        text_r_w = metrics_r.width(text_r)
        text_r_x = int(cx - (text_r_w / 2))
        text_r_y = int(cy + (r * scale) + 18) 
        painter.setPen(QColor("#0D47A1")) 
        painter.drawText(text_r_x, text_r_y, text_r)

        # ── 外弧尺寸 R 文字（深紅色，貼在紅色弧線下方） ──
        text_R = f"R = {R:.1f}"
        metrics_R = painter.fontMetrics()
        text_R_w = metrics_R.width(text_R)
        text_R_x = int(cx - (text_R_w / 2))
        text_R_y = int(cy + (R * scale) + 22) 
        painter.setPen(QColor("#B71C1C")) 
        painter.drawText(text_R_x, text_R_y, text_R)

        # ── 🚀 修正標註：顯示為「兩弧線的垂直線長（斜面尺寸）」 ──
        if H_slant > 0:
            text_H = f"H = {H_slant:.1f}"
            # 放在內弧與外弧正中間的中軸虛線旁
            mid_y = cy + ((r + R) * scale / 2.0)
            text_H_x = int(cx + 8)
            text_H_y = int(mid_y + 5) 
            painter.setPen(QColor("#2E7D32")) 
            painter.drawText(text_H_x, text_H_y, text_H)

        # ── 展開夾角 θ 文字（深紫色，標註在圓心夾角弧線的右上方） ──
        if theta > 0:
            text_theta = f"θ = {theta:.1f}°"
            text_theta_x = int(cx + 12)
            text_theta_y = int(cy - 8)
            painter.setPen(QColor("#7B1FA2")) 
            painter.drawText(text_theta_x, text_theta_y, text_theta)
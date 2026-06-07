# calculations.py
import math
import os
import pandas as pd

class SheetMetalCalc:
    """專職負責板金展開、幾何放樣的所有數據計算與 Excel 檢索"""
    def __init__(self, excel_file="bend_parameters.xlsx"):
        self.excel_file = excel_file
        self.all_sheets = {}
        self.load_all_excel_sheets()

    def load_all_excel_sheets(self):
        if os.path.exists(self.excel_file):
            try:
                xls = pd.ExcelFile(self.excel_file)
                for name in xls.sheet_names:
                    # 讀取時將 index 與 columns 都轉成字串，避免類型不匹配
                    df = pd.read_excel(xls, sheet_name=name, index_col=0)
                    df.columns = [str(col).strip() for col in df.columns]
                    df.index = [str(idx).strip() for idx in df.index]
                    self.all_sheets[name] = df
            except Exception as e:
                raise RuntimeError(f"讀取 Excel 失敗: {str(e)}")
        else:
            raise FileNotFoundError(f"找不到指定的 Excel 檔案: {self.excel_file}")

    def get_bend_sheets(self):
        """關鍵修改：自動抓取名稱包含 '倍板金係數' 的所有工作表，未來增加 6倍、7倍 都會自動偵測"""
        return [s for s in self.all_sheets.keys() if "倍板金係數" in s]

    def get_sheet_data(self, sheet_name):
        return self.all_sheets.get(sheet_name, pd.DataFrame())

    def get_k90_value(self, sheet, r, c):
        if sheet in self.all_sheets:
            df = self.all_sheets[sheet]
            # 去除前後空白確保比對成功
            r_str, c_str = str(r).strip(), str(c).strip()
            if r_str in df.index and c_str in df.columns:
                try:
                    return df.loc[r_str, c_str]
                except:
                    pass
        return ""

    def calculate_bend_length(self, k90, sides, angles):
        """動態折彎：內邊相加法"""
        sum_l = sum(sides)
        sum_k = sum([(k90 / 90.0) * (180.0 - a) for a in angles])
        return sum_l + sum_k

    def calculate_cylinder(self, d, T, H):
        L = math.pi * (d + T)
        return {"L": L, "H": H}

    def calculate_cone(self, D_in, d_in, H, T):
        if D_in == d_in:
            raise ValueError("大小端內徑不能相同，請選用圓柱體模式。")
        D = D_in + T
        d = d_in + T
        
        slant = math.sqrt(H**2 + ((D - d) / 2.0)**2)
        R_val = (D * slant) / (D - d)
        r_val = R_val - slant
        theta = (D * 180.0) / R_val
        return {"R": R_val, "r": r_val, "theta": theta}
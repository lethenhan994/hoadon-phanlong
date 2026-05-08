import pandas as pd
import re
import os

class ZinZinMisaEngine:
    def __init__(self):
        self.rules_file = 'Mapping_Rules.xlsx'
        self.rules = []
        self.load_rules()

    def load_rules(self):
        """Đọc quy tắc từ file Excel: Từ khóa, Tên chuẩn, TK Nợ, TK Có"""
        if os.path.exists(self.rules_file):
            try:
                df = pd.read_excel(self.rules_file).dropna(subset=['Tu_Khoa', 'Ten_Chuan'])
                # Ưu tiên từ khóa dài để bắt chính xác hơn
                df['len'] = df['Tu_Khoa'].astype(str).str.len()
                df = df.sort_values(by='len', ascending=False)
                for _, row in df.iterrows():
                    self.rules.append({
                        'tk': str(row['Tu_Khoa']).strip().upper(),
                        'tc': str(row['Ten_Chuan']).strip(),
                        'tk_no': str(row.get('TK_No', '')).strip() if pd.notna(row.get('TK_No')) else None,
                        'tk_co': str(row.get('TK_Co', '')).strip() if pd.notna(row.get('TK_Co')) else None
                    })
            except Exception:
                pass

    def is_tax_text(self, text):
        """Nhận diện dòng Thuế GTGT"""
        t = str(text).upper()
        tax_keywords = ["THUẾ GTGT", "THUẾ GIÁ TRỊ GIA TĂNG", "THUẾ VAT", "VAT 8%", "VAT 10%"]
        return any(kw in t for kw in tax_keywords)

    def process_block(self, group):
        """Xử lý phân tích diễn giải và tìm tài khoản cho một cụm chứng từ"""
        found_items = []
        found_tk_no = None
        found_tk_co = None
        plates = []
        
        # Lấy tên khách hàng để xét logic sửa xe ô tô
        ten_kh = str(group.get('TENKH', pd.Series([''])).dropna().iloc[0]).upper() if 'TENKH' in group.columns and not group.get('TENKH').dropna().empty else ""
        
        # Chỉ lấy các dòng hàng (không phải thuế) để phân tích
        goods_rows = group[~group['DIENGIAI'].apply(self.is_tax_text)]
        
        if goods_rows.empty:
            return str(group['DIENGIAI'].iloc[0]), None, None

        for _, row in goods_rows.iterrows():
            row_text = ""
            for col in ['DIENGIAI', 'TENDM', 'MATHANG']:
                if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != "":
                    row_text += str(row[col]).upper() + " "
            
            row_text = row_text.strip()
            
            # 1. Quét biển số xe
            for match in re.finditer(r'XE\s+([\w\d\-]{6,11})|([0-9]{2}[A-Z][\w\d\-]{4,8})', row_text):
                p = match.group(1) if match.group(1) else match.group(2)
                if p and p not in plates: plates.append(p)

            # 2. Dò tìm theo quy tắc trong file Mapping_Rules
            for rule in self.rules:
                if rule['tk'] in row_text:
                    if rule['tc'] not in found_items:
                        found_items.append(rule['tc'])
                        # Ưu tiên lấy tài khoản Nợ/Có từ quy tắc khớp đầu tiên
                        if not found_tk_no: found_tk_no = rule['tk_no']
                        if not found_tk_co: found_tk_co = rule['tk_co']
                    break 

        # 3. Logic đặc biệt cho phí sửa chữa xe
        full_text = " ".join([str(x) for x in goods_rows['DIENGIAI']])
        repair_keywords = ["ẮC QUY", "CUROA", "VÒNG BI", "SKF", "LỐP", "SỬA CHỮA", "BẢO DƯỠNG", "THAY"]
        is_repair = any(k in full_text.upper() for k in repair_keywords)
        is_oto = "Ô TÔ" in ten_kh or "AUTO" in ten_kh

        if plates and (is_repair or is_oto):
            prefix = "Phí sửa ô tô" if is_oto else "Phí sửa xe"
            if prefix not in found_items:
                found_items.insert(0, prefix)

        # 4. Ráp chuỗi diễn giải cuối cùng
        if not found_items:
            final_desc = str(goods_rows['DIENGIAI'].fillna('').iloc[0]).strip()
        else:
            final_desc = ", ".join(found_items)

        # Thêm tiền tố "Phí" nếu chưa có
        prefixes = ["Phí", "Thuế", "Thay", "Lệ phí", "Phụ", "Bình", "Dây", "Lốp", "Vòng", "Cước"]
        if final_desc and not any(final_desc.startswith(p) for p in prefixes):
            final_desc = "Phí " + final_desc
            
        return final_desc, found_tk_no, found_tk_co

def process_data(df_raw):
    """Tạo Bảng 2: Chuẩn hóa diễn giải và áp luật tài khoản Phan Long"""
    engine = ZinZinMisaEngine()
    df = df_raw.copy()
    df.columns = df.columns.astype(str).str.strip()
    
    if 'SOCT' not in df.columns: return df
    
    # Gom nhóm theo số chứng từ
    df['SOCT_Clean'] = df['SOCT'].astype(str).str.strip().replace(['nan', ''], pd.NA).ffill()
            
    summary_results = {}
    for soct, group in df.groupby('SOCT_Clean'):
        summary_results[soct] = engine.process_block(group)
        
    def apply_logic(row):
        # Nếu là dòng thuế -> Giữ nguyên diễn giải và tài khoản gốc
        if engine.is_tax_text(row['DIENGIAI']):
            return row['DIENGIAI'], row.get('TKNO'), row.get('TKCO')
        
        # Lấy kết quả phân tích
        res_desc, res_no, res_co = summary_results.get(row['SOCT_Clean'], (row['DIENGIAI'], None, None))
        
        # Xử lý TK Nợ
        final_no = res_no if res_no else row.get('TKNO')
        
        # Xử lý TK Có theo luật số tiền và Thuê tài chính
        try:
            # Ưu tiên lấy tổng tiền thanh toán (TTVND_TT) để xét ngưỡng 5 triệu
            val_tt = str(row.get('TTVND_TT', 0)).replace(',', '').replace(' ', '')
            amount = float(val_tt) if val_tt != "" else 0
        except:
            amount = 0

        if res_co == "341":
            final_co = "341"
        elif amount >= 5000000:
            final_co = "331"
        else:
            final_co = "1111"

        return res_desc, final_no, final_co

    # Cập nhật dữ liệu
    results = df.apply(apply_logic, axis=1)
    df['DIENGIAI'] = [x[0] for x in results]
    df['TKNO'] = [x[1] for x in results]
    df['TKCO'] = [x[2] for x in results]
    
    return df.drop(columns=['SOCT_Clean'], errors='ignore')

def generate_table_3(df_b2):
    """Tạo Bảng 3: Gom các dòng trùng Số HD và Diễn giải, cộng dồn số tiền"""
    df_t3 = df_b2.copy()
    
    # Gom theo các cột định danh
    group_cols = ['SOCT', 'SO_HD', 'DIENGIAI', 'TKNO', 'TKCO']
    for col in group_cols:
        if col in df_t3.columns:
            df_t3[col] = df_t3[col].fillna("")

    # Danh sách cột tiền và lượng cần cộng dồn
    num_cols = ['LUONG_CTU', 'LUONG', 'TTVND', 'THUEVND', 'TTVND_TT', 'TIENHANG']
    for col in num_cols:
        if col in df_t3.columns:
            df_t3[col] = pd.to_numeric(df_t3[col].astype(str).str.replace(',', '').str.replace(' ', ''), errors='coerce').fillna(0)
            
    agg_funcs = {col: ('sum' if col in num_cols else 'first') for col in df_t3.columns if col not in group_cols}
    
    df_grouped = df_t3.groupby(group_cols, as_index=False, dropna=False).agg(agg_funcs)
    
    # Định dạng lại số tiền cho đẹp (bỏ đuôi .0)
    for col in num_cols:
        if col in df_grouped.columns:
            df_grouped[col] = df_grouped[col].apply(lambda x: f"{x:g}" if x != 0 else "")
            
    return df_grouped[df_b2.columns]
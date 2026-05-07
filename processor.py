import pandas as pd
import re
import os

class ZinZinMisaEngine:
    def __init__(self):
        self.rules_file = 'Mapping_Rules.xlsx'
        self.rules = []
        self.load_rules()

    def load_rules(self):
        if os.path.exists(self.rules_file):
            try:
                df = pd.read_excel(self.rules_file).dropna(subset=['Tu_Khoa', 'Ten_Chuan'])
                df['len'] = df['Tu_Khoa'].astype(str).str.len()
                df = df.sort_values(by='len', ascending=False)
                for _, row in df.iterrows():
                    self.rules.append((str(row['Tu_Khoa']).strip().upper(), str(row['Ten_Chuan']).strip()))
            except Exception:
                pass

    def is_tax_text(self, text):
        t = str(text).upper()
        tax_keywords = ["THUẾ GTGT", "THUẾ GIÁ TRỊ GIA TĂNG", "THUẾ VAT", "VAT 8%", "VAT 10%"]
        return any(kw in t for kw in tax_keywords)

    def process_block(self, group):
        found_items = []
        full_text_list = []
        plates = []
        
        ten_kh = str(group.get('TENKH', pd.Series([''])).dropna().iloc[0]).upper() if 'TENKH' in group.columns and not group.get('TENKH').dropna().empty else ""
        goods_rows = group[~group['DIENGIAI'].apply(self.is_tax_text)]
        
        if goods_rows.empty:
            return str(group['DIENGIAI'].iloc[0])

        for _, row in goods_rows.iterrows():
            row_text = ""
            for col in ['DIENGIAI', 'TENDM', 'MATHANG']:
                if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != "":
                    row_text += str(row[col]).upper() + " "
                    
            row_text = row_text.strip()
            full_text_list.append(row_text)
            
            for match in re.finditer(r'XE\s+([\w\d\-]{6,11})|([0-9]{2}[A-Z][\w\d\-]{4,8})', row_text):
                p = match.group(1) if match.group(1) else match.group(2)
                if p and p not in plates: plates.append(p)

            toll_keywords = ["TRẠM THU PHÍ", "VÉ TRẠM", "VÉ QUA TRẠM", "VÉ CẦU ĐƯỜNG", "VÉ LƯỢT", "VETC", "EPASS", "SỬ DỤNG ĐB", "SỬ DỤNG ĐƯỜNG BỘ", "CƯỚC ĐƯỜNG BỘ", "CƯỚC ĐB", "PHÍ ĐƯỜNG BỘ"]
            if any(kw in row_text for kw in toll_keywords):
                return "Cước đường bộ"

            for tk, tc in self.rules:
                if tk in row_text:
                    if tc not in found_items:
                        found_items.append(tc)
                    break 

        full_text = " ".join(full_text_list)

        repair_keywords = ["ẮC QUY", "CUROA", "VÒNG BI", "SKF", "LỐP", "SỬA CHỮA", "BẢO DƯỠNG", "THAY"]
        is_repair = any(k in full_text for k in repair_keywords)
        is_oto = "Ô TÔ" in ten_kh or "AUTO" in ten_kh

        if plates and (is_repair or is_oto):
            prefix = "Phí sửa ô tô" if is_oto else "Phí sửa xe"
            if prefix not in found_items:
                found_items.insert(0, prefix)

        if not found_items:
            main_desc = str(goods_rows['DIENGIAI'].fillna('').iloc[0]).strip()
        else:
            main_desc = ", ".join(found_items)

        final_desc = main_desc.strip()
        final_desc = re.sub(r'(container|cont)\s+cont\b', r'\1', final_desc, flags=re.IGNORECASE)
        final_desc = final_desc.replace("Phí Phí", "Phí")
        
        prefixes = ["Phí", "Thuế", "Thay", "Lệ phí", "Phụ", "Bình", "Dây", "Lốp", "Vòng", "Cước"]
        if not any(final_desc.startswith(p) for p in prefixes) and final_desc:
            final_desc = "Phí " + final_desc
            
        return final_desc

def process_data(df_raw):
    engine = ZinZinMisaEngine()
    df = df_raw.copy()
    
    df.columns = df.columns.astype(str).str.strip()
    
    if 'SOCT' not in df.columns:
        return df
        
    df['SOCT_Clean'] = df['SOCT'].astype(str).str.strip().replace(['nan', ''], pd.NA).ffill()
            
    summary_map = {}
    for soct, group in df.groupby('SOCT_Clean'):
        summary_map[soct] = engine.process_block(group)
        
    def finalize_desc(row):
        if engine.is_tax_text(row['DIENGIAI']):
            return row['DIENGIAI']
        return summary_map.get(row['SOCT_Clean'], row['DIENGIAI'])

    df['DIENGIAI'] = df.apply(finalize_desc, axis=1)
    df = df.drop(columns=['SOCT_Clean'], errors='ignore')
    
    return df

# =======================================================
# BƯỚC 3: THUẬT TOÁN GOM NHÓM & CỘNG TIỀN (MỚI)
# =======================================================
def generate_table_3(df_b2):
    df_t3 = df_b2.copy()
    
    # 1. Trám rỗng các cột làm gốc để không bị lỗi rớt dòng
    group_cols = ['SOCT', 'SO_HD', 'DIENGIAI']
    for col in group_cols:
        if col in df_t3.columns:
            df_t3[col] = df_t3[col].fillna("")

    # 2. Danh sách các cột CẦN TÍNH TỔNG (Số lượng, Số tiền)
    num_cols = ['LUONG_CTU', 'LUONG', 'TTVND', 'THUEVND', 'TTVND_TT', 'TIENHANG']
    
    # Xóa dấu phẩy của file Excel và ép thành số thực (float) để máy tính cộng được
    for col in num_cols:
        if col in df_t3.columns:
            df_t3[col] = pd.to_numeric(df_t3[col].astype(str).str.replace(',', '').str.replace(' ', ''), errors='coerce').fillna(0)
            
    # 3. Tạo luật Gom Nhóm
    agg_funcs = {}
    for col in df_t3.columns:
        if col not in group_cols:
            if col in num_cols:
                agg_funcs[col] = 'sum'   # Cột Tiền/Số lượng thì CỘNG LẠI
            else:
                agg_funcs[col] = 'first' # Các cột khác (Mã KH, Ngày, TK Nợ/Có) thì LẤY GIÁ TRỊ DÒNG ĐẦU TIÊN
                
    # 4. THỰC THI GOM NHÓM 
    # Gom theo Số Hóa đơn & Diễn Giải (Kèm SOCT để bảo vệ mã bút toán)
    df_grouped = df_t3.groupby(group_cols, as_index=False, dropna=False).agg(agg_funcs)
    
    # Ép thứ tự cột trở lại y hệt form chuẩn 146 cột ban đầu
    df_grouped = df_grouped[df_b2.columns]
    
    # 5. Định dạng lại cột tiền (Bỏ số .0 ở cuối cho kế toán dễ nhìn)
    for col in num_cols:
        if col in df_grouped.columns:
            df_grouped[col] = df_grouped[col].apply(lambda x: f"{x:g}" if x != 0 else "")
            
    return df_grouped
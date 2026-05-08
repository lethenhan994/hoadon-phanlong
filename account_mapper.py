import pandas as pd
import os

def apply_account_rules(df_b3):
    """
    Hàm chuyên dụng cho Phan Long:
    1. Đọc TK Nợ/Có từ Mapping_Rules.xlsx.
    2. Ưu tiên TK Nợ của nhóm 'Dịch vụ cảng' khi có ghép phí.
    3. Tự động điền TK Có: 341 (ưu tiên), >= 5tr là 331, < 5tr là 1111.
    4. Tự động điền LCTG (Cột đầu tiên): 1111 -> PC, 331/341 -> PKT.
    """
    df_res = df_b3.copy()
    
    # 1. Đọc file Mapping_Rules.xlsx để lấy quy tắc tài khoản
    tk_rules = {}
    if os.path.exists('Mapping_Rules.xlsx'):
        try:
            df_tk = pd.read_excel('Mapping_Rules.xlsx')
            cols = {str(c).strip().lower(): c for c in df_tk.columns}
            col_no = cols.get('tkno') or cols.get('tk_no')
            col_co = cols.get('tkco') or cols.get('tk_co')
            
            if 'Ten_Chuan' in df_tk.columns:
                df_tk = df_tk.dropna(subset=['Ten_Chuan'])
                for _, row in df_tk.iterrows():
                    dg = str(row['Ten_Chuan']).strip().upper()
                    # Xử lý lấy số tài khoản (bỏ phần .0 nếu có)
                    t_no = str(row[col_no]).split('.')[0] if col_no and pd.notna(row[col_no]) else None
                    t_co = str(row[col_co]).split('.')[0] if col_co and pd.notna(row[col_co]) else None
                    tk_rules[dg] = {'no': t_no, 'co': t_co}
        except: 
            pass
            
    # Bộ từ khóa nhận diện nhóm "Dịch vụ cảng" của Phan Long
    port_keywords = ["NÂNG", "HẠ", "CONTAINER", "CONT", "BÃI", "KIỂM HÓA", "CÂN XE", "VỆ SINH", "GIAO", "CỔNG"]
            
    # 2. Logic xử lý từng dòng dữ liệu
    def process_row_logic(row):
        dg = str(row.get('DIENGIAI', '')).strip().upper()
        
        # Lấy giá trị hiện tại làm mặc định
        tk_no_final = row.get('TKNO', '')
        tk_co_final = row.get('TKCO', '')
        lctg_final = ''

        # --- A. XỬ LÝ TK NỢ ---
        # Tìm các quy tắc khớp trong diễn giải (có thể ghép nhiều phí)
        matched_rules = {k: v for k, v in tk_rules.items() if k in dg}
        tk_co_rule_from_excel = None

        if matched_rules:
            # ƯU TIÊN: Nếu có phí thuộc nhóm Dịch vụ cảng, bốc TK Nợ của nó
            priority_key = next((k for k in matched_rules if any(pw in k for pw in port_keywords)), None)
            chosen_rule = matched_rules[priority_key] if priority_key else list(matched_rules.values())[0]
            
            if chosen_rule.get('no') and chosen_rule.get('no') != 'None':
                tk_no_final = chosen_rule['no']
            
            # Kiểm tra xem trong các quy tắc khớp có ông nào yêu cầu TK Có là 341 không
            if any(r.get('co') == '341' for r in matched_rules.values()):
                tk_co_rule_from_excel = '341'

        # --- B. XỬ LÝ TK CÓ (Theo luật 5 triệu) ---
        try:
            amt_text = str(row.get('TTVND_TT', 0)).replace(',', '').replace(' ', '')
            amt = float(amt_text) if amt_text != "" else 0
        except: 
            amt = 0
            
        if tk_co_rule_from_excel == '341':
            tk_co_final = '341'
        elif amt >= 5000000:
            tk_co_final = '331'
        else:
            tk_co_final = '1111'

        # --- C. XỬ LÝ LCTG (CỘT ĐẦU TIÊN) ---
        # Quy tắc: 1111 -> PC, 331/341 -> PKT
        if tk_co_final == '1111':
            lctg_final = 'PC'
        elif tk_co_final in ['331', '341']:
            lctg_final = 'PKT'
            
        # Riêng dòng Thuế: Giữ nguyên diễn giải gốc và thường là PKT/331
        if "THUẾ GTGT" in dg or "VAT" in dg:
            return 'PKT', row.get('TKNO', ''), '331'
            
        return lctg_final, tk_no_final, tk_co_final

    # Áp dụng logic và ghi đè vào bảng dữ liệu
    results = df_res.apply(process_row_logic, axis=1)
    
    # Đảm bảo các cột này tồn tại để ghi dữ liệu
    df_res['LCTG'] = [x[0] for x in results]
    df_res['TKNO'] = [x[1] for x in results]
    df_res['TKCO'] = [x[2] for x in results]

    return df_res
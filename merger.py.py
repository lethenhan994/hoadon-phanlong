import pandas as pd

def create_final_table(df_table_1_header, df_table_2_lines):
    """
    Ghép nối Bảng 1 (Thông tin chung) và Bảng 2 (Chi tiết) thành Bảng 3 (Kết quả cuối cùng)
    - df_table_1_header: Dataframe chứa thông tin Ngày, KH, MST... (duy nhất theo SOCT)
    - df_table_2_lines: Dataframe chi tiết (đã chạy qua processor.py để có TENDM chuẩn)
    """
    # 1. Đảm bảo SOCT ở cả 2 bảng đều là chuỗi (string) để ghép không bị lệch
    df_table_1_header['SOCT'] = df_table_1_header['SOCT'].astype(str).str.strip()
    df_table_2_lines['SOCT'] = df_table_2_lines['SOCT'].astype(str).str.strip()

    # Xóa các dòng trùng lặp ở Bảng 1 (chỉ giữ lại 1 dòng Header cho mỗi SOCT)
    df_table_1_header = df_table_1_header.drop_duplicates(subset=['SOCT'], keep='first')

    # 2. Ghép nối (Left Join): Giữ nguyên mọi dòng của Bảng 2, kéo thông tin Bảng 1 đắp vào
    df_merged = pd.merge(df_table_2_lines, df_table_1_header, on='SOCT', how='left', suffixes=('', '_header'))

    # 3. Định nghĩa bản đồ cột (Từ tên gốc hệ thống -> Tên cột hiển thị Bảng 3)
    # Lưu ý: Sắp xếp đúng theo thứ tự bạn muốn hiển thị trên Excel cuối cùng
    mapping_columns = {
        'SOCT': 'Số CT',
        'NGAYCT': 'Ngày CT',
        'SO_HD': 'Số HĐ',
        'NGAY_HD': 'Ngày HĐ',
        'TENDM': 'Diễn giải',
        'TKNO': 'TK Nợ',
        'TKCO': 'TK Có',
        'TTVND': 'Số tiền',
        'THUEVND': 'Thuế',
        'TTVND_TT': 'Tổng thanh toán',
        'TENKH': 'Tên KH',
        'DIACHI': 'Địa chỉ',
        'MS_DN': 'MST'
    }

    # 4. Kiểm tra xem các cột cần thiết có tồn tại không, nếu không có thì tạo cột trống
    for col in mapping_columns.keys():
        if col not in df_merged.columns:
            df_merged[col] = ""

    # 5. Đổi tên cột và sắp xếp thứ tự
    df_final = df_merged[list(mapping_columns.keys())].rename(columns=mapping_columns)

    # 6. Làm sạch dữ liệu (Xóa chữ 'nan' hoặc 'NaN' do ghép nối sinh ra)
    df_final = df_final.fillna("")
    df_final = df_final.replace(['nan', 'NaN', 'None'], "")

    return df_final

# Hướng dẫn cách sử dụng file này (Đoạn này để test hoặc gọi từ file main)
if __name__ == "__main__":
    print("Đây là Module Merger. Hàm create_final_table đã sẵn sàng được gọi!")
    
    # Ví dụ cách bạn sẽ gọi nó trong luồng chạy chính:
    # import processor
    # import merger
    
    # df_raw = pd.read_excel("dulieu.xlsx")
    # df_table_1 = ... (Tách lấy các cột header)
    
    # BƯỚC 1: Gọi file processor (KHÔNG sửa đổi processor)
    # df_table_2_processed = processor.process_data(df_raw)
    
    # BƯỚC 2: Gọi file merger để xuất Bảng 3
    # df_bang_3 = merger.create_final_table(df_table_1, df_table_2_processed)
    # df_bang_3.to_excel("Ket_Qua_Cuoi_Cung.xlsx", index=False)
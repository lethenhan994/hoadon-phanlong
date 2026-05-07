import streamlit as st
import pandas as pd
import io
import processor

# Cấu hình giao diện
st.set_page_config(page_title="Nhập Hoá Đơn Phan Long", layout="wide")
st.title("🏢 NHẬP HOÁ ĐƠN PHAN LONG")
st.markdown("---")

# DANH SÁCH 146 CỘT CHUẨN (Giữ nguyên như cũ)
EXPECTED_COLUMNS = [
    "LCTG", "SR_HD", "SO_HD", "NGAY_HD", "SOCT", "SO_PC", "NGAYCT", "DIENGIAI", "TKNO", "MADTPNNO", 
    "MAYTCPNO", "MADMNO", "LO_NHAP", "TKCO", "MADTPNCO", "MAYTCPCO", "MADMCO", "LO_XUAT", "MA_CT", 
    "MADTGT", "TENDM", "DONVI_CTU", "DONVI", "LUONG_CTU", "LUONG", "DGVND", "TTVND", "PT_CK", 
    "CHIETKHAU", "HDVAT", "TKTHUE", "TS_GTGT", "THUEVND", "TTVND_TT", "MATHANG", "MAKH", "TENKH", 
    "MS_DN", "DIACHI", "DIACHI_NGD", "KHACHHANG", "GHICHU", "NV_BAN", "DGVON", "GTVON", "LAI_GOP", 
    "TYGIA", "TTUSD", "THUEUSD", "TTUSD_TT", "SOCTGS", "NGAYCTGS", "STT_SC", "DT_NHAN", "DT_XUAT", 
    "THANG", "SO_HOPDONG", "MAUSER", "MATTTU", "NGAY_DAO_HAN", "MADVTT", "TENDVTT", "TKNHNO", 
    "TENTKNHNO", "MADVNT", "TENDVNT", "TKNHCO", "TENTKNHCO", "TENVUNG", "KYHIEU", "MANGD", 
    "DIENGIAI2", "DGUSD", "TIENHANG", "TENHH_IN", "HTTT", "INVOI", "MASANPHAM", "DONTRONG", 
    "MA_NV_BAN", "SOCT_U", "COL_11", "COL_12", "COL_13", "DANHDAU", "TRANGTHAI", "CHIETKHAU_USD", 
    "DG_GC", "DG_VC", "HANSUDUNG", "ID_CHUNGTU", "KHM_HD", "KXLDG", "LENHSX", "MA_HD", "MA_HH_GC", 
    "MA_TIEP_THI", "MA_TT", "MA_VUNG", "MANHOM", "MANHOM1", "MANHOM2", "MODEL", "MS_TM", 
    "NGAY_HOPDONG_SC", "NGAY_TT", "NGAYLO", "NHACUNGCAP", "NHASANXUAT", "NHOMKH", "NOIDUNG_SC", 
    "SL_GC", "SO_PT", "SOKHEUOC", "STT_TT", "TEN_TIEP_THI", "TENCT_SC", "TENYTCPNO", "THANG_N", 
    "THUEEUR", "TK_CHIETKHAU", "TK_XUATKHO", "TNK_USD", "TNK_VND", "TS_NK", "TT_GC", "TT_VC", 
    "TTEUR", "DONVI_1", "DONVI_2", "HSQD_DVT", "LUONG_1", "LUONG_2", "MABPSX", "STT_BT", "TENDTGT", 
    "IMEI", "MADONHANG", "BARCODE", "MADONHANG_MUA", "MANHOMDT1", "TENNHOMDT1", "CHIETKHAU2", 
    "ID_UPDATE", "PT_CK2", "ID_NGHIEPVU", "MACHINHANH", "STT_SAPXEP", "DGMAUSAC", "GUID", "KYDUYET", 
    "MAMAUSAC", "SOCT_N", "TENMAUSAC", "THOIGIANNHAP", "TTMAUSAC", "SO_PC_N", "SO_PT_N"
]

uploaded_file = st.file_uploader("Bước 1: Chọn file dữ liệu MISA (Excel hoặc CSV)", type=['csv', 'xlsx'])

if uploaded_file is not None:
    if 'file_name' not in st.session_state or st.session_state['file_name'] != uploaded_file.name:
        try:
            with st.spinner('Đang nạp và căn chỉnh Form chuẩn Phan Long...'):
                uploaded_file.seek(0)
                if uploaded_file.name.endswith('.csv'):
                    df_raw = pd.read_csv(uploaded_file, dtype=str)
                else:
                    df_raw = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
                
                df_raw.columns = df_raw.columns.astype(str).str.strip()
                df_raw = df_raw.reindex(columns=EXPECTED_COLUMNS).fillna("")
                
                st.session_state['df_raw'] = df_raw
                st.session_state['file_name'] = uploaded_file.name
                st.session_state['processed'] = False
                st.session_state['df_b2'] = None
                st.session_state['df_b3'] = None
                # Trạng thái nút tải
                st.session_state['download_done'] = False
                
        except Exception as e:
            st.error(f"Lỗi khi đọc file: {e}")

    if 'df_raw' in st.session_state and st.session_state['df_raw'] is not None:
        total_rows = len(st.session_state['df_raw'])
        st.subheader(f"📋 BẢNG 1: DỮ LIỆU THÔ ({total_rows} dòng)")
        st.dataframe(st.session_state['df_raw'].head(100), height=300)
        st.markdown("---")
            
        if st.button("🚀 Xử lý Diễn giải & Gom nhóm (Tạo Bảng 2 & 3)"):
            with st.spinner('Hệ thống đang chuẩn hóa và gom nhóm tự động...'):
                try:
                    df_b2 = processor.process_data(st.session_state['df_raw'].copy())
                    st.session_state['df_b2'] = df_b2
                    df_b3 = processor.generate_table_3(df_b2.copy())
                    st.session_state['df_b3'] = df_b3
                    st.session_state['processed'] = True
                except Exception as err:
                    st.error(f"Lỗi hệ thống: {err}")

        if st.session_state.get('processed'):
            st.subheader(f"✨ BẢNG 2: KẾT QUẢ DIỄN GIẢI CHI TIẾT ({len(st.session_state['df_b2'])} dòng)")
            st.dataframe(st.session_state['df_b2'], height=400)
            
            st.subheader(f"🔥 BẢNG 3: DỮ LIỆU ĐÃ GOM NHÓM ({len(st.session_state['df_b3'])} dòng)")
            st.dataframe(st.session_state['df_b3'], height=400)

            # Xuất dữ liệu Bảng 3 ra bộ nhớ đệm
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                st.session_state['df_b3'].to_excel(writer, index=False, sheet_name="Bang_3_Gom_Nhom")
            
            # Logic đổi tên nút sau khi bấm
            btn_label = "⬇️ TẢI BẢNG 3 VỀ MÁY"
            if st.session_state.get('download_done'):
                btn_label = "✅ ĐÃ TẢI XONG RỒI NHÉ NINH XINH ĐẸP"

            # Nút tải
            if st.download_button(
                label=btn_label,
                data=output.getvalue(),
                file_name="HoaDon_PhanLong_Bang3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_btn_ninh"
            ):
                st.session_state['download_done'] = True
                # Force rerun để cập nhật label nút ngay lập tức
                st.rerun()
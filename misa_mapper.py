import pandas as pd

def fill_table_3(df_table_2):
    """
    Điền kết quả từ Bảng 2 vào cấu trúc 157 cột MISA
    """
    misa_columns = [
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
        "ID_UPDATE", "PT_CK2", "ID_NGHIEPVU", "MACHINHANH", "STT_SAPXEP", "GUID", "THOIGIANNHAP",
        "KYDUYET", "DGMAUSAC", "MAMAUSAC", "SOCT_N", "TENMAUSAC", "TTMAUSAC"
    ]
    df_3 = pd.DataFrame(index=df_table_2.index, columns=misa_columns)
    for col in df_table_2.columns:
        if col in df_3.columns:
            df_3[col] = df_table_2[col]
    if 'TENDM' in df_table_2.columns:
        df_3['DIENGIAI'] = df_table_2['TENDM']
    return df_3
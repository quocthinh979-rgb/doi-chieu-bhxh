import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import io
import re
from datetime import datetime

# ===== CẤU HÌNH TRANG =====
st.set_page_config(
    page_title="Đối chiếu BHXH - BHYT - BHTN",
    page_icon="📊",
    layout="wide"
)

# ===== CSS =====
st.markdown("""
<style>
    .header-box {
        background: linear-gradient(90deg, #000080 0%, #1a1a9e 100%);
        padding: 15px 20px; border-radius: 8px; color: white; margin-bottom: 20px;
    }
    .header-box h2 { color: white; margin: 0; font-size: 22px; }
    .header-box p { color: #e0e0ff; margin: 5px 0 0 0; font-size: 14px; }
    .metric-box {
        background: white; padding: 15px; border-radius: 8px;
        border: 1px solid #e0e0e0; box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .metric-title { font-size: 14px; color: #666; margin-bottom: 5px; }
    .metric-value { font-size: 28px; font-weight: bold; color: #000080; }
    .row-item {
        display: flex; justify-content: space-between;
        padding: 8px 0; border-bottom: 1px dashed #eee;
    }
    .row-item span:last-child { font-weight: bold; color: #000080; }
    .warning-box {
        background: #fff9e6; border-left: 4px solid #ffc107;
        padding: 12px 15px; border-radius: 4px;
        font-size: 13px; color: #664d03;
    }
</style>
""", unsafe_allow_html=True)

# ===== ĐỌC SECRETS =====
def get_secret(key, default=None):
    try:
        return st.secrets[key]
    except:
        return default

# ===== HÀM ĐỌC FILE CƠ BẢN =====
def doc_file_excel(file):
    """Đọc Excel, hỗ trợ cả .xls và .xlsx"""
    try:
        ten_file = file.name.lower()
        if ten_file.endswith('.xls'):
            # File Excel cũ cần xlrd
            return pd.read_excel(file, header=None, engine='xlrd')
        else:
            return pd.read_excel(file, header=None)
    except Exception as e:
        st.error(f"Lỗi đọc Excel: {e}")
        return None

def doc_file_pdf_text(file):
    try:
        text = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    except Exception as e:
        st.error(f"Lỗi đọc PDF text: {e}")
        return None

def doc_file_pdf_scan(file):
    try:
        images = convert_from_path(file)
        text = ""
        for img in images:
            text += pytesseract.image_to_string(img, lang='vie+eng')
        return text
    except Exception as e:
        st.error(f"Lỗi OCR PDF scan: {e}")
        return None

def doc_file_pdf_tu_dong(file):
    try:
        with pdfplumber.open(file) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text() or ""
            if len(text.strip()) < 50:
                st.info("📷 Phát hiện PDF scan, đang dùng OCR...")
                return doc_file_pdf_scan(file)
            else:
                st.info("📄 Phát hiện PDF text, đang đọc trực tiếp...")
                return doc_file_pdf_text(file)
    except Exception as e:
        st.error(f"Lỗi đọc PDF: {e}")
        return None

def doc_file_word(file):
    try:
        from docx import Document
        doc = Document(file)
        text = ""
        for para in doc.paragraphs:
            text += para.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + "\t"
                text += "\n"
        return text
    except Exception as e:
        st.error(f"Lỗi đọc Word: {e}")
        return None

def doc_file_bat_ky(file):
    if file is None:
        return None
    ten_file = file.name.lower()
    if ten_file.endswith(('.xlsx', '.xls')):
        return doc_file_excel(file)
    elif ten_file.endswith('.pdf'):
        return doc_file_pdf_tu_dong(file)
    elif ten_file.endswith('.docx'):
        return doc_file_word(file)
    elif ten_file.endswith('.csv'):
        return pd.read_csv(file)
    else:
        st.warning(f"Định dạng không hỗ trợ: {ten_file}")
        return None

# ===== HÀM XỬ LÝ HEADER 2 DÒNG =====
def gop_header_2_dong(df_raw, dong_header_1, dong_header_2):
    """Gộp 2 dòng header bị tách do merge cell trong Excel"""
    header = []
    for col in range(df_raw.shape[1]):
        h1 = df_raw.iloc[dong_header_1, col]
        h2 = df_raw.iloc[dong_header_2, col]
        
        # Lấy giá trị không NaN
        h1_str = str(h1).strip() if pd.notna(h1) else ""
        h2_str = str(h2).strip() if pd.notna(h2) else ""
        
        # Gộp: nếu h1 có và h2 có → "h1 h2", nếu chỉ 1 trong 2 có → lấy cái đó
        if h1_str and h2_str and h1_str != h2_str:
            header.append(f"{h1_str} {h2_str}".strip())
        elif h1_str:
            header.append(h1_str)
        elif h2_str:
            header.append(h2_str)
        else:
            header.append(f"Unnamed_{col}")
    
    return header

# ===== HÀM TÌM DÒNG HEADER =====
def tim_dong_header(df_raw, tu_khoa_list):
    """Tìm dòng header chứa các từ khóa"""
    for i in range(min(20, len(df_raw))):
        row_str = " ".join([str(x) for x in df_raw.iloc[i].values if pd.notna(x)]).lower()
        if any(tk.lower() in row_str for tk in tu_khoa_list):
            return i
    return None

# ===== ĐỌC BẢNG LƯƠNG =====
def doc_bang_luong(file):
    """Đọc bảng lương, trích xuất Họ tên và Tổng tiền lương còn được lĩnh"""
    df_raw = doc_file_excel(file)
    if df_raw is None:
        return None
    
    # Tìm dòng header (dòng có chứa "STT" và "Họ" hoặc "Mã")
    dong_header_1 = tim_dong_header(df_raw, ["STT", "Họ và tên", "Mã NV"])
    
    if dong_header_1 is None:
        st.error("❌ Không tìm thấy dòng header trong bảng lương")
        return None
    
    # Kiểm tra xem có header 2 dòng không (dòng tiếp theo có chứa "Tăng ca", "Phụ cấp"...)
    dong_header_2 = None
    if dong_header_1 + 1 < len(df_raw):
        row_tiep = " ".join([str(x) for x in df_raw.iloc[dong_header_1 + 1].values if pd.notna(x)]).lower()
        if any(tk in row_tiep for tk in ["tăng ca", "phụ cấp", "thường", "chủ nhật", "giờ"]):
            dong_header_2 = dong_header_1 + 1
    
    # Gộp header
    if dong_header_2 is not None:
        header_moi = gop_header_2_dong(df_raw, dong_header_1, dong_header_2)
        du_lieu_bat_dau = dong_header_2 + 1
    else:
        header_moi = [str(x).strip() if pd.notna(x) else f"Unnamed_{i}" 
                      for i, x in enumerate(df_raw.iloc[dong_header_1].values)]
        du_lieu_bat_dau = dong_header_1 + 1
    
    # Tạo DataFrame mới
    df = df_raw.iloc[du_lieu_bat_dau:].copy()
    df.columns = header_moi
    df = df.reset_index(drop=True)
    
    st.write("🔍 **Cột tìm thấy trong bảng lương:**")
    st.write(list(df.columns))
    
    # Tìm cột Họ tên
    col_hoten = None
    for c in df.columns:
        c_lower = str(c).lower()
        if "họ và tên" in c_lower or "họ tên" in c_lower or c_lower == "ho va ten":
            col_hoten = c
            break
    
    if col_hoten is None:
        # Thử tìm cột "Mã NV" để lấy tên
        for c in df.columns:
            if "mã nv" in str(c).lower() or "mã nhân viên" in str(c).lower():
                col_hoten = c
                break
    
    # Tìm cột Lương (ưu tiên "Tổng tiền lương còn được lĩnh")
    col_luong = None
    for c in df.columns:
        c_lower = str(c).lower()
        if "tổng tiền lương còn được lĩnh" in c_lower:
            col_luong = c
            break
    
    # Nếu không có, tìm "Tổng thu nhập tính thuế"
    if col_luong is None:
        for c in df.columns:
            if "tổng thu nhập tính thuế" in str(c).lower():
                col_luong = c
                break
    
    # Nếu vẫn không có, tìm "Tổng lương"
    if col_luong is None:
        for c in df.columns:
            if "tổng lương" in str(c).lower():
                col_luong = c
                break
    
    if col_hoten is None:
        st.error(f"❌ Không tìm thấy cột 'Họ và tên'. Cột hiện có: {list(df.columns)}")
        return None
    
    if col_luong is None:
        st.error(f"❌ Không tìm thấy cột lương. Cột hiện có: {list(df.columns)}")
        return None
    
    st.success(f"✅ Tìm thấy cột: **{col_hoten}** và **{col_luong}**")
    
    # Lọc dữ liệu
    ket_qua = df[[col_hoten, col_luong]].copy()
    ket_qua.columns = ["Họ tên", "Lương trả thực tế"]
    ket_qua = ket_qua.dropna(subset=["Họ tên"])
    ket_qua["Họ tên"] = ket_qua["Họ tên"].astype(str).str.strip()
    ket_qua["Lương trả thực tế"] = pd.to_numeric(ket_qua["Lương trả thực tế"], errors='coerce').fillna(0)
    
    # Loại bỏ dòng tổng cộng
    ket_qua = ket_qua[~ket_qua["Họ tên"].str.lower().str.contains("tổng|cộng", na=False)]
    ket_qua = ket_qua[ket_qua["Họ tên"] != ""]
    
    return ket_qua.reset_index(drop=True)

# ===== ĐỌC D02-TS =====
def doc_d02(file):
    """Đọc D02-TS, trích xuất Họ tên và Tiền lương tiền công"""
    df_raw = doc_file_excel(file)
    if df_raw is None:
        return None
    
    # Tìm dòng header
    dong_header = tim_dong_header(df_raw, ["STT", "Họ và tên", "Họ tên", "Mã số BHXH"])
    
    if dong_header is None:
        st.error("❌ Không tìm thấy dòng header trong D02")
        return None
    
    # Kiểm tra header 2 dòng
    dong_header_2 = None
    if dong_header + 1 < len(df_raw):
        row_tiep = " ".join([str(x) for x in df_raw.iloc[dong_header + 1].values if pd.notna(x)]).lower()
        if any(tk in row_tiep for tk in ["tiền lương", "tiền công", "phụ cấp", "bổ sung"]):
            dong_header_2 = dong_header + 1
    
    if dong_header_2 is not None:
        header_moi = gop_header_2_dong(df_raw, dong_header, dong_header_2)
        du_lieu_bat_dau = dong_header_2 + 1
    else:
        header_moi = [str(x).strip() if pd.notna(x) else f"Unnamed_{i}" 
                      for i, x in enumerate(df_raw.iloc[dong_header].values)]
        du_lieu_bat_dau = dong_header + 1
    
    df = df_raw.iloc[du_lieu_bat_dau:].copy()
    df.columns = header_moi
    df = df.reset_index(drop=True)
    
    st.write("🔍 **Cột tìm thấy trong D02:**")
    st.write(list(df.columns))
    
    # Tìm cột Họ tên
    col_hoten = None
    for c in df.columns:
        c_lower = str(c).lower()
        if "họ và tên" in c_lower or "họ tên" in c_lower:
            col_hoten = c
            break
    
    # Tìm cột Tiền lương tiền công
    col_luong = None
    for c in df.columns:
        c_lower = str(c).lower()
        if "tiền lương tiền công" in c_lower or "tiền lương đóng" in c_lower:
            col_luong = c
            break
    
    # Nếu không có, tìm cột "Tiền lương"
    if col_luong is None:
        for c in df.columns:
            if "tiền lương" in str(c).lower() and "tiền công" not in str(c).lower():
                col_luong = c
                break
    
    # Tìm cột Mã số BHXH
    col_bhxh = None
    for c in df.columns:
        if "mã số bhxh" in str(c).lower() or "số bhxh" in str(c).lower():
            col_bhxh = c
            break
    
    # Tìm cột Ngày sinh
    col_ngaysinh = None
    for c in df.columns:
        if "ngày sinh" in str(c).lower():
            col_ngaysinh = c
            break
    
    if col_hoten is None:
        st.error(f"❌ Không tìm thấy cột 'Họ và tên' trong D02. Cột hiện có: {list(df.columns)}")
        return None
    
    if col_luong is None:
        st.error(f"❌ Không tìm thấy cột lương trong D02. Cột hiện có: {list(df.columns)}")
        return None
    
    st.success(f"✅ Tìm thấy cột: **{col_hoten}** và **{col_luong}**")
    
    # Lọc dữ liệu
    cols = [col_hoten]
    if col_bhxh:
        cols.append(col_bhxh)
    if col_ngaysinh:
        cols.append(col_ngaysinh)
    cols.append(col_luong)
    
    ket_qua = df[cols].copy()
    rename_map = {col_hoten: "Họ tên", col_luong: "Lương đóng D02"}
    if col_bhxh:
        rename_map[col_bhxh] = "Mã số BHXH"
    if col_ngaysinh:
        rename_map[col_ngaysinh] = "Ngày sinh"
    ket_qua = ket_qua.rename(columns=rename_map)
    
    ket_qua = ket_qua.dropna(subset=["Họ tên"])
    ket_qua["Họ tên"] = ket_qua["Họ tên"].astype(str).str.strip()
    ket_qua["Lương đóng D02"] = pd.to_numeric(ket_qua["Lương đóng D02"], errors='coerce').fillna(0)
    
    # Loại bỏ dòng tổng cộng
    ket_qua = ket_qua[~ket_qua["Họ tên"].str.lower().str.contains("tổng|cộng", na=False)]
    ket_qua = ket_qua[ket_qua["Họ tên"] != ""]
    
    return ket_qua.reset_index(drop=True)

# ===== ĐỐI CHIẾU =====
def chuan_hoa_ten(ten):
    """Chuẩn hóa tên: bỏ dấu, viết thường, bỏ khoảng trắng thừa"""
    ten = str(ten).strip().lower()
    ten = re.sub(r'\s+', ' ', ten)
    # Bỏ dấu tiếng Việt
    ten = re.sub(r'[àáảãạăằắẳẵặâầấẩẫậ]', 'a', ten)
    ten = re.sub(r'[èéẻẽẹêềếểễệ]', 'e', ten)
    ten = re.sub(r'[ìíỉĩị]', 'i', ten)
    ten = re.sub(r'[òóỏõọôồốổỗộơờớởỡợ]', 'o', ten)
    ten = re.sub(r'[ùúủũụưừứửữự]', 'u', ten)
    ten = re.sub(r'[ỳýỷỹỵ]', 'y', ten)
    ten = re.sub(r'[đ]', 'd', ten)
    return ten.strip()

def doi_chieu(bang_luong, d02):
    """Đối chiếu 2 bảng"""
    bang_luong = bang_luong.copy()
    d02 = d02.copy()
    
    bang_luong["ten_chuan"] = bang_luong["Họ tên"].apply(chuan_hoa_ten)
    d02["ten_chuan"] = d02["Họ tên"].apply(chuan_hoa_ten)
    
    merged = pd.merge(
        bang_luong,
        d02,
        on="ten_chuan",
        how="outer",
        suffixes=("_BL", "_D02")
    )
    
    ket_qua = []
    stt = 0
    
    for _, row in merged.iterrows():
        ho_ten_bl = row.get("Họ tên_BL")
        ho_ten_d02 = row.get("Họ tên_D02")
        ho_ten = ho_ten_bl if pd.notna(ho_ten_bl) else ho_ten_d02
        
        luong_thuc_te = row.get("Lương trả thực tế", 0)
        luong_thuc_te = 0 if pd.isna(luong_thuc_te) else luong_thuc_te
        
        luong_d02 = row.get("Lương đóng D02", 0)
        luong_d02 = 0 if pd.isna(luong_d02) else luong_d02
        
        ma_bhxh = row.get("Mã số BHXH", "")
        ma_bhxh = "" if pd.isna(ma_bhxh) else str(ma_bhxh)
        
        ngay_sinh = row.get("Ngày sinh", "")
        ngay_sinh = "" if pd.isna(ngay_sinh) else str(ngay_sinh)
        
        # Xác định loại
        if pd.isna(row.get("Lương đóng D02")):
            loai = "Truy đóng"
            chenh_lech = luong_thuc_te
            khoan = "Chưa có tên trên D02"
            dien_giai = "Chưa tham gia BHXH - Cần đăng ký theo Điều 31 Luật BHXH 2024"
        elif luong_thuc_te > luong_d02:
            loai = "Truy thu"
            chenh_lech = luong_thuc_te - luong_d02
            khoan = "Chênh lệch tiền lương đóng BHXH"
            dien_giai = "Điều 31 Luật BHXH 2024 - Tiền lương làm căn cứ đóng BHXH"
        else:
            continue
        
        stt += 1
        ket_qua.append({
            "STT": stt,
            "Họ tên": ho_ten,
            "Mã số BHXH": ma_bhxh,
            "Ngày sinh": ngay_sinh,
            "Loại": loai,
            "Lương đóng D02": luong_d02,
            "Lương trả thực tế": luong_thuc_te,
            "Chênh lệch": chenh_lech,
            "Khoản truy thu": khoan,
            "Số tháng": 12,
            "Số tiền truy thu": chenh_lech * 0.32,
            "Diễn giải pháp lý": dien_giai
        })
    
    return pd.DataFrame(ket_qua)

# ===== HEADER =====
st.markdown("""
<div class="header-box">
    <h2>📊 CÔNG CỤ ĐỐI CHIẾU MỨC ĐÓNG BHXH, BHYT, BHTN</h2>
    <p>Hệ thống quản lý tự động - Căn cứ Luật BHXH 2024 (Luật số 41/2024/QH15)</p>
</div>
""", unsafe_allow_html=True)

# ===== CHỌN ĐƠN VỊ =====
danh_sach_don_vi = get_secret("danh_sach_don_vi", ["[TF0372F] CÔNG TY TNHH K&D GARMENT"])

col_a, col_b = st.columns([4, 1])
with col_a:
    don_vi = st.selectbox("🏢 Đơn vị đang chọn:", danh_sach_don_vi)
with col_b:
    st.write("")
    st.write("")
    if st.button("📁 Mở Lưu Trữ", use_container_width=True):
        st.info("Chức năng đang phát triển")

# ===== MENU TABS =====
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Đối chiếu",
    "📂 Nhập dữ liệu",
    "📜 Danh sách đơn vị",
    "💼 Mức lương tối thiểu",
    "📚 Thư viện pháp luật",
    "⚙️ Quản trị"
])

# ============================================================
# TAB 1: ĐỐI CHIẾU
# ============================================================
with tab1:
    st.markdown("### 📉 Mức Đóng & Tỷ Lệ Trích Nộp Bảo Hiểm Bắt Buộc")
    st.caption("Áp dụng cho HĐLĐ từ 01 tháng trở lên | Mức tham chiếu 2.340.000 VNĐ | Lương tối thiểu Vùng II 4.410.000 VNĐ")

    st.markdown("**Tổng trích nộp: 32%**")
    st.markdown("---")

    col_emp, col_er = st.columns(2)

    with col_emp:
        st.markdown("""
        <div class="metric-box">
            <div class="metric-title">👤 Người Lao Động (NLĐ) Trích Nộp</div>
            <div class="metric-value" style="color:#1976d2;">10.5%</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        st.markdown("""
        <div class="row-item"><span>🟢 BHXH (Hưu trí & Tử tuất)</span><span>8.0%</span></div>
        <div class="row-item"><span>❤️ BHYT (Bảo hiểm Y tế)</span><span>1.5%</span></div>
        <div class="row-item"><span>🟦 BHTN (Bảo hiểm Thất nghiệp)</span><span>1.0%</span></div>
        """, unsafe_allow_html=True)

    with col_er:
        st.markdown("""
        <div class="metric-box">
            <div class="metric-title">🏢 Người Sử Dụng Lao Động (NSDLĐ) Đóng</div>
            <div class="metric-value" style="color:#388e3c;">21.5%</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        st.markdown("""
        <div class="row-item"><span>🟢 BHXH (14% + 3%)</span><span>17.0%</span></div>
        <div class="row-item"><span>❤️ BHYT</span><span>3.0%</span></div>
        <div class="row-item"><span>🟦 BHTN</span><span>1.0%</span></div>
        <div class="row-item"><span>⚠️ BHTN-BNN</span><span>0.5%*</span></div>
        """, unsafe_allow_html=True)

    st.write("")
    st.markdown("""
    <div class="warning-box">
    <b>⚠️ Lưu ý:</b> Mức trần BHXH, BHYT: 46.800.000 VNĐ/tháng | Mức trần BHTN Vùng II: 88.200.000 VNĐ/tháng
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("📋 Kết quả đối chiếu - Danh sách truy thu & truy đóng")

    if "df_ket_qua" in st.session_state and st.session_state.df_ket_qua is not None:
        df = st.session_state.df_ket_qua
        
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            st.metric("Tổng lao động", len(df))
        with col_s2:
            so_truy_thu = len(df[df["Loại"] == "Truy thu"])
            st.metric("🔴 Truy thu", so_truy_thu)
        with col_s3:
            so_truy_dong = len(df[df["Loại"] == "Truy đóng"])
            st.metric("🟠 Truy đóng", so_truy_dong)
        with col_s4:
            tong = df["Số tiền truy thu"].sum()
            st.metric("Tổng tiền", f"{tong:,.0f} VNĐ")
        
        st.dataframe(
            df,
            use_container_width=True,
            column_config={
                "Lương đóng D02": st.column_config.NumberColumn("Lương đóng D02", format="%,.0f"),
                "Lương trả thực tế": st.column_config.NumberColumn("Lương trả thực tế", format="%,.0f"),
                "Chênh lệch": st.column_config.NumberColumn("Chênh lệch", format="%,.0f"),
                "Số tiền truy thu": st.column_config.NumberColumn("Số tiền truy thu", format="%,.0f")
            }
        )
        
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 Tải kết quả (CSV)",
            csv,
            f"ket_qua_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv"
        )
        
        if st.button("🗑️ Xóa kết quả này"):
            st.session_state.df_ket_qua = None
            st.rerun()
    else:
        st.info("Chưa có dữ liệu. Vui lòng nhập ở tab 'Nhập dữ liệu'.")

# ============================================================
# TAB 2: NHẬP DỮ LIỆU
# ============================================================
with tab2:
    st.header("📂 Nhập dữ liệu")
    st.caption("Hỗ trợ: Excel (.xlsx, .xls), CSV, Word (.docx), PDF (text + scan)")

    col_up1, col_up2 = st.columns(2)

    with col_up1:
        st.subheader("1️⃣ Bảng lương")
        file_luong = st.file_uploader(
            "Chọn file bảng lương",
            type=["xlsx", "xls", "csv", "pdf", "docx"],
            key="file_luong"
        )
        if file_luong:
            st.success(f"✅ Đã tải: {file_luong.name}")

    with col_up2:
        st.subheader("2️⃣ Bảng chấm công")
        file_cong = st.file_uploader(
            "Chọn file bảng chấm công",
            type=["xlsx", "xls", "csv", "pdf", "docx"],
            key="file_cong"
        )
        if file_cong:
            st.success(f"✅ Đã tải: {file_cong.name}")

    col_up3, col_up4 = st.columns(2)

    with col_up3:
        st.subheader("3️⃣ Mẫu D02-TS")
        file_d02 = st.file_uploader(
            "Chọn file D02-TS",
            type=["xlsx", "xls", "csv", "docx", "pdf"],
            key="file_d02"
        )
        if file_d02:
            st.success(f"✅ Đã tải: {file_d02.name}")

    with col_up4:
        st.subheader("4️⃣ Quy chế chi trả lương")
        file_quyche = st.file_uploader(
            "Chọn file quy chế chi trả lương",
            type=["xlsx", "xls", "csv", "docx", "pdf"],
            key="file_quyche"
        )
        if file_quyche:
            st.success(f"✅ Đã tải: {file_quyche.name}")

    st.markdown("---")
    
    if st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU", type="primary", use_container_width=True):
        if file_luong and file_d02:
            with st.spinner("Đang xử lý..."):
                try:
                    st.markdown("### 📖 Đang đọc bảng lương...")
                    bang_luong = doc_bang_luong(file_luong)
                    
                    st.markdown("### 📖 Đang đọc D02-TS...")
                    d02 = doc_d02(file_d02)
                    
                    if bang_luong is None or d02 is None:
                        st.error("❌ Không đọc được dữ liệu. Kiểm tra lại cấu trúc file.")
                    else:
                        st.success(f"✅ Bảng lương: {len(bang_luong)} lao động")
                        st.success(f"✅ D02-TS: {len(d02)} lao động")
                        
                        with st.expander("👁️ Xem dữ liệu đã đọc"):
                            st.write("**Bảng lương:**")
                            st.dataframe(bang_luong.head(10), use_container_width=True)
                            st.write("**D02-TS:**")
                            st.dataframe(d02.head(10), use_container_width=True)
                        
                        ket_qua = doi_chieu(bang_luong, d02)
                        
                        if len(ket_qua) > 0:
                            st.session_state.df_ket_qua = ket_qua
                            st.success(f"✅ Đối chiếu xong! Tìm thấy {len(ket_qua)} lao động cần xử lý.")
                            st.balloons()
                        else:
                            st.info("✅ Không có lao động nào cần truy thu/truy đóng.")
                except Exception as e:
                    st.error(f"Lỗi xử lý: {e}")
                    st.exception(e)
        else:
            st.warning("⚠️ Vui lòng tải lên ít nhất file Bảng lương và D02-TS.")

# ============================================================
# TAB 3-6 (giữ nguyên như bản trước)
# ============================================================
with tab3:
    st.header("📜 Danh sách đơn vị đã đối chiếu")
    if "danh_sach_da_doi_chieu" not in st.session_state:
        st.session_state.danh_sach_da_doi_chieu = []
    if st.session_state.danh_sach_da_doi_chieu:
        st.dataframe(pd.DataFrame(st.session_state.danh_sach_da_doi_chieu), use_container_width=True)
        if st.button("🗑️ Xóa toàn bộ danh sách"):
            st.session_state.danh_sach_da_doi_chieu = []
            st.rerun()
    else:
        st.info("Chưa có đơn vị nào được đối chiếu trong phiên làm việc này.")

with tab4:
    st.header("💼 Mức lương tối thiểu vùng")
    st.caption("Căn cứ Nghị định 74/2024/NĐ-CP")
    col_v1, col_v2, col_v3, col_v4 = st.columns(4)
    with col_v1:
        st.number_input("Vùng I (đ/tháng)", value=4960000, step=100000)
    with col_v2:
        st.number_input("Vùng II (đ/tháng)", value=4410000, step=100000)
    with col_v3:
        st.number_input("Vùng III (đ/tháng)", value=3860000, step=100000)
    with col_v4:
        st.number_input("Vùng IV (đ/tháng)", value=3450000, step=100000)

with tab5:
    st.header("📚 Thư viện căn cứ pháp luật")
    with st.expander("📖 Luật BHXH 2024 (Luật số 41/2024/QH15)"):
        st.markdown("""
        - **Điều 31**: Tiền lương làm căn cứ đóng BHXH bắt buộc
        - **Điều 32**: Tỷ lệ đóng BHXH
        """)
    with st.expander("📖 Nghị định 74/2024/NĐ-CP"):
        st.markdown("- Quy định mức lương tối thiểu vùng")

with tab6:
    st.header("⚙️ Quản trị hệ thống")
    mk = st.text_input("Nhập mật khẩu admin:", type="password")
    mat_khau_admin = get_secret("mat_khau_admin", "admin123")
    
    if mk == mat_khau_admin:
        st.success("✅ Đăng nhập thành công")
        st.markdown("---")
        st.subheader("🗑️ Xóa dữ liệu")
        st.warning("⚠️ Thao tác này không thể hoàn tác!")
        
        col_x1, col_x2, col_x3 = st.columns(3)
        with col_x1:
            if st.button("🗑️ Xóa kết quả đối chiếu"):
                st.session_state.df_ket_qua = None
                st.success("Đã xóa!")
                st.rerun()
        with col_x2:
            if st.button("🗑️ Xóa danh sách đơn vị"):
                st.session_state.danh_sach_da_doi_chieu = []
                st.success("Đã xóa!")
                st.rerun()
        with col_x3:
            if st.button("🗑️ XÓA TOÀN BỘ"):
                st.session_state.df_ket_qua = None
                st.session_state.danh_sach_da_doi_chieu = []
                st.success("Đã xóa toàn bộ!")
                st.rerun()
    elif mk:
        st.error("❌ Sai mật khẩu")

st.markdown("---")
st.caption("📌 Công cụ hỗ trợ đối chiếu tự động dựa trên Luật BHXH 2024.")

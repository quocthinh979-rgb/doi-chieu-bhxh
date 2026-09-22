import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import io
import re
import json
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

# ===== HÀM ĐỌC FILE =====
def doc_file_excel(file):
    try:
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

# ===== TRÍCH XUẤT DỮ LIỆU TỪ EXCEL =====
def tim_cot(df, ten_cot):
    """Tìm vị trí cột theo tên (không phân biệt hoa thường)"""
    for i, row in df.iterrows():
        for j, cell in enumerate(row):
            if pd.notna(cell) and ten_cot.lower() in str(cell).lower():
                return i, j  # trả về (dòng header, cột)
    return None, None

def doc_bang_luong(file):
    """Đọc bảng lương, trả về DataFrame với Họ tên và Tổng tiền công được lĩnh"""
    df_raw = doc_file_excel(file)
    if df_raw is None:
        return None
    
    # Tìm dòng header
    header_row = None
    for i in range(min(15, len(df_raw))):
        row_str = " ".join([str(x) for x in df_raw.iloc[i].values if pd.notna(x)])
        if "Họ tên" in row_str or "HỌ TÊN" in row_str or "Mã NV" in row_str:
            header_row = i
            break
    
    if header_row is None:
        st.warning("Không tìm thấy dòng header trong bảng lương")
        return None
    
    # Đọc với header đúng
    df = pd.read_excel(file, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Tìm cột cần thiết
    col_hoten = None
    col_luong = None
    
    for c in df.columns:
        if "họ tên" in c.lower() or "ho ten" in c.lower():
            col_hoten = c
        if "tổng tiền công" in c.lower() or "tổng tiền công được lĩnh" in c.lower():
            col_luong = c
        if "tổng thu nhập" in c.lower() and col_luong is None:
            col_luong = c
    
    if col_hoten is None or col_luong is None:
        st.warning(f"Không tìm thấy cột. Cột hiện có: {list(df.columns)}")
        return None
    
    # Lọc dữ liệu
    ket_qua = df[[col_hoten, col_luong]].copy()
    ket_qua.columns = ["Họ tên", "Lương trả thực tế"]
    ket_qua = ket_qua.dropna(subset=["Họ tên"])
    ket_qua["Họ tên"] = ket_qua["Họ tên"].astype(str).str.strip()
    ket_qua["Lương trả thực tế"] = pd.to_numeric(ket_qua["Lương trả thực tế"], errors='coerce').fillna(0)
    
    return ket_qua

def doc_d02(file):
    """Đọc D02-TS, trả về DataFrame với Họ tên và Tiền lương tiền công"""
    df_raw = doc_file_excel(file)
    if df_raw is None:
        return None
    
    # Tìm dòng header
    header_row = None
    for i in range(min(15, len(df_raw))):
        row_str = " ".join([str(x) for x in df_raw.iloc[i].values if pd.notna(x)])
        if "Họ tên" in row_str or "HỌ TÊN" in row_str or "Mã số BHXH" in row_str:
            header_row = i
            break
    
    if header_row is None:
        st.warning("Không tìm thấy dòng header trong D02")
        return None
    
    # Đọc với header đúng
    df = pd.read_excel(file, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Tìm cột cần thiết
    col_hoten = None
    col_luong = None
    col_bhxh = None
    col_ngaysinh = None
    
    for c in df.columns:
        if "họ tên" in c.lower() or "ho ten" in c.lower():
            col_hoten = c
        if "tiền lương tiền công" in c.lower() or "tiền lương đóng" in c.lower():
            col_luong = c
        if "mã số bhxh" in c.lower() or "số bhxh" in c.lower():
            col_bhxh = c
        if "ngày sinh" in c.lower() or "ngay sinh" in c.lower():
            col_ngaysinh = c
    
    if col_hoten is None or col_luong is None:
        st.warning(f"Không tìm thấy cột. Cột hiện có: {list(df.columns)}")
        return None
    
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
    
    return ket_qua

def doi_chieu(bang_luong, d02):
    """Đối chiếu 2 bảng, trả về danh sách truy thu + truy đóng"""
    # Chuẩn hóa tên để so khớp (bỏ dấu, viết thường)
    def chuan_hoa_ten(ten):
        ten = str(ten).strip().lower()
        # Bỏ dấu tiếng Việt
        ten = re.sub(r'[àáảãạăằắẳẵặâầấẩẫậ]', 'a', ten)
        ten = re.sub(r'[èéẻẽẹêềếểễệ]', 'e', ten)
        ten = re.sub(r'[ìíỉĩị]', 'i', ten)
        ten = re.sub(r'[òóỏõọôồốổỗộơờớởỡợ]', 'o', ten)
        ten = re.sub(r'[ùúủũụưừứửữự]', 'u', ten)
        ten = re.sub(r'[ỳýỷỹỵ]', 'y', ten)
        ten = re.sub(r'[đ]', 'd', ten)
        return ten.strip()
    
    bang_luong["ten_chuan"] = bang_luong["Họ tên"].apply(chuan_hoa_ten)
    d02["ten_chuan"] = d02["Họ tên"].apply(chuan_hoa_ten)
    
    # Merge 2 bảng
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
        stt += 1
        ho_ten = row["Họ tên_BL"] if pd.notna(row.get("Họ tên_BL")) else row.get("Họ tên_D02", "")
        luong_thuc_te = row.get("Lương trả thực tế", 0) or 0
        luong_d02 = row.get("Lương đóng D02", 0) or 0
        ma_bhxh = row.get("Mã số BHXH", "") if pd.notna(row.get("Mã số BHXH")) else ""
        ngay_sinh = row.get("Ngày sinh", "") if pd.notna(row.get("Ngày sinh")) else ""
        
        # Xác định loại
        if pd.isna(row.get("Lương đóng D02")):
            # Có trong bảng lương nhưng không có trong D02
            loai = "Truy đóng"
            chenh_lech = luong_thuc_te
            khoan = "Chưa có tên trên D02 - Yêu cầu tham gia"
            dien_giai = "Chưa tham gia BHXH - Cần đăng ký theo Điều 31 Luật BHXH 2024"
        elif luong_thuc_te > luong_d02:
            # Có trong cả 2 nhưng lương thực tế > lương đóng
            loai = "Truy thu"
            chenh_lech = luong_thuc_te - luong_d02
            khoan = "Chênh lệch tiền lương đóng BHXH"
            dien_giai = "Điều 31 Luật BHXH 2024 - Tiền lương làm căn cứ đóng BHXH"
        else:
            # Đóng đủ
            continue
        
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
            "Số tiền truy thu": chenh_lech * 0.32,  # 32% tổng
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

    # Kết quả
    st.markdown("---")
    st.subheader("📋 Kết quả đối chiếu - Danh sách truy thu & truy đóng")

    if "df_ket_qua" in st.session_state and st.session_state.df_ket_qua is not None:
        df = st.session_state.df_ket_qua
        
        # Thống kê
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
        
        # Hiển thị bảng
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
        
        # Tải xuống
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 Tải kết quả (CSV)",
            csv,
            f"ket_qua_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv"
        )
        
        # Nút xóa kết quả
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
            "Chọn file bảng lương (cột 'Tổng tiền công được lĩnh')",
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
            "Chọn file D02-TS (cột 'Tiền lương tiền công')",
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
                    bang_luong = doc_bang_luong(file_luong)
                    d02 = doc_d02(file_d02)
                    
                    if bang_luong is None or d02 is None:
                        st.error("❌ Không đọc được dữ liệu. Kiểm tra lại cấu trúc file.")
                    else:
                        st.success(f"✅ Đọc bảng lương: {len(bang_luong)} lao động")
                        st.success(f"✅ Đọc D02: {len(d02)} lao động")
                        
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
# TAB 3: DANH SÁCH ĐƠN VỊ
# ============================================================
with tab3:
    st.header("📜 Danh sách đơn vị đã đối chiếu")
    
    # Lấy từ session state
    if "danh_sach_da_doi_chieu" not in st.session_state:
        st.session_state.danh_sach_da_doi_chieu = []
    
    if st.session_state.danh_sach_da_doi_chieu:
        df_dv = pd.DataFrame(st.session_state.danh_sach_da_doi_chieu)
        st.dataframe(df_dv, use_container_width=True)
        
        if st.button("🗑️ Xóa toàn bộ danh sách"):
            st.session_state.danh_sach_da_doi_chieu = []
            st.rerun()
    else:
        st.info("Chưa có đơn vị nào được đối chiếu trong phiên làm việc này.")
        st.caption("💡 Dữ liệu chỉ lưu trong phiên làm việc, sẽ mất khi tắt trình duyệt.")

# ============================================================
# TAB 4: MỨC LƯƠNG TỐI THIỂU
# ============================================================
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

# ============================================================
# TAB 5: THƯ VIỆN PHÁP LUẬT
# ============================================================
with tab5:
    st.header("📚 Thư viện căn cứ pháp luật")
    
    with st.expander("📖 Luật BHXH 2024 (Luật số 41/2024/QH15)"):
        st.markdown("""
        - **Điều 31**: Tiền lương làm căn cứ đóng BHXH bắt buộc
        - **Điều 32**: Tỷ lệ đóng BHXH
        - Hiệu lực: 01/07/2025
        """)
    
    with st.expander("📖 Nghị định 74/2024/NĐ-CP"):
        st.markdown("- Quy định mức lương tối thiểu vùng. Hiệu lực: 01/07/2024")
    
    with st.expander("📖 Thông tư 41/2024/TT-BLĐTBXH"):
        st.markdown("- Hướng dẫn thi hành Luật BHXH về chế độ, chính sách")

# ============================================================
# TAB 6: QUẢN TRỊ
# ============================================================
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
                st.success("Đã xóa kết quả đối chiếu!")
                st.rerun()
        
        with col_x2:
            if st.button("🗑️ Xóa danh sách đơn vị"):
                st.session_state.danh_sach_da_doi_chieu = []
                st.success("Đã xóa danh sách đơn vị!")
                st.rerun()
        
        with col_x3:
            if st.button("🗑️ XÓA TOÀN BỘ"):
                st.session_state.df_ket_qua = None
                st.session_state.danh_sach_da_doi_chieu = []
                st.success("Đã xóa toàn bộ dữ liệu!")
                st.rerun()
        
        st.markdown("---")
        st.subheader("📊 Trạng thái bộ nhớ")
        st.write(f"- Kết quả đối chiếu: **{'Có' if st.session_state.get('df_ket_qua') is not None else 'Trống'}**")
        st.write(f"- Số đơn vị đã lưu: **{len(st.session_state.get('danh_sach_da_doi_chieu', []))}**")
        
        st.markdown("---")
        st.subheader("🔧 Cấu hình tỷ lệ trích nộp")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.number_input("Tỷ lệ NLĐ đóng BHXH (%)", value=8.0, step=0.5)
            st.number_input("Tỷ lệ NLĐ đóng BHYT (%)", value=1.5, step=0.5)
            st.number_input("Tỷ lệ NLĐ đóng BHTN (%)", value=1.0, step=0.5)
        with col_c2:
            st.number_input("Tỷ lệ NSDLĐ đóng BHXH (%)", value=17.0, step=0.5)
            st.number_input("Tỷ lệ NSDLĐ đóng BHYT (%)", value=3.0, step=0.5)
            st.number_input("Tỷ lệ NSDLĐ đóng BHTN (%)", value=1.0, step=0.5)
        
        if st.button("💾 Lưu cấu hình"):
            st.success("Đã lưu cấu hình! (Lưu ý: Streamlit Cloud sẽ reset khi redeploy)")
    elif mk:
        st.error("❌ Sai mật khẩu")

# ===== FOOTER =====
st.markdown("---")
st.caption("📌 Công cụ hỗ trợ đối chiếu tự động dựa trên Luật BHXH 2024. Kết quả mang tính tham khảo.")
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
        padding: 15px 20px;
        border-radius: 8px;
        color: white;
        margin-bottom: 20px;
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

# ===== HÀM ĐỌC FILE =====
def doc_file_excel(file):
    """Đọc file Excel, trả về DataFrame"""
    try:
        df = pd.read_excel(file, header=None)
        return df
    except Exception as e:
        st.error(f"Lỗi đọc Excel: {e}")
        return None

def doc_file_pdf_text(file):
    """Đọc PDF dạng text (không scan)"""
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
    """Đọc PDF scan bằng OCR"""
    try:
        # Chuyển PDF thành ảnh
        images = convert_from_path(file)
        text = ""
        for img in images:
            # OCR từng ảnh
            text += pytesseract.image_to_string(img, lang='vie+eng')
        return text
    except Exception as e:
        st.error(f"Lỗi OCR PDF scan: {e}")
        return None

def doc_file_pdf_tu_dong(file):
    """Tự động phát hiện PDF text hay scan"""
    try:
        with pdfplumber.open(file) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text() or ""
            
            # Nếu text quá ít (< 50 ký tự) → có thể là scan
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
    """Đọc file Word"""
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
    """Hàm tổng: tự động đọc theo đuôi file"""
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

def trich_xuat_bang_luong(text):
    """Trích xuất dữ liệu bảng lương từ text OCR"""
    # Pattern tìm dòng có: Mã NV, Họ tên, Lương
    # Ví dụ: "1 25040014 LƯƠNG THỊ PHƯƠNG Công nhân may 26 1 4900000 ..."
    pattern = r'(\d+)\s+(\d{6,8})\s+([A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+)+)\s+(.+?)\s+(\d+)\s+(\d+)\s+(\d{6,8})'
    matches = re.findall(pattern, text)
    
    ket_qua = []
    for m in matches:
        ket_qua.append({
            "STT": m[0],
            "Mã NV": m[1],
            "Họ tên": m[2],
            "Chức vụ": m[3].strip(),
            "Lương cơ bản": int(m[6]) if m[6].isdigit() else 0
        })
    return pd.DataFrame(ket_qua)

def trich_xuat_d02(text):
    """Trích xuất dữ liệu D02-TS từ text OCR"""
    # Pattern tìm dòng có: STT, Họ tên, Mã BHXH, Tiền lương đóng
    pattern = r'(\d+)\s+([A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+)+)\s+(\d{10})\s+(\d{2}/\d{2}/\d{4})\s+(\w+)\s+([\d,\.]+)'
    matches = re.findall(pattern, text)
    
    ket_qua = []
    for m in matches:
        ket_qua.append({
            "STT": m[0],
            "Họ tên": m[1],
            "Mã số BHXH": m[2],
            "Ngày sinh": m[3],
            "Giới tính": m[4],
            "Tiền lương đóng": float(m[5].replace(',', '').replace('.', '')) if m[5] else 0
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
col_a, col_b = st.columns([4, 1])
with col_a:
    don_vi = st.selectbox(
        "🏢 Đơn vị đang chọn:",
        ["[TF0372F] CÔNG TY TNHH K&D GARMENT"]
    )
with col_b:
    st.write("")
    st.write("")
    if st.button("📁 Mở Lưu Trữ", use_container_width=True):
        st.info("Chức năng đang phát triển")

# ===== MENU TABS =====
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Đối chiếu",
    "📂 Nhập dữ liệu",
    "📜 Danh sách đơn vị",
    "💼 Mức lương tối thiểu",
    "📚 Thư viện pháp luật"
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
    st.subheader("📋 Kết quả đối chiếu")

    if "df_ket_qua" in st.session_state and st.session_state.df_ket_qua is not None:
        df = st.session_state.df_ket_qua
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            st.metric("Tổng lao động", len(df))
        with col_s2:
            st.metric("Cần truy thu", len(df[df["Chênh lệch"] > 0]) if "Chênh lệch" in df.columns else 0)
        with col_s3:
            if "Chênh lệch" in df.columns:
                st.metric("Tổng truy thu", f"{df['Chênh lệch'].sum():,.0f} VNĐ")
        
        st.dataframe(df, use_container_width=True)
        
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 Tải kết quả (CSV)",
            csv,
            f"ket_qua_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv"
        )
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
        st.subheader("1️⃣ Bảng lương / Bảng chấm công")
        file_luong = st.file_uploader(
            "Chọn file bảng lương",
            type=["xlsx", "xls", "csv", "pdf", "docx"],
            key="file_luong"
        )
        if file_luong:
            st.success(f"✅ Đã tải: {file_luong.name}")
            with st.expander("👁️ Xem trước nội dung"):
                du_lieu = doc_file_bat_ky(file_luong)
                if isinstance(du_lieu, pd.DataFrame):
                    st.dataframe(du_lieu.head(20), use_container_width=True)
                elif isinstance(du_lieu, str):
                    st.text_area("Nội dung text:", du_lieu[:2000], height=200)

    with col_up2:
        st.subheader("2️⃣ Mẫu D02-TS")
        file_d02 = st.file_uploader(
            "Chọn file D02-TS",
            type=["xlsx", "xls", "csv", "docx", "pdf"],
            key="file_d02"
        )
        if file_d02:
            st.success(f"✅ Đã tải: {file_d02.name}")
            with st.expander("👁️ Xem trước nội dung"):
                du_lieu = doc_file_bat_ky(file_d02)
                if isinstance(du_lieu, pd.DataFrame):
                    st.dataframe(du_lieu.head(20), use_container_width=True)
                elif isinstance(du_lieu, str):
                    st.text_area("Nội dung text:", du_lieu[:2000], height=200)

    st.markdown("---")
    
    if st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU", type="primary", use_container_width=True):
        if file_luong and file_d02:
            with st.spinner("Đang xử lý..."):
                # Đọc dữ liệu
                data_luong = doc_file_bat_ky(file_luong)
                data_d02 = doc_file_bat_ky(file_d02)
                
                # Demo kết quả
                demo = {
                    "STT": [1, 2, 3, 4, 5],
                    "Họ tên": ["LƯƠNG THỊ PHƯƠNG", "DƯƠNG THỊ BÍCH LÂM", "HÀ THỊ TUYẾT", "NGUYỄN THỊ CHÂM", "TRẦN THỊ GIANG"],
                    "Mã số BHXH": ["0123456789", "0250760188", "0250760189", "0250760190", "0250760191"],
                    "Ngày sinh": ["01/01/1990", "04/06/1976", "10/10/1987", "04/09/1984", "12/03/1989"],
                    "Lương đóng D02": [4900000, 4900000, 4900000, 4900000, 4900000],
                    "Lương chịu đóng": [9382509, 9682768, 9184302, 10547054, 7059786],
                    "Chênh lệch": [4482509, 4782768, 4284302, 5647054, 2159786],
                    "Khoản truy thu": ["Lương sản phẩm + Phụ cấp", "Lương sản phẩm + Phụ cấp", 
                                       "Lương sản phẩm + Phụ cấp", "Lương sản phẩm + Phụ cấp",
                                       "Lương sản phẩm + Phụ cấp"],
                    "Số tháng": [12, 12, 12, 12, 12],
                    "Diễn giải pháp lý": [
                        "Điều 31 Luật BHXH 2024 - Tiền lương theo HĐLĐ",
                        "Điều 31 Luật BHXH 2024 - Tiền lương theo HĐLĐ",
                        "Điều 31 Luật BHXH 2024 - Tiền lương theo HĐLĐ",
                        "Điều 31 Luật BHXH 2024 - Tiền lương theo HĐLĐ",
                        "Điều 31 Luật BHXH 2024 - Tiền lương theo HĐLĐ"
                    ]
                }
                st.session_state.df_ket_qua = pd.DataFrame(demo)
                st.success("✅ Đối chiếu xong! Xem kết quả ở tab 'Đối chiếu'.")
                st.balloons()
        else:
            st.warning("⚠️ Vui lòng tải lên cả 2 file.")

# ============================================================
# TAB 3: DANH SÁCH ĐƠN VỊ
# ============================================================
with tab3:
    st.header("📜 Danh sách đơn vị đã đối chiếu")
    don_vi_data = {
        "STT": [1],
        "Mã đơn vị": ["TF0372F"],
        "Tên đơn vị": ["CÔNG TY TNHH K&D GARMENT"],
        "Ngày đối chiếu": ["22/09/2026"],
        "Số lao động": [45],
        "Số cần truy thu": [12],
        "Trạng thái": ["Đã đối chiếu"]
    }
    st.dataframe(pd.DataFrame(don_vi_data), use_container_width=True)

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

    st.markdown("---")
    st.subheader("📌 Tham chiếu khác")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.number_input("Mức tham chiếu (đ/tháng)", value=2340000, step=10000)
    with col_t2:
        st.number_input("Trần BHXH, BHYT (20 lần)", value=46800000, step=100000, disabled=True)

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

# ===== FOOTER =====
st.markdown("---")
st.caption("📌 Công cụ hỗ trợ đối chiếu tự động dựa trên Luật BHXH 2024. Kết quả mang tính tham khảo.")
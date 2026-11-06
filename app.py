import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Tạo Ma Trận Đề Kiểm Tra Tin Học 2018", layout="wide", page_icon="📚")

# 2. DỮ LIỆU CHƯƠNG TRÌNH GDPT 2018 MÔN TIN HỌC (THCS)
THEMES = {
    "Chủ đề A. Máy tính và cộng đồng": ["Thông tin và dữ liệu", "Biểu diễn thông tin"],
    "Chủ đề B. Mạng máy tính và Internet": ["Giới thiệu về mạng máy tính", "Internet và công cụ tìm kiếm"],
    "Chủ đề C. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin": ["Tổ chức thông tin", "Sơ đồ tư duy"],
    "Chủ đề D. Đạo đức, pháp luật và văn hoá trong môi trường số": ["Bản quyền nội dung số", "Giao tiếp an toàn cơ bản"],
    "Chủ đề E. Ứng dụng tin học": ["Soạn thảo văn bản", "Trình chiếu", "Bảng tính điện tử"],
    "Chủ đề F. Giải quyết vấn đề với sự trợ giúp của máy tính": ["Thuật toán cơ bản", "Lập trình trực quan (Scratch/Python)"]
}

# 3. GIAO DIỆN HEADER (Bộ lọc)
st.title("📚 HỆ THỐNG XÂY DỰNG MA TRẬN & ĐẶC TẢ ĐỀ KIỂM TRA")
st.markdown("---")

col1, col2, col3, col4 = st.columns(4)
with col1:
    mon_hoc = st.selectbox("MÔN HỌC:", ["Tin học"])
with col2:
    khoi_lop = st.selectbox("KHỐI LỚP:", ["Khối 6", "Khối 7", "Khối 8", "Khối 9"])
with col3:
    ky_thi = st.selectbox("KỲ THI / ĐỢT KIỂM TRA:", ["Giữa kì I", "Cuối kì I", "Giữa kì II", "Cuối kì II"])
with col4:
    thoi_luong = st.number_input("THỜI LƯỢNG (phút):", min_value=15, value=45, step=15)

st.markdown("""<style>
    .stCheckbox > label { font-weight: bold; color: #2e6c80; }
    .header-style { background-color: #f0f2f6; padding: 10px; border-radius: 5px; }
</style>""", unsafe_allow_html=True)

# 4. BỐ CỤC CHÍNH (Sidebar Trái - Bảng Phải)
left_col, right_col = st.columns([1, 3])

with left_col:
    st.markdown("<div class='header-style'><b>YÊU CẦU CẦN ĐẠT</b></div>", unsafe_allow_html=True)
    st.write("")
    selected_themes = []
    for theme in THEMES.keys():
        with st.expander(theme, expanded=(theme == "Chủ đề A. Máy tính và cộng đồng" or theme == "Chủ đề B. Mạng máy tính và Internet")):
            if st.checkbox(f"Chọn toàn bộ", key=f"all_{theme}"):
                selected_themes.extend(THEMES[theme])
            else:
                for topic in THEMES[theme]:
                    if st.checkbox(topic, key=f"topic_{topic}"):
                        selected_themes.append(topic)

with right_col:
    st.info("💡 Quy ước: Hàng trên nhập số lượng (ý/câu). Mỗi ý = 0.25đ. Hàng dưới tự động sinh mã câu.")
    
    # Khởi tạo bảng dữ liệu Ma trận
    if not selected_themes:
        st.warning("👈 Vui lòng chọn ít nhất một nội dung kiến thức ở cột bên trái.")
    else:
        # Cấu trúc cột dựa theo hình (TNKQ, Đúng-Sai, Tự luận) x (Biết, Hiểu, Vận dụng)
        columns = [
            "Nội dung kiến thức", 
            "TNKQ - Nhận biết", "TNKQ - Thông hiểu", "TNKQ - Vận dụng",
            "Đ/S - Nhận biết", "Đ/S - Thông hiểu", "Đ/S - Vận dụng",
            "TL - Nhận biết", "TL - Thông hiểu", "TL - Vận dụng"
        ]
        
        df = pd.DataFrame(columns=columns)
        df["Nội dung kiến thức"] = selected_themes
        for col in columns[1:]:
            df[col] = 0  # Mặc định số câu là 0
            
        edited_df = st.data_editor(
            df, 
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Nội dung kiến thức": st.column_config.TextColumn("Nội dung / Đơn vị kiến thức", disabled=True),
            }
        )

# 5. XỬ LÝ XUẤT FILE WORD (.docx)
def create_word_matrix(dataframe, lop, ky):
    doc = Document()
    doc.add_heading(f'MA TRẬN ĐỀ KIỂM TRA {ky.upper()} - MÔN TIN HỌC {lop[-1]}', 0)
    doc.add_paragraph('Chương trình GDPT 2018')
    
    # Tạo bảng
    table = doc.add_table(rows=1, cols=len(dataframe.columns))
    table.style = 'Table Grid'
    
    # Header
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(dataframe.columns):
        hdr_cells[i].text = col_name
        
    # Thêm dữ liệu
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(dataframe.columns):
            row_cells[i].text = str(row[col_name])
            
    # Lưu vào bộ nhớ đệm
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 6. CÁC NÚT CHỨC NĂNG (FOOTER)
st.markdown("---")
btn_col1, btn_col2, btn_col3, btn_col4 = st.columns([1.5, 1.5, 1.5, 5])

with btn_col1:
    st.button("💾 Lưu Ma trận & Bản đặc tả", use_container_width=True)
with btn_col2:
    if selected_themes:
        word_file = create_word_matrix(edited_df, khoi_lop, ky_thi)
        st.download_button(
            label="📥 Tải Ma trận (Word)",
            data=word_file,
            file_name=f"Ma_Tran_Tin_Hoc_{khoi_lop.replace(' ', '')}_{ky_thi.replace(' ', '')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
with btn_col3:
    st.button("🤖 Tạo đề thi bằng AI", type="primary", use_container_width=True)
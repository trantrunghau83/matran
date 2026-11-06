import streamlit as st
import pandas as pd
from docx import Document
import io

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Ma Trận Đề Kiểm Tra GDPT 2018", layout="wide", page_icon="📚")

# 2. DỮ LIỆU CHƯƠNG TRÌNH
THEMES = {
    "Chủ đề A. Máy tính và cộng đồng": ["Thông tin và dữ liệu", "Biểu diễn thông tin"],
    "Chủ đề B. Mạng máy tính và Internet": ["Giới thiệu về mạng máy tính", "Internet và công cụ tìm kiếm"],
    "Chủ đề C. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin": ["Tổ chức thông tin", "Sơ đồ tư duy"],
    "Chủ đề D. Đạo đức, pháp luật và văn hoá trong môi trường số": ["Bản quyền nội dung số", "Giao tiếp an toàn cơ bản"],
    "Chủ đề E. Ứng dụng tin học": ["Soạn thảo văn bản", "Trình chiếu", "Bảng tính điện tử"],
    "Chủ đề F. Giải quyết vấn đề với sự trợ giúp của máy tính": ["Thuật toán cơ bản", "Lập trình trực quan"]
}

# 3. HEADER & BỘ LỌC
st.title("📚 HỆ THỐNG XÂY DỰNG MA TRẬN & ĐẶC TẢ ĐỀ KIỂM TRA")
st.markdown("---")

col1, col2, col3, col4 = st.columns(4)
mon_hoc = col1.selectbox("MÔN HỌC:", ["Tin học"])
khoi_lop = col2.selectbox("KHỐI LỚP:", ["Khối 6", "Khối 7", "Khối 8", "Khối 9"])
ky_thi = col3.selectbox("KỲ THI / ĐỢT:", ["Giữa kì I", "Cuối kì I", "Giữa kì II", "Cuối kì II"])
thoi_luong = col4.number_input("THỜI LƯỢNG (phút):", min_value=15, value=45, step=15)

st.markdown("""
<style>
    .header-style { background-color: #f0f2f6; padding: 10px; border-radius: 5px; text-align: center;}
</style>
""", unsafe_allow_html=True)

# 4. CHIA CỘT GIAO DIỆN
left_col, right_col = st.columns([1, 2.5])

# CỘT TRÁI: CHỌN CHỦ ĐỀ
with left_col:
    st.markdown("<div class='header-style'><b>BỘ KHUNG KIẾN THỨC YÊU CẦU</b></div>", unsafe_allow_html=True)
    st.write("")
    
    selected_themes = []
    for theme, topics in THEMES.items():
        # Mặc định mở rộng 2 chủ đề đầu để giao diện giống mẫu
        is_expanded = (theme in ["Chủ đề A. Máy tính và cộng đồng", "Chủ đề B. Mạng máy tính và Internet"])
        with st.expander(theme, expanded=is_expanded):
            # Tuỳ chọn check tất cả
            if st.checkbox("✅ Chọn toàn bộ chủ đề này", key=f"all_{theme}"):
                selected_themes.extend(topics)
            else:
                for topic in topics:
                    if st.checkbox(f"• {topic}", key=f"topic_{topic}"):
                        selected_themes.append(topic)

# CỘT PHẢI: BẢNG MA TRẬN
edited_df = None
with right_col:
    st.info("💡 Quy ước: Nhập số lượng ý/câu vào bảng bên dưới. Mỗi ý/câu tương ứng 0.25đ. Hệ thống sẽ tự động tổng hợp.")
    
    if len(selected_themes) == 0:
        st.warning("👈 Vui lòng đánh dấu chọn ít nhất một nội dung kiến thức ở cột bên trái để hiển thị bảng ma trận.")
    else:
        # Cấu trúc các cột điểm
        columns = [
            "Nội dung kiến thức", 
            "TNKQ - Nhận biết", "TNKQ - Thông hiểu", "TNKQ - Vận dụng",
            "Đ/S - Nhận biết", "Đ/S - Thông hiểu", "Đ/S - Vận dụng",
            "TL - Nhận biết", "TL - Thông hiểu", "TL - Vận dụng"
        ]
        
        # Tạo khung dữ liệu ban đầu
        df = pd.DataFrame(columns=columns)
        df["Nội dung kiến thức"] = selected_themes
        for col in columns[1:]:
            df[col] = 0  # Gán giá trị mặc định là 0
            
        # Hiển thị bảng cho phép nhập liệu
        edited_df = st.data_editor(
            df, 
            use_container_width=True,
            hide_index=True,
            disabled=["Nội dung kiến thức"], # Không cho phép sửa tên bài học
        )

# 5. HÀM TẠO FILE WORD
def create_word_matrix(dataframe, lop, ky):
    doc = Document()
    
    # Tiêu đề file Word
    doc.add_heading(f'MA TRẬN ĐỀ KIỂM TRA {ky.upper()} - MÔN TIN HỌC {lop[-1]}', level=1)
    doc.add_paragraph('Chương trình Giáo dục phổ thông 2018')
    
    # Tạo bảng trong Word
    table = doc.add_table(rows=1, cols=len(dataframe.columns))
    table.style = 'Table Grid'
    
    # Đổ Header
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(dataframe.columns):
        hdr_cells[i].text = col_name
        
    # Đổ dữ liệu từ Streamlit vào Word
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(dataframe.columns):
            row_cells[i].text = str(row[col_name])
            
    # Lưu vào RAM (BytesIO) để tải xuống
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 6. KHU VỰC NÚT BẤM (FOOTER)
st.markdown("---")
btn_col1, btn_col2, btn_col3, btn_col4 = st.columns([2, 2, 2, 4])

with btn_col1:
    st.button("💾 Lưu Cấu Hình Tạm", use_container_width=True)
    
with btn_col2:
    if edited_df is not None: # Nếu bảng đã được tạo
        try:
            word_file = create_word_matrix(edited_df, khoi_lop, ky_thi)
            st.download_button(
                label="📥 Xuất file Word",
                data=word_file,
                file_name=f"Ma_Tran_Tin_Hoc_{khoi_lop.replace(' ', '')}_{ky_thi.replace(' ', '')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
        except Exception as e:
            st.error(f"Lỗi tạo file Word: {e}")
            
with btn_col3:
    st.button("🤖 AI Sinh Đề (Demo)", use_container_width=True)

import streamlit as st
import pandas as pd
from docx import Document
import io

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Ma Trận Đề Kiểm Tra GDPT 2018", layout="wide", page_icon="📚")

# 2. DỮ LIỆU CHƯƠNG TRÌNH GDPT 2018 (Từ lớp 6 đến lớp 9)
CURRICULUM = {
    "Khối 6": {
        "Chủ đề A. Máy tính và cộng đồng": ["Thông tin và dữ liệu", "Biểu diễn thông tin"],
        "Chủ đề B. Mạng máy tính và Internet": ["Giới thiệu về mạng máy tính", "Internet và công cụ tìm kiếm"],
        "Chủ đề C. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin": ["Tổ chức thông tin", "Sơ đồ tư duy"],
        "Chủ đề D. Đạo đức, pháp luật và văn hoá trong môi trường số": ["Bản quyền nội dung số", "Giao tiếp an toàn cơ bản"],
        "Chủ đề E. Ứng dụng tin học": ["Soạn thảo văn bản", "Trình chiếu cơ bản"],
        "Chủ đề F. Giải quyết vấn đề với sự trợ giúp của máy tính": ["Thuật toán cơ bản", "Lập trình trực quan (Scratch)"]
    },
    "Khối 7": {
        "Chủ đề A. Máy tính và cộng đồng": ["Thiết bị vào - ra", "Hệ điều hành và phần mềm ứng dụng"],
        "Chủ đề B. Mạng máy tính và Internet": ["Mạng xã hội và trao đổi thông tin"],
        "Chủ đề C. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin": ["Quản lí dữ liệu trong máy tính"],
        "Chủ đề D. Đạo đức, pháp luật và văn hoá trong môi trường số": ["Ứng xử trên mạng xã hội", "Bảo vệ dữ liệu cá nhân"],
        "Chủ đề E. Ứng dụng tin học": ["Phần mềm bảng tính", "Phần mềm trình chiếu (Nâng cao)"],
        "Chủ đề F. Giải quyết vấn đề với sự trợ giúp của máy tính": ["Thuật toán tìm kiếm", "Thuật toán sắp xếp"]
    },
    "Khối 8": {
        "Chủ đề A. Máy tính và cộng đồng": ["Lịch sử phát triển máy tính", "Dữ liệu số trong thời đại thông tin"],
        "Chủ đề B. Mạng máy tính và Internet": ["Thực hành khai thác thông tin số"],
        "Chủ đề C. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin": ["Sắp xếp và lọc dữ liệu"],
        "Chủ đề D. Đạo đức, pháp luật và văn hoá trong môi trường số": ["Đạo đức và văn hoá sử dụng công nghệ kĩ thuật số"],
        "Chủ đề E. Ứng dụng tin học": ["Xử lý đồ hoạ/hình ảnh", "Bản mẫu trình chiếu", "Thực hành Bảng tính (Địa chỉ ô)"],
        "Chủ đề F. Giải quyết vấn đề với sự trợ giúp của máy tính": ["Lập trình trực quan (Rẽ nhánh và Lặp)"]
    },
    "Khối 9": {
        "Chủ đề A. Máy tính và cộng đồng": ["Tin học và xã hội", "Chất lượng thông tin"],
        "Chủ đề B. Mạng máy tính và Internet": ["Khám phá thế giới qua Internet", "Giao tiếp qua không gian mạng"],
        "Chủ đề C. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin": ["Tổ chức cây thư mục nâng cao", "Công cụ tìm kiếm nâng cao"],
        "Chủ đề D. Đạo đức, pháp luật và văn hoá trong môi trường số": ["Pháp luật và bản quyền số", "Bảo mật thông tin cơ bản"],
        "Chủ đề E. Ứng dụng tin học": ["Phần mềm sơ đồ tư duy", "Phần mềm làm video", "Thực tại ảo/Thực tại tăng cường"],
        "Chủ đề F. Giải quyết vấn đề với sự trợ giúp của máy tính": ["Ngôn ngữ lập trình Python/C++ cơ bản", "Chương trình con và mảng"]
    }
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

# LẤY DỮ LIỆU CHỦ ĐỀ THEO KHỐI LỚP ĐƯỢC CHỌN
THEMES = CURRICULUM[khoi_lop]

# 4. CHIA CỘT GIAO DIỆN
left_col, right_col = st.columns([1, 2.5])

# CỘT TRÁI: CHỌN CHỦ ĐỀ
with left_col:
    st.markdown(f"<div class='header-style'><b>BỘ KHUNG KIẾN THỨC YÊU CẦU ({khoi_lop.upper()})</b></div>", unsafe_allow_html=True)
    st.write("")
    
    selected_themes = []
    for theme, topics in THEMES.items():
        is_expanded = (theme in ["Chủ đề A. Máy tính và cộng đồng", "Chủ đề B. Mạng máy tính và Internet"])
        with st.expander(theme, expanded=is_expanded):
            # Tạo key duy nhất để không bị lỗi khi chuyển khối lớp
            if st.checkbox("✅ Chọn toàn bộ", key=f"all_{khoi_lop}_{theme}"):
                selected_themes.extend(topics)
            else:
                for topic in topics:
                    if st.checkbox(f"• {topic}", key=f"topic_{khoi_lop}_{topic}"):
                        selected_themes.append(topic)

# CỘT PHẢI: BẢNG MA TRẬN
edited_df = None
with right_col:
    st.info("💡 Quy ước: Nhập số lượng ý/câu vào bảng bên dưới. Mỗi ý/câu tương ứng 0.25đ. Hàng dưới tự động sinh mã câu.")
    
    if len(selected_themes) == 0:
        st.warning(f"👈 Vui lòng đánh dấu chọn ít nhất một nội dung kiến thức ở cột bên trái của {khoi_lop}.")
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
            df[col] = 0 
            
        edited_df = st.data_editor(
            df, 
            use_container_width=True,
            hide_index=True,
            disabled=["Nội dung kiến thức"],
        )

# 5. HÀM TẠO FILE WORD
def create_word_matrix(dataframe, lop, ky):
    doc = Document()
    doc.add_heading(f'MA TRẬN ĐỀ KIỂM TRA {ky.upper()} - MÔN TIN HỌC {lop[-1]}', level=1)
    doc.add_paragraph('Chương trình Giáo dục phổ thông 2018')
    
    table = doc.add_table(rows=1, cols=len(dataframe.columns))
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(dataframe.columns):
        hdr_cells[i].text = col_name
        
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(dataframe.columns):
            row_cells[i].text = str(row[col_name])
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 6. KHU VỰC NÚT BẤM (FOOTER)
st.markdown("---")
btn_col1, btn_col2, btn_col3, btn_col4 = st.columns([2, 2, 2, 4])

with btn_col1:
    st.button("💾 Lưu Cấu Hình Tạm", use_container_width=True)
    
with btn_col2:
    if edited_df is not None:
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

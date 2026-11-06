import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import io

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Ma Trận & Đặc Tả GDPT 2018", layout="wide", page_icon="📚")

# 2. DỮ LIỆU CHƯƠNG TRÌNH & YÊU CẦU CẦN ĐẠT (Trích mẫu)
CURRICULUM = {
    "Khối 6": {
        "Chủ đề 1: Máy tính và cộng đồng": {
            "Thông tin và dữ liệu": "- Nhận biết: Định nghĩa thông tin, dữ liệu, vật mang tin.\n- Thông hiểu: Phân biệt dữ liệu và thông tin.\n- Vận dụng: Đánh giá ý nghĩa thông tin để ra quyết định.",
            "Xử lí thông tin": "- Nhận biết: Các bước xử lí thông tin.\n- Thông hiểu: Mối quan hệ giữa máy tính và con người trong xử lí thông tin.\n- Vận dụng: Đưa ra ví dụ thực tế.",
            "Thông tin trong máy tính": "- Nhận biết: Biểu diễn bit, byte.\n- Thông hiểu: Đổi các đơn vị đo.\n- Vận dụng: Ước lượng dung lượng file."
        },
        "Chủ đề 2: Mạng máy tính và Internet": {
            "Mạng máy tính": "- Nhận biết: Khái niệm mạng, các thành phần.\n- Thông hiểu: Vai trò của thiết bị mạng.\n- Vận dụng: Phân biệt mạng có dây, không dây."
        }
    },
    # (Để ngắn gọn, bạn có thể tự thêm dữ liệu Khối 7,8,9 tương tự cấu trúc trên)
    "Khối 7": {"Chủ đề 1: Máy tính và cộng đồng": {"Thiết bị vào - ra": "- YCCĐ: ...", "Hệ điều hành": "- YCCĐ: ..."}},
    "Khối 8": {"Chủ đề 1: Máy tính và cộng đồng": {"Lịch sử máy tính": "- YCCĐ: ...", "Dữ liệu số": "- YCCĐ: ..."}},
    "Khối 9": {"Chủ đề 1: Máy tính và cộng đồng": {"Tin học và xã hội": "- YCCĐ: ...", "Chất lượng thông tin": "- YCCĐ: ..."}}
}

# 3. HEADER & BỘ LỌC
st.title("📚 HỆ THỐNG XÂY DỰNG MA TRẬN & ĐẶC TẢ ĐỀ KIỂM TRA (CHUẨN MẪU)")
st.markdown("---")

col1, col2, col3, col4 = st.columns(4)
mon_hoc = col1.selectbox("MÔN HỌC:", ["Tin học"])
khoi_lop = col2.selectbox("KHỐI LỚP:", ["Khối 6", "Khối 7", "Khối 8", "Khối 9"])
ky_thi = col3.selectbox("KỲ THI / ĐỢT:", ["Giữa học kì 1", "Cuối học kì 1", "Giữa học kì 2", "Cuối học kì 2"])
thoi_luong = col4.number_input("THỜI LƯỢNG (phút):", min_value=15, value=45, step=15)

THEMES = CURRICULUM[khoi_lop]

# 4. CHIA CỘT GIAO DIỆN
left_col, right_col = st.columns([1, 3])

with left_col:
    st.markdown(f"<div style='background-color:#f0f2f6; padding:10px; text-align:center;'><b>BỘ KHUNG KIẾN THỨC ({khoi_lop.upper()})</b></div>", unsafe_allow_html=True)
    st.write("")
    selected_topics = []
    
    for theme, topics in THEMES.items():
        with st.expander(theme, expanded=True):
            for topic, yccd in topics.items():
                if st.checkbox(f"• {topic}", key=f"topic_{khoi_lop}_{topic}"):
                    selected_topics.append({"Chủ đề": theme, "Nội dung": topic, "YCCĐ": yccd})

edited_df = None
with right_col:
    st.info("💡 Quy ước: Nhập số lượng ý/câu. Trắc nghiệm NLC (0.25đ/câu), Đúng-Sai (0.25đ/ý), Tự luận (Điểm linh hoạt tự chia).")
    
    if len(selected_topics) == 0:
        st.warning(f"👈 Vui lòng chọn ít nhất một nội dung kiến thức.")
    else:
        columns = [
            "Chủ đề", "Nội dung kiến thức", 
            "NLC-Biết", "NLC-Hiểu", "NLC-VD",
            "ĐS-Biết", "ĐS-Hiểu", "ĐS-VD",
            "TL-Biết", "TL-Hiểu", "TL-VD"
        ]
        
        df = pd.DataFrame(columns=columns)
        for i, item in enumerate(selected_topics):
            df.loc[i] = [item["Chủ đề"], item["Nội dung"]] + [0]*9
            
        edited_df = st.data_editor(
            df, use_container_width=True, hide_index=True, disabled=["Chủ đề", "Nội dung kiến thức"]
        )

# 5. HÀM GHI TẬP TIN WORD THEO ĐÚNG MẪU ĐẶC TẢ
def set_cell(cell, text, bold=False, center=True):
    cell.text = str(text)
    for p in cell.paragraphs:
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = bold
            r.font.name = 'Times New Roman'
            r.font.size = Pt(11)

def create_word_export(dataframe, yccd_data, lop, ky):
    doc = Document()
    
    # -----------------------------
    # PHẦN 1: MA TRẬN ĐỀ KIỂM TRA
    # -----------------------------
    doc.add_heading(f'MA TRẬN ĐỀ KIỂM TRA {ky.upper()} - MÔN TIN HỌC {lop[-1]}', level=1)
    doc.add_paragraph('NĂM HỌC 2026-2027\n').alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Bảng 16 cột: TT, Chủ đề, Nội dung, [NLC: 3], [ĐS: 3], [TL: 3], Tổng TN, Tổng ĐS, Tổng TL, Điểm
    t = doc.add_table(rows=4, cols=16)
    t.style = 'Table Grid'
    
    # Merge Header 
    t.cell(0,0).merge(t.cell(3,0)); set_cell(t.cell(0,0), "TT", True)
    t.cell(0,1).merge(t.cell(3,1)); set_cell(t.cell(0,1), "Chủ đề/\nChương", True)
    t.cell(0,2).merge(t.cell(3,2)); set_cell(t.cell(0,2), "Nội dung/\nĐơn vị kiến thức", True)
    
    t.cell(0,3).merge(t.cell(0,11)); set_cell(t.cell(0,3), "Mức độ đánh giá", True)
    t.cell(1,3).merge(t.cell(1,8)); set_cell(t.cell(1,3), "TNKQ", True)
    t.cell(1,9).merge(t.cell(2,11)); set_cell(t.cell(1,9), "Tự luận", True)
    
    t.cell(2,3).merge(t.cell(2,5)); set_cell(t.cell(2,3), "Nhiều lựa chọn", True)
    t.cell(2,6).merge(t.cell(2,8)); set_cell(t.cell(2,6), "Đúng - Sai", True)
    
    for i in range(3):
        set_cell(t.cell(3, 3+i), ["Biết", "Hiểu", "VD"][i])
        set_cell(t.cell(3, 6+i), ["Biết", "Hiểu", "VD"][i])
        set_cell(t.cell(3, 9+i), ["Biết", "Hiểu", "VD"][i])
        
    t.cell(0,12).merge(t.cell(1,14)); set_cell(t.cell(0,12), "Tổng", True)
    t.cell(2,12).merge(t.cell(3,12)); set_cell(t.cell(2,12), "TN")
    t.cell(2,13).merge(t.cell(3,13)); set_cell(t.cell(2,

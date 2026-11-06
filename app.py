import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import io

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Ma Trận & Đặc Tả GDPT 2018", layout="wide", page_icon="📚")

# 2. DỮ LIỆU CHƯƠNG TRÌNH & YÊU CẦU CẦN ĐẠT
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
    "Khối 7": {"Chủ đề 1: Máy tính và cộng đồng": {"Thiết bị vào - ra": "- Nhận biết: ...\n- Thông hiểu: ...", "Hệ điều hành": "- Nhận biết: ...\n- Thông hiểu: ..."}},
    "Khối 8": {"Chủ đề 1: Máy tính và cộng đồng": {"Lịch sử máy tính": "- Nhận biết: ...\n- Thông hiểu: ...", "Dữ liệu số": "- Nhận biết: ...\n- Thông hiểu: ..."}},
    "Khối 9": {"Chủ đề 1: Máy tính và cộng đồng": {"Tin học và xã hội": "- Nhận biết: ...\n- Thông hiểu: ...", "Chất lượng thông tin": "- Nhận biết: ...\n- Thông hiểu: ..."}}
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
    st.info("💡 Quy ước: Nhập số lượng ý/câu. Trắc nghiệm NLC (0.25đ/câu), Đúng-Sai (0.25đ/ý), Tự luận (Điểm linh hoạt).")
    
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
    
    # --- PHẦN 1: MA TRẬN ---
    doc.add_heading(f'MA TRẬN ĐỀ KIỂM TRA {ky.upper()} - MÔN TIN HỌC {lop[-1]}', level=1)
    doc.add_paragraph('NĂM HỌC 2026-2027\n').alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    t = doc.add_table(rows=4, cols=16)
    t.style = 'Table Grid'
    
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
    t.cell(2,13).merge(t.cell(3,13)); set_cell(t.cell(2,13), "ĐS")
    t.cell(2,14).merge(t.cell(3,14)); set_cell(t.cell(2,14), "TL")
    
    t.cell(0,15).merge(t.cell(3,15)); set_cell(t.cell(0,15), "Điểm", True)
    
    tt = 1
    sum_cols = [0]*13 
    current_theme = ""
    for index, row in dataframe.iterrows():
        r = t.add_row().cells
        if row["Chủ đề"] != current_theme:
            set_cell(r[0], str(tt), center=True)
            set_cell(r[1], row["Chủ đề"], center=False)
            current_theme = row["Chủ đề"]
            tt += 1
        
        set_cell(r[2], row["Nội dung kiến thức"], center=False)
        
        for i in range(9):
            val = int(row[columns[2+i]])
            set_cell(r[3+i], val if val>0 else "")
            sum_cols[i] += val
            
        sum_tn = int(row["NLC-Biết"]) + int(row["NLC-Hiểu"]) + int(row["NLC-VD"])
        sum_ds = int(row["ĐS-Biết"]) + int(row["ĐS-Hiểu"]) + int(row["ĐS-VD"])
        sum_tl = int(row["TL-Biết"]) + int(row["TL-Hiểu"]) + int(row["TL-VD"])
        diem_tn = sum_tn * 0.25 + sum_ds * 0.25 
        
        set_cell(r[12], sum_tn if sum_tn>0 else "")
        set_cell(r[13], sum_ds if sum_ds>0 else "")
        set_cell(r[14], sum_tl if sum_tl>0 else "")
        set_cell(r[15], str(diem_tn))
        
        sum_cols[9] += sum_tn; sum_cols[10] += sum_ds; sum_cols[11] += sum_tl; sum_cols[12] += diem_tn

    r_total = t.add_row().cells
    r_total[0].merge(r_total[2]); set_cell(r_total[0], "Tổng số câu", True)
    for i in range(13):
        set_cell(r_total[3+i], sum_cols[i], True)

    doc.add_page_break()

    # --- PHẦN 2: BẢN ĐẶC TẢ ---
    doc.add_heading(f'BẢNG ĐẶC TẢ ĐỀ KIỂM TRA {ky.upper()} - TIN HỌC {lop[-1]}', level=1)
    doc.add_paragraph('NĂM HỌC 2026-2027\n').alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    t2 = doc.add_table(rows=4, cols=16)
    t2.style = 'Table Grid'
    
    t2.cell(0,0).merge(t2.cell(3,0)); set_cell(t2.cell(0,0), "TT", True)
    t2.cell(0,1).merge(t2.cell(3,1)); set_cell(t2.cell(0,1), "Chủ đề", True)
    t2.cell(0,2).merge(t2.cell(3,2)); set_cell(t2.cell(0,2), "Nội dung", True)
    t2.cell(0,3).merge(t2.cell(3,3)); set_cell(t2.cell(0,3), "Yêu cầu cần đạt", True)
    
    t2.cell(0,4).merge(t2.cell(0,12)); set_cell(t2.cell(0,4), "Mức độ đánh giá", True)
    t2.cell(1,4).merge(t2.cell(1,9)); set_cell(t2.cell(1,4), "TNKQ", True)
    t2.cell(1,10).merge(t2.cell(2,12)); set_cell(t2.cell(1,10), "Tự luận", True)
    t2.cell(2,4).merge(t2.cell(2,6)); set_cell(t2.cell(2,4), "Nhiều lựa chọn", True)
    t2.cell(2,7).merge(t2.cell(2,9)); set_cell(t2.cell(2,7), "Đúng - Sai", True)
    
    for i in range(3):
        set_cell(t2.cell(3, 4+i), ["Biết", "Hiểu", "VD"][i])
        set_cell(t2.cell(3, 7+i), ["Biết", "Hiểu", "VD"][i])
        set_cell(t2.cell(3, 10+i), ["Biết", "Hiểu", "VD"][i])
        
    t2.cell(0,13).merge(t2.cell(1,15)); set_cell(t2.cell(0,13), "Tổng", True)
    t2.cell(2,13).merge(t2.cell(3,13)); set_cell(t2.cell(2,13), "TN")
    t2.cell(2,14).merge(t2.cell(3,14)); set_cell(t2.cell(2,14), "ĐS")
    t2.cell(2,15).merge(t2.cell(3,15)); set_cell(t2.cell(2,15), "TL")

    tt2 = 1
    current_theme2 = ""
    for index, row in dataframe.iterrows():
        r = t2.add_row().cells
        if row["Chủ đề"] != current_theme2:
            set_cell(r[0], str(tt2), center=True)
            set_cell(r[1], row["Chủ đề"], center=False)
            current_theme2 = row["Chủ đề"]
            tt2 += 1
            
        set_cell(r[2], row["Nội dung kiến thức"], center=False)
        
        yccd = ""
        for item in yccd_data:
            if item["Nội dung"] == row["Nội dung kiến thức"]:
                yccd = item["YCCĐ"]
                break
        
        set_cell(r[3], yccd, center=False)
        
        for i in range(9):
            val = int(row[columns[2+i]])
            set_cell(r[4+i], val if val>0 else "")
            
        sum_tn = int(row["NLC-Biết"]) + int(row["NLC-Hiểu"]) + int(row["NLC-VD"])
        sum_ds = int(row["ĐS-Biết"]) + int(row["ĐS-Hiểu"]) + int(row["ĐS-VD"])
        sum_tl = int(row["TL-Biết"]) + int(row["TL-Hiểu"]) + int(row["TL-VD"])
        set_cell(r[13], sum_tn if sum_tn>0 else "")
        set_cell(r[14], sum_ds if sum_ds>0 else "")
        set_cell(r[15], sum_tl if sum_tl>0 else "")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 6. NÚT CHỨC NĂNG TẢI XUỐNG
st.markdown("---")
btn_col = st.columns([2, 4, 2])
with btn_col[1]:
    if edited_df is not None:
        word_file = create_word_export(edited_df, selected_topics, khoi_lop, ky_thi)
        st.download_button(
            label="📥 Tải File Word (Chuẩn Ma Trận & Đặc Tả)",
            data=word_file,
            file_name=f"Ma_Tran_Dac_Ta_Tin_Hoc_{khoi_lop.replace(' ', '')}_{ky_thi.replace(' ', '')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary"
        )

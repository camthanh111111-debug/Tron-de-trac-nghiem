import streamlit as st
import os
import tempfile
import shutil
import docx
import re
from dtmix_core import DTMIXApp 

# ================= CẤU HÌNH GIAO DIỆN =================
st.set_page_config(page_title="DTMIX - Trộn Đề Online", page_icon="🚀", layout="wide")

st.markdown("""
    <style>
    .main-header { color: #1E3A8A; font-weight: bold; text-align: center; margin-bottom: 0px;}
    .sub-header { color: #3B82F6; text-align: center; font-size: 18px; margin-bottom: 20px;}
    .correct-ans { color: #E11D48; font-weight: bold; background-color: #ffe4e6; padding: 2px 6px; border-radius: 4px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-header'>✨ DTMIX ONLINE</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Công cụ nhận diện và trộn đề trắc nghiệm thông minh</p>", unsafe_allow_html=True)

# ================= THANH BÊN (SIDEBAR) THÔNG TIN =================
with st.sidebar:
    st.markdown("### 📝 THÔNG TIN ĐẦU TRANG")
    so = st.text_input("Sở GD&ĐT / Phòng", "SỞ GIÁO DỤC VÀ ĐÀO TẠO")
    truong = st.text_input("Tên Trường", "TRƯỜNG THPT...")
    kythi = st.text_input("Tên Kỳ Thi", "KIỂM TRA HỌC KỲ II")
    namhoc = st.text_input("Năm Học", "NĂM HỌC 2025 - 2026")
    monthi = st.text_input("Môn Thi", "Môn: TIẾNG ANH")
    thoigian = st.text_input("Thời Gian Làm Bài", "Thời gian làm bài: 45 phút")
    
    st.markdown("---")
    st.markdown("### ⚙️ CÀI ĐẶT TRỘN ĐỀ")
    made_input = st.text_input("Nhập mã đề (cách nhau dấu phẩy)", "101, 102, 103, 104")

# ================= KHU VỰC CHÍNH (3 TABS) =================
tab1, tab2, tab3 = st.tabs(["📂 1. TẢI ĐỀ GỐC", "🔍 2. RÀ SOÁT & XEM TRƯỚC", "🚀 3. XUẤT ĐỀ"])

with tab1:
    st.markdown("### ⚙️ TẢI LÊN TỆP ĐỀ GỐC (.DOCX)")
    
    # Nút bật chế độ Tiếng Anh làm nổi bật
    st.markdown("---")
    st.markdown("**🇬🇧 TÙY CHỌN CHO MÔN TIẾNG ANH / CÂU CHÙM:**")
    youngmix_mode = st.toggle("Bật nhận diện Tiếng Anh (Bài đọc điền từ, bài đọc hiểu dùng thẻ <g3>, <g2>...)", value=False)
    st.info("Nếu đề có bài đọc (câu chùm), hãy chắc chắn bạn đã đánh dấu các thẻ <g0>, <g1>, <g2>, <g3> trong file Word gốc.")
    st.markdown("---")
    
    uploaded_file = st.file_uploader("Kéo thả file đề gốc (.docx) vào đây", type=["docx"])
    
    if uploaded_file is not None:
        if "last_uploaded" not in st.session_state or st.session_state.last_uploaded != uploaded_file.name or st.session_state.ym_mode != youngmix_mode:
            with st.spinner("Đang đọc và phân tích cấu trúc đề..."):
                temp_dir = tempfile.mkdtemp()
                temp_input_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Phân tích
                parser = DTMIXApp(temp_input_path, temp_dir, so, truong, kythi, namhoc, monthi, thoigian, youngmix_mode, made_input)
                doc = docx.Document(temp_input_path)
                parser.convert_auto_numbering(doc)
                parser.split_soft_returns(doc)
                
                parsed_data, errors = parser.parse_document_structure(doc)
                st.session_state.parsed_data = parsed_data
                
                if youngmix_mode:
                    parser.parse_youngmix_structure(doc)
                    st.session_state.youngmix_data = parser.youngmix_data
                else:
                    st.session_state.youngmix_data = None
                
                st.session_state.last_uploaded = uploaded_file.name
                st.session_state.ym_mode = youngmix_mode
                st.session_state.temp_input_path = temp_input_path
                st.session_state.temp_dir = temp_dir
            
        st.success("✅ Đã tải và phân tích xong! Chuyển sang Tab 2 để kiểm tra xem phần mềm nhận diện đúng chưa.")

with tab2:
    st.markdown("### 📊 KIỂM TRA NỘI DUNG ĐÃ NHẬN DIỆN")
    
    if "parsed_data" not in st.session_state:
        st.info("⚠️ Vui lòng tải file đề gốc lên ở Tab 1 trước.")
    else:
        parsed_data = st.session_state.parsed_data
        is_ym = st.session_state.ym_mode
        
        st.markdown("*Lưu ý: Các đáp án được phần mềm hiểu là **ĐÁP ÁN ĐÚNG** sẽ được tô nền <span class='correct-ans'>màu đỏ</span> bên dưới. Hãy kiểm tra kĩ bằng mắt!*", unsafe_allow_html=True)
        st.write("---")

        with st.container(height=600): # Khung cuộn xem trước đề
            if not is_ym:
                # HIỂN THỊ ĐỀ THƯỜNG
                for p_idx, part in enumerate(parsed_data.get("parts", [])):
                    p_type = part.get("type")
                    part_name = f"PHẦN {('I', 'II', 'III', 'IV')[p_type-1] if p_type <= 4 else p_type}"
                    st.markdown(f"#### 🏷️ {part_name}")
                    
                    for muc in part.get("mucs", []):
                        for q_idx, q in enumerate(muc.get("questions", [])):
                            q_text = q.get('raw_text', '')
                            st.write(f"**{q_text}**")
                            
                            for a_idx, ans in enumerate(q.get("answers", [])):
                                ans_text = ans.get('raw_text', ans.get('text', f'Phương án {a_idx+1}'))
                                if ans.get("is_true"):
                                    st.markdown(f"<span class='correct-ans'>✔ {ans_text}</span>", unsafe_allow_html=True)
                                else:
                                    st.write(f"○ {ans_text}")
                            st.write("") 
            else:
                # HIỂN THỊ ĐỀ TIẾNG ANH (YOUNGMIX)
                ym_data = st.session_state.youngmix_data
                for g_idx, group in enumerate(ym_data):
                    st.markdown(f"#### 🏷️ NHÓM {g_idx + 1} (Chế độ: {group.get('type', 'Không rõ')})")
                    # Hiển thị tiêu đề bài đọc / đoạn văn chung
                    header = group.get('header', '')
                    if header:
                        st.info(header)
                    
                    for q in group.get("questions", []):
                        st.write(f"**{q.get('raw_text', '')}**")
                        for a_idx, ans in enumerate(q.get("answers", [])):
                            ans_text = ans.get('raw_text', ans.get('text', ''))
                            if ans.get("is_true"):
                                st.markdown(f"<span class='correct-ans'>✔ {ans_text}</span>", unsafe_allow_html=True)
                            else:
                                st.write(f"○ {ans_text}")
                        st.write("")

with tab3:
    st.markdown("### 🚀 TRỘN ĐỀ VÀ LƯU FILE")
    st.info("Nếu cấu trúc và đáp án đúng ở Tab 2 đã hiển thị chính xác, bạn có thể trộn đề.")
    
    if "parsed_data" not in st.session_state:
        st.warning("⚠️ Vui lòng tải file đề gốc lên ở Tab 1 trước.")
    else:
        if st.button("🚀 TIẾN HÀNH TRỘN ĐỀ", type="primary", use_container_width=True):
            with st.spinner("Đang hoán vị và đóng gói file Word... Vui lòng đợi trong giây lát!"):
                try:
                    temp_input_path = st.session_state.temp_input_path
                    temp_dir = st.session_state.temp_dir
                    out_dir = os.path.join(temp_dir, "KetQua")
                    os.makedirs(out_dir, exist_ok=True)

                    app = DTMIXApp(
                        filepath=temp_input_path, out_dir=out_dir,
                        so=so, truong=truong, kythi=kythi, 
                        namhoc=namhoc, monthi=monthi, thoigian=thoigian,
                        youngmix=st.session_state.ym_mode, ma_de_str=made_input
                    )
                    
                    app.process_web()

                    # Nén ZIP
                    zip_path = os.path.join(temp_dir, "De_Da_Tron")
                    shutil.make_archive(zip_path, 'zip', out_dir)

                    with open(zip_path + ".zip", "rb") as f:
                        st.session_state['download_data'] = f.read()
                        st.session_state['success'] = True
                        
                except Exception as e:
                    st.error(f"❌ Có lỗi xảy ra trong quá trình trộn: {e}")

        # HIỂN THỊ NÚT TẢI XUỐNG SAU KHI TRỘN XONG
        if st.session_state.get('success'):
            st.success("✅ Trộn đề hoàn tất! Bấm nút bên dưới để chọn nơi lưu.")
            st.download_button(
                label="💾 LƯU BỘ ĐỀ ĐÃ TRỘN VÀO MÁY (.ZIP)",
                data=st.session_state['download_data'],
                file_name="Bo_De_DTMIX_Online.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

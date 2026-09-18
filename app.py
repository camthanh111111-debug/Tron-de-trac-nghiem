import streamlit as st
import os
import tempfile
import shutil
import docx
import re
import pandas as pd
from dtmix_core import DTMIXApp 

# ================= CẤU HÌNH GIAO DIỆN (Giống Desktop) =================
st.set_page_config(page_title="DTMIX - Trộn Đề Trắc Nghiệm", page_icon="📝", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f0f2f6; }
    .main-title { color: #0f172a; font-weight: 800; font-size: 26px; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; margin-bottom: 20px;}
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .correct-txt { color: #16a34a; font-weight: bold; }
    .error-txt { color: #dc2626; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>📝 PHẦN MỀM TRỘN ĐỀ TRẮC NGHIỆM DTMIX</div>", unsafe_allow_html=True)

# ================= THANH CÔNG CỤ BÊN TRÁI (THÔNG TIN KỲ THI) =================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135692.png", width=80)
    st.markdown("### ⚙️ THÔNG TIN ĐỀ THI")
    so = st.text_input("Đơn vị (Sở/Phòng)", "SỞ GIÁO DỤC VÀ ĐÀO TẠO")
    truong = st.text_input("Trường", "TRƯỜNG THPT...")
    kythi = st.text_input("Kỳ thi", "KIỂM TRA HỌC KỲ II")
    namhoc = st.text_input("Năm học", "NĂM HỌC 2025 - 2026")
    monthi = st.text_input("Môn thi", "Môn: TIẾNG ANH")
    thoigian = st.text_input("Thời gian", "Thời gian làm bài: 45 phút")
    
    st.markdown("---")
    st.markdown("### 🛠️ TÙY CHỌN TRỘN")
    made_input = st.text_input("Mã đề (VD: 101, 102, 103, 104)", "101, 102, 103, 104")
    youngmix_mode = st.checkbox("🇬🇧 Chế độ Ngoại ngữ (Câu chùm <g3>, <g2>...)", value=False)

# ================= KHU VỰC CHÍNH (TABS NHƯ BẢN OFFLINE) =================
tab1, tab2 = st.tabs(["📊 RÀ SOÁT CẤU TRÚC ĐỀ GỐC", "🚀 XUẤT ĐỀ"])

with tab1:
    col_upload, col_action = st.columns([3, 1])
    with col_upload:
        uploaded_file = st.file_uploader("📂 Bấm hoặc kéo thả file đề gốc (.docx) vào đây để nạp dữ liệu", type=["docx"])
    
    if uploaded_file is not None:
        # Xử lý File 
        if "last_uploaded" not in st.session_state or st.session_state.last_uploaded != uploaded_file.name or st.session_state.ym_mode != youngmix_mode:
            with st.spinner("Đang quét cấu trúc đề gốc..."):
                temp_dir = tempfile.mkdtemp()
                temp_input_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Gọi Core xử lý
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

        # ================= GIAO DIỆN HIỂN THỊ DỮ LIỆU NHƯ OFFLINE =================
        parsed_data = st.session_state.parsed_data
        is_ym = st.session_state.ym_mode
        
        table_data = []
        total_q = 0
        error_count = 0

        # Trích xuất dữ liệu để đưa vào Bảng Grid
        if not is_ym:
            for p_idx, part in enumerate(parsed_data.get("parts", [])):
                p_type = part.get("type")
                part_name = f"Phần {('I', 'II', 'III', 'IV')[p_type-1] if p_type <= 4 else p_type}"
                for muc in part.get("mucs", []):
                    for q_idx, q in enumerate(muc.get("questions", [])):
                        total_q += 1
                        m = re.search(r'((?:Câu|Question)\s*\d+)', q.get('raw_text', ''), re.IGNORECASE)
                        q_name = m.group(1) if m else f"Câu {total_q}"
                        
                        ans_count = len(q.get("answers", []))
                        correct_ans_list = [chr(65+i) for i, a in enumerate(q.get("answers", [])) if a.get("is_true")]
                        correct_str = ", ".join(correct_ans_list) if correct_ans_list else "Trống"
                        
                        status = "✅ Hợp lệ"
                        if p_type == 1 and len(correct_ans_list) != 1:
                            status = "❌ Lỗi đáp án"
                            error_count += 1
                        
                        table_data.append({"Vị trí": part_name, "Câu hỏi": q_name, "Số P.Án": ans_count, "Đáp án đúng": correct_str, "Trạng thái": status})
        else:
            ym_data = st.session_state.youngmix_data
            for g_idx, group in enumerate(ym_data):
                g_name = f"Nhóm {g_idx + 1} ({group.get('type', '')})"
                for q in group.get("questions", []):
                    total_q += 1
                    m = re.search(r'((?:Câu|Question)\s*\d+)', q.get('raw_text', ''), re.IGNORECASE)
                    q_name = m.group(1) if m else f"Câu {total_q}"
                    
                    ans_count = len(q.get("answers", []))
                    correct_ans_list = [chr(65+i) for i, a in enumerate(q.get("answers", [])) if a.get("is_true")]
                    correct_str = ", ".join(correct_ans_list) if correct_ans_list else "Trống"
                    
                    status = "✅ Hợp lệ"
                    if ans_count == 0:
                        status = "⚠️ Cảnh báo (Tự luận?)"
                    
                    table_data.append({"Vị trí": g_name, "Câu hỏi": q_name, "Số P.Án": ans_count, "Đáp án đúng": correct_str, "Trạng thái": status})

        # HIỂN THỊ THỐNG KÊ (KPI)
        col1, col2, col3 = st.columns(3)
        col1.info(f"📚 **Tổng số câu hỏi:** {total_q}")
        col2.success(f"✅ **Hợp lệ:** {total_q - error_count}")
        if error_count > 0:
            col3.error(f"❌ **Lỗi phát hiện:** {error_count}")
        else:
            col3.success("✨ **Lỗi phát hiện:** 0")

        st.write("---")

        # HIỂN THỊ BẢNG LƯỚI BÊN TRÁI & NỘI DUNG ĐỀ BÊN PHẢI
        col_grid, col_preview = st.columns([1.5, 1])
        
        with col_grid:
            st.markdown("#### 📋 BẢNG RÀ SOÁT CẤU TRÚC")
            df = pd.DataFrame(table_data)
            # Hiển thị bảng dạng Dataframe tương tự Treeview của Tkinter
            st.dataframe(df, use_container_width=True, height=500, hide_index=True)
            
        with col_preview:
            st.markdown("#### 🔎 XEM TRƯỚC CHI TIẾT GỐC")
            with st.container(height=500, border=True):
                if not is_ym:
                    for p_idx, part in enumerate(parsed_data.get("parts", [])):
                        p_type = part.get("type")
                        st.markdown(f"**PHẦN {('I', 'II', 'III', 'IV')[p_type-1] if p_type <= 4 else p_type}**")
                        for muc in part.get("mucs", []):
                            for q in muc.get("questions", []):
                                st.write(f"*{q.get('raw_text', '')}*")
                                for a_idx, ans in enumerate(q.get("answers", [])):
                                    ans_text = ans.get('raw_text', ans.get('text', ''))
                                    prefix = chr(65+a_idx) + "."
                                    if ans.get("is_true"):
                                        st.markdown(f"<span class='correct-txt'>{prefix} {ans_text}</span>", unsafe_allow_html=True)
                                    else:
                                        st.write(f"{prefix} {ans_text}")
                                st.write("---")
                else:
                    ym_data = st.session_state.youngmix_data
                    for group in ym_data:
                        st.markdown(f"**[{group.get('type', '')}] {group.get('header', '')}**")
                        for q in group.get("questions", []):
                            st.write(f"*{q.get('raw_text', '')}*")
                            for a_idx, ans in enumerate(q.get("answers", [])):
                                ans_text = ans.get('raw_text', ans.get('text', ''))
                                prefix = chr(65+a_idx) + "."
                                if ans.get("is_true"):
                                    st.markdown(f"<span class='correct-txt'>{prefix} {ans_text}</span>", unsafe_allow_html=True)
                                else:
                                    st.write(f"{prefix} {ans_text}")
                            st.write("---")

with tab2:
    st.markdown("### 🚀 TIẾN HÀNH TRỘN ĐỀ")
    if "parsed_data" not in st.session_state:
        st.warning("⚠️ Vui lòng tải và rà soát file đề gốc ở Tab 1 trước khi trộn.")
    else:
        st.info("💡 Nếu bảng rà soát ở Tab 1 báo 'Hợp lệ' tất cả các câu, bạn có thể yên tâm xuất đề.")
        
        col_btn1, col_btn2 = st.columns([1, 2])
        with col_btn1:
            btn_tron = st.button("⚡ BẮT ĐẦU TRỘN ĐỀ", type="primary", use_container_width=True)
            
        if btn_tron:
            with st.spinner("Đang hoán vị câu hỏi và phương án..."):
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

                    zip_path = os.path.join(temp_dir, "De_Da_Tron")
                    shutil.make_archive(zip_path, 'zip', out_dir)

                    with open(zip_path + ".zip", "rb") as f:
                        st.session_state['download_data'] = f.read()
                        st.session_state['success'] = True
                except Exception as e:
                    st.error(f"❌ Có lỗi khi trộn: {e}")

        if st.session_state.get('success'):
            st.success("✅ Đã xuất đề xong! Hệ thống đã đóng gói thành file ZIP chứa Word và Excel.")
            st.download_button(
                label="💾 LƯU BỘ ĐỀ VÀO MÁY TÍNH (.ZIP)",
                data=st.session_state['download_data'],
                file_name="Bo_De_DTMIX_Online.zip",
                mime="application/zip",
                type="primary"
            )

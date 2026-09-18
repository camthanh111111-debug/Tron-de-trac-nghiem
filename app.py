import streamlit as st
import os
import tempfile
import shutil
import docx
import re
import pandas as pd
from dtmix_core import DTMIXApp 

# ================= CẤU HÌNH GIAO DIỆN =================
st.set_page_config(page_title="DTMIX - Trộn Đề Trắc Nghiệm", page_icon="📝", layout="wide")

st.markdown("""
    <style>
    .main-title { color: #0f172a; font-weight: 800; font-size: 26px; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; margin-bottom: 20px;}
    .correct-txt { color: #dc2626; font-weight: bold; background-color: #fee2e2; padding: 2px 5px; border-radius: 4px;} /* Màu đỏ cho đáp án đúng */
    .normal-txt { color: #334155; }
    .q-title { font-weight: bold; color: #1e40af; margin-top: 15px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>📝 PHẦN MỀM TRỘN ĐỀ TRẮC NGHIỆM DTMIX</div>", unsafe_allow_html=True)

# ================= THANH CÔNG CỤ BÊN TRÁI =================
with st.sidebar:
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

# ================= KHU VỰC CHÍNH (3 TABS) =================
tab1, tab2, tab3 = st.tabs(["📂 1. NẠP ĐỀ & BÁO CÁO", "🔎 2. XEM TRƯỚC CHI TIẾT", "🚀 3. XUẤT ĐỀ"])

with tab1:
    uploaded_file = st.file_uploader("📂 Bấm hoặc kéo thả file đề gốc (.docx) vào đây để nạp dữ liệu", type=["docx"])
    
    if uploaded_file is not None:
        if "last_uploaded" not in st.session_state or st.session_state.last_uploaded != uploaded_file.name or st.session_state.ym_mode != youngmix_mode:
            with st.spinner("Đang phân tích cấu trúc đề..."):
                temp_dir = tempfile.mkdtemp()
                temp_input_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
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

        # === BẢNG THÔNG BÁO TỪNG PHẦN ===
        st.markdown("### 📊 BÁO CÁO CẤU TRÚC ĐỀ THI")
        parsed_data = st.session_state.parsed_data
        is_ym = st.session_state.ym_mode
        
        summary_data = []
        total_q = 0
        total_errors = 0

        if not is_ym:
            for p_idx, part in enumerate(parsed_data.get("parts", [])):
                p_type = part.get("type")
                part_name = f"Phần {('I', 'II', 'III', 'IV')[p_type-1] if p_type <= 4 else p_type}"
                q_count = 0
                error_count = 0
                
                for muc in part.get("mucs", []):
                    for q in muc.get("questions", []):
                        q_count += 1
                        total_q += 1
                        # Kiểm tra đáp án
                        correct_ans = [a for a in q.get("answers", []) if a.get("is_true")]
                        if p_type == 1 and len(correct_ans) != 1:
                            error_count += 1
                            total_errors += 1
                            
                status = "✅ Đủ đáp án" if error_count == 0 else f"❌ Lỗi/Thiếu đáp án ở {error_count} câu"
                summary_data.append({"Khu vực": part_name, "Số câu hỏi": q_count, "Tình trạng": status})
        else:
            ym_data = st.session_state.youngmix_data
            for g_idx, group in enumerate(ym_data):
                g_name = f"Nhóm {g_idx + 1} ({group.get('type', 'Tự do')})"
                q_count = len(group.get("questions", []))
                total_q += q_count
                
                # Check thiếu phương án
                error_count = sum(1 for q in group.get("questions", []) if len(q.get("answers", [])) == 0)
                total_errors += error_count
                
                status = "✅ Đủ dữ liệu" if error_count == 0 else f"⚠️ {error_count} câu không có phương án"
                summary_data.append({"Khu vực": g_name, "Số câu hỏi": q_count, "Tình trạng": status})

        # Hiển thị tóm tắt
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)
        
        if total_errors > 0:
            st.error(f"⚠️ Phát hiện {total_errors} lỗi trong đề gốc. Vui lòng sang Tab 2 để xem chi tiết câu nào bị lỗi!")
        else:
            st.success(f"✨ Cấu trúc hoàn hảo! Đã nhận diện thành công {total_q} câu hỏi.")

with tab2:
    st.markdown("### 🔎 XEM TRƯỚC CHI TIẾT ĐỀ GỐC")
    if "parsed_data" not in st.session_state:
        st.info("👈 Vui lòng nạp file đề ở Tab 1 trước.")
    else:
        st.markdown("*Kiểm tra lại nội dung nhận diện. **Đáp án đúng** sẽ được bôi nền <span class='correct-txt'>màu đỏ</span>.*", unsafe_allow_html=True)
        st.write("---")
        
        parsed_data = st.session_state.parsed_data
        is_ym = st.session_state.ym_mode
        
        # Hiển thị cuộn rộng rãi
        with st.container(height=650, border=True):
            if not is_ym:
                for part in parsed_data.get("parts", []):
                    p_type = part.get("type")
                    st.markdown(f"#### 🏷️ PHẦN {('I', 'II', 'III', 'IV')[p_type-1] if p_type <= 4 else p_type}")
                    for muc in part.get("mucs", []):
                        for q in muc.get("questions", []):
                            # Tên câu hỏi và nội dung
                            st.markdown(f"<div class='q-title'>{q.get('raw_text', '')}</div>", unsafe_allow_html=True)
                            
                            # Hiển thị các phương án
                            ans_list = q.get("answers", [])
                            if not ans_list:
                                st.markdown("<span style='color:red;'>⚠️ (Không tìm thấy phương án/đáp án cho câu này)</span>", unsafe_allow_html=True)
                            else:
                                for a_idx, ans in enumerate(ans_list):
                                    prefix = chr(65+a_idx) + "." # A., B., C., D.
                                    ans_text = ans.get('raw_text', ans.get('text', ''))
                                    
                                    if ans.get("is_true"):
                                        st.markdown(f"<span class='correct-txt'>{prefix} {ans_text}</span>", unsafe_allow_html=True)
                                    else:
                                        st.markdown(f"<span class='normal-txt'>{prefix} {ans_text}</span>", unsafe_allow_html=True)
                            st.write("") # Dòng trống phân cách
            else:
                ym_data = st.session_state.youngmix_data
                for g_idx, group in enumerate(ym_data):
                    st.markdown(f"#### 🏷️ NHÓM {g_idx + 1} - Chế độ: {group.get('type', '')}")
                    if group.get('header'):
                        st.info(group.get('header'))
                    
                    for q in group.get("questions", []):
                        st.markdown(f"<div class='q-title'>{q.get('raw_text', '')}</div>", unsafe_allow_html=True)
                        ans_list = q.get("answers", [])
                        for a_idx, ans in enumerate(ans_list):
                            prefix = chr(65+a_idx) + "."
                            ans_text = ans.get('raw_text', ans.get('text', ''))
                            if ans.get("is_true"):
                                st.markdown(f"<span class='correct-txt'>{prefix} {ans_text}</span>", unsafe_allow_html=True)
                            else:
                                st.markdown(f"<span class='normal-txt'>{prefix} {ans_text}</span>", unsafe_allow_html=True)
                        st.write("")

with tab3:
    st.markdown("### 🚀 TIẾN HÀNH TRỘN ĐỀ")
    if "parsed_data" not in st.session_state:
        st.warning("⚠️ Vui lòng nạp file đề ở Tab 1 trước.")
    else:
        st.info("💡 Bấm nút dưới đây để phần mềm tạo ra các mã đề và đáp án (Word + Excel).")
        
        if st.button("⚡ BẮT ĐẦU TRỘN ĐỀ", type="primary"):
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
            st.success("✅ Đã xuất đề xong! Hệ thống đã đóng gói thành file ZIP.")
            st.download_button(
                label="💾 LƯU BỘ ĐỀ VÀO MÁY TÍNH (.ZIP)",
                data=st.session_state['download_data'],
                file_name="Bo_De_DTMIX_Online.zip",
                mime="application/zip",
                type="primary"
            )

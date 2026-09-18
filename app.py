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
    .st-emotion-cache-1y4p8pa { padding-top: 2rem; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-header'>✨ DTMIX ONLINE</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Công cụ nhận diện và trộn đề trắc nghiệm thông minh</p>", unsafe_allow_html=True)

# ================= THANH BÊN (SIDEBAR) THÔNG TIN =================
with st.sidebar:
    st.markdown("### 📝 THÔNG TIN ĐẦU TRANG")
    so = st.text_input("Sở GD&ĐT / Phòng", "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH")
    truong = st.text_input("Tên Trường", "TRƯỜNG THPT DƯƠNG BẠCH MAI")
    kythi = st.text_input("Tên Kỳ Thi", "KIỂM TRA HỌC KỲ II")
    namhoc = st.text_input("Năm Học", "NĂM HỌC 2025 - 2026")
    monthi = st.text_input("Môn Thi", "Môn: HÓA HỌC")
    thoigian = st.text_input("Thời Gian Làm Bài", "Thời gian làm bài: 45 phút")
    
    st.markdown("---")
    st.markdown("### ⚙️ CÀI ĐẶT TRỘN ĐỀ")
    made_input = st.text_input("Nhập mã đề (cách nhau dấu phẩy)", "101, 102, 103, 104")

# ================= KHU VỰC CHÍNH (3 TABS) =================
tab1, tab2, tab3 = st.tabs(["📂 1. TẢI ĐỀ GỐC", "🔍 2. RÀ SOÁT CẤU TRÚC", "🚀 3. XUẤT ĐỀ"])

with tab1:
    st.markdown("### ⚙️ TẢI LÊN TỆP ĐỀ GỐC")
    youngmix_mode = st.toggle("Bật chế độ nhóm dạng <g3>, <g2>... và đề Tiếng Anh", value=False)
    
    uploaded_file = st.file_uploader("Kéo thả file đề gốc (.docx) vào đây", type=["docx"])
    
    if uploaded_file is not None:
        # Xử lý phân tích file ngay khi vừa tải lên để chuyển sang Tab 2
        if "last_uploaded" not in st.session_state or st.session_state.last_uploaded != uploaded_file.name or st.session_state.ym_mode != youngmix_mode:
            with st.spinner("Đang đọc và phân tích cấu trúc đề..."):
                temp_dir = tempfile.mkdtemp()
                temp_input_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Gọi lõi DTMIX để phân tích (không trộn)
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
            
        st.success("✅ Đã tải và phân tích xong! Hãy chuyển sang Tab **2. RÀ SOÁT CẤU TRÚC** để kiểm tra lỗi.")

with tab2:
    st.markdown("### 📊 RÀ SOÁT ĐỀ GỐC")
    
    if "parsed_data" not in st.session_state:
        st.info("⚠️ Vui lòng tải file đề gốc lên ở Tab 1 trước.")
    else:
        parsed_data = st.session_state.parsed_data
        is_ym = st.session_state.ym_mode
        
        total_q = 0
        valid_q = 0
        error_q = 0
        warnings = []
        part_summaries = []

        # --- LOGIC RÀ SOÁT CHO CHẾ ĐỘ THƯỜNG ---
        if not is_ym:
            for p_idx, part in enumerate(parsed_data.get("parts", [])):
                p_type = part.get("type")
                part_name = f"Phần {('I', 'II', 'III', 'IV')[p_type-1] if p_type <= 4 else p_type}"
                p_count = 0
                
                for muc in part.get("mucs", []):
                    for q_idx, q in enumerate(muc.get("questions", [])):
                        total_q += 1
                        p_count += 1
                        
                        m = re.search(r'((?:Câu|Question)\s*\d+)', q.get('raw_text', ''), re.IGNORECASE)
                        q_name = m.group(1) if m else f"Câu {q_idx+1}"
                        
                        has_correct = False
                        if p_type == 1:
                            correct_count = sum(1 for a in q.get("answers", []) if a.get("is_true"))
                            if correct_count == 1: has_correct = True
                            elif correct_count > 1: warnings.append(f"❌ {part_name} - {q_name}: Có {correct_count} đáp án được đánh dấu đúng (Chỉ được phép 1).")
                            else: warnings.append(f"❌ {part_name} - {q_name}: Chưa gạch chân/tô đỏ đáp án đúng.")
                        elif p_type == 2:
                            has_correct = True 
                        elif p_type == 3:
                            has_correct = False
                            for elm in q.get('q_elements', []):
                                if elm.tag == docx.oxml.ns.qn('w:p'):
                                    txt = "".join([t.text for t in elm.iter(docx.oxml.ns.qn('w:t')) if t.text])
                                    if re.search(r'(?i)(đáp\s+án\s*[:\.]|(?:^|\n|\uFFFC)\s*A\.)\s*(.*)', txt):
                                        has_correct = True
                                        break
                            if not has_correct: warnings.append(f"❌ {part_name} - {q_name}: Thiếu chữ 'Đáp án: ...' (hoặc đáp án không chứa số).")
                        elif p_type == 4:
                            has_correct = True

                        if has_correct: valid_q += 1
                        else: error_q += 1
                part_summaries.append((part_name, p_count))
                
        # --- LOGIC RÀ SOÁT CHO CHẾ ĐỘ YOUNGMIX ---
        else:
            ym_data = st.session_state.youngmix_data
            for group in ym_data:
                q_in_group = group.get("q_count", 0)
                total_q += q_in_group
                valid_q += q_in_group # Ở chế độ YM tạm coi là hợp lệ, chi tiết sẽ báo ở warnings
                part_summaries.append((group["name"], q_in_group))
            
            # Quét lỗi cơ bản cho YM (Câu không có phương án)
            for part in parsed_data.get("parts", []):
                for muc in part.get("mucs", []):
                    for q_idx, q in enumerate(muc.get("questions", [])):
                        if len(q.get("answers", [])) == 0:
                            m = re.search(r'((?:Câu|Question)\s*\d+)', q.get('raw_text', ''), re.IGNORECASE)
                            q_name = m.group(1) if m else f"Câu {q_idx+1}"
                            warnings.append(f"⚠️ {q_name}: Không nhận diện được phương án (hoặc đây là câu tự luận/trả lời ngắn).")

        # --- HIỂN THỊ KPI ---
        c1, c2, c3 = st.columns(3)
        c1.metric("🔢 Tổng số câu hỏi", total_q)
        c2.metric("✅ Hợp lệ / Có đáp án", valid_q)
        c3.metric("⚠️ Câu lỗi / Cần kiểm tra", error_q if not is_ym else len(warnings))
        
        st.write("---")
        
        # --- HIỂN THỊ CHI TIẾT ---
        col_detail, col_error = st.columns([1, 1])
        with col_detail:
            st.markdown("**📋 CẤU TRÚC NHẬN DIỆN:**")
            for name, count in part_summaries:
                with st.expander(f"🔹 {name} ({count} câu)"):
                    st.write(f"Đã nhận diện thành công {count} câu hỏi thuộc phần này.")
                    
        with col_error:
            st.markdown("**🚨 BÁO CÁO LỖI:**")
            if not warnings and error_q == 0:
                st.success("Tuyệt vời! Không phát hiện lỗi định dạng nào trong đề gốc.")
            else:
                for w in warnings:
                    st.error(w)

with tab3:
    st.markdown("### 🚀 TRỘN ĐỀ VÀ XUẤT FILE")
    st.info("Hãy chắc chắn bạn đã rà soát lỗi ở Tab 2 trước khi tiến hành trộn đề.")
    
    if "parsed_data" not in st.session_state:
        st.warning("⚠️ Vui lòng tải file đề gốc lên ở Tab 1 trước.")
    else:
        if st.button("🚀 BẮT ĐẦU TRỘN ĐỀ", type="primary", use_container_width=True):
            with st.spinner("Đang tiến hành hoán vị và xuất file Word, Excel..."):
                try:
                    # Lấy lại đường dẫn đã lưu
                    temp_input_path = st.session_state.temp_input_path
                    temp_dir = st.session_state.temp_dir
                    out_dir = os.path.join(temp_dir, "KetQua")
                    os.makedirs(out_dir, exist_ok=True)

                    # Gọi code trộn
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

        if st.session_state.get('success'):
            st.success("✅ Trộn đề hoàn tất!")
            st.download_button(
                label="⬇️ BẤM VÀO ĐÂY ĐỂ TẢI BỘ ĐỀ ĐÃ TRỘN (.ZIP)",
                data=st.session_state['download_data'],
                file_name="Bo_De_DTMIX_Online.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

import streamlit as st
import os
import tempfile
import shutil
from dtmix_core import DTMIXApp 

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="DTMIX Online", page_icon="🚀", layout="wide")

# --- CSS TÙY CHỈNH CHO ĐẸP ---
st.markdown("""
    <style>
    .main-title { color: #1E3A8A; font-weight: bold; text-align: center; }
    .stButton>button { background-color: #F43F5E; color: white; font-weight: bold; border-radius: 8px; height: 50px; width: 100%; }
    .stButton>button:hover { background-color: #E11D48; }
    .card { background-color: #F8FAFC; padding: 20px; border-radius: 10px; border: 1px solid #E2E8F0; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-title'>✨ DTMIX ONLINE - TRỘN ĐỀ TRẮC NGHIỆM</h1>", unsafe_allow_html=True)
st.write("---")

# --- TẠO CÁC TAB GIAO DIỆN ---
tab1, tab2, tab3 = st.tabs(["⚙️ 1. Cài Đặt & Trộn Đề", "🛠️ 2. Tùy Chọn Nâng Cao", "📖 3. Hướng Dẫn Sử Dụng"])

with tab1:
    col_info, col_file = st.columns([1, 1], gap="large")
    
    with col_info:
        st.markdown("### 📝 THÔNG TIN ĐẦU TRANG ĐỀ")
        st.markdown("<div class='card'>", unsafe_allow_html=True)
        so = st.text_input("Sở GD&ĐT / Phòng", "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH")
        truong = st.text_input("Tên Trường", "TRƯỜNG THPT DƯƠNG BẠCH MAI")
        
        c1, c2 = st.columns(2)
        with c1:
            kythi = st.text_input("Tên Kỳ Thi", "KIỂM TRA HỌC KỲ II")
            monthi = st.text_input("Môn Thi", "Môn: HÓA HỌC")
        with c2:
            namhoc = st.text_input("Năm Học", "NĂM HỌC 2025 - 2026")
            thoigian = st.text_input("Thời Gian Làm Bài", "Thời gian làm bài: 45 phút")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_file:
        st.markdown("### 📁 TẢI LÊN & XUẤT ĐỀ")
        st.markdown("<div class='card'>", unsafe_allow_html=True)
        
        youngmix_mode = st.toggle("🚀 Bật chế độ Nhóm (<g3>, <g2>...) / Tiếng Anh", value=False)
        uploaded_file = st.file_uploader("Tải file đề gốc (.docx)", type=["docx"], help="Chỉ chấp nhận file định dạng Word (.docx)")
        
        st.write("---")
        st.markdown("**Cấu hình Mã đề:**")
        so_luong_de = st.selectbox("Số lượng đề cần trộn:", [str(i) for i in range(1, 25)], index=3)
        made_input = st.text_input("Nhập mã đề (Cách nhau dấu phẩy)", "101, 102, 103, 104")
        
        st.markdown("</div>", unsafe_allow_html=True)
        st.write("") # Tạo khoảng trống
        
        # NÚT TRỘN ĐỀ
        btn_tron = st.button("🚀 TRỘN ĐỀ VÀ XUẤT FILE")

with tab2:
    st.markdown("### 🛠️ CÁC TÙY CHỌN TRỘN ĐỀ")
    st.info("Các tùy chọn này áp dụng cho toàn bộ cấu trúc đề tải lên.")
    
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        continuous_numbering = st.checkbox("Đánh số câu liên tiếp (Cho tất cả các phần)", value=False)
        master_fix = st.checkbox("Cố định vị trí tất cả các nhóm/phần (Không hoán vị nhóm)", value=True)
    with col_opt2:
        keep_muc = st.checkbox("Giữ lại tiêu đề của Nhóm/Phần trong đề xuất ra", value=False)
        st.write("*Lưu ý: Tính năng Rà soát cấu trúc chi tiết (hiển thị từng câu, từng đáp án) trên Web đang được ẩn đi để tối ưu tốc độ. Đề vẫn sẽ được trộn theo đúng nguyên tắc nhận diện của lõi DTMIX.*")

with tab3:
    st.markdown("### 📖 HƯỚNG DẪN SỬ DỤNG DTMIX ONLINE")
    st.markdown("""
    **1. CHUẨN BỊ FILE GỐC (.docx):**
    * Soạn thảo bằng MS Word, lưu định dạng `.docx`.
    * **KHÔNG** sử dụng tính năng đánh số tự động (Numbering). Phải gõ thủ công như: 'Câu 1.', 'Question 1:', 'A.', 'B.', 'C.', 'D.'.
    * Phân chia các phần bằng từ khóa (viết hoa): `PHẦN I`, `PHẦN II`, `PHẦN III`, `PHẦN IV`.
    * Đáp án đúng ở Phần I và Phần II bắt buộc phải được **Gạch chân** (khuyến nghị) hoặc **tô đỏ**.
    * Cố định phương án hoặc câu: Gõ thêm dấu `#` liền trước đối tượng.

    **2. CHẾ ĐỘ NHÓM (<g3>, <g2>...):**
    * `<g0>`: Nhóm không trộn.
    * `<g1>`: Chỉ hoán vị câu hỏi.
    * `<g2>`: Chỉ hoán vị phương án.
    * `<g3>`: Hoán vị cả câu hỏi và phương án.
    * `<g4>`: Câu hỏi tự luận.
    """)

# ================= XỬ LÝ LOGIC TRỘN ĐỀ =================
if btn_tron:
    if not uploaded_file:
        st.error("❌ Vui lòng tải file đề gốc lên trước khi trộn!")
    else:
        with st.spinner("⏳ Hệ thống đang phân tích và trộn đề, vui lòng chờ..."):
            temp_dir = tempfile.mkdtemp()
            temp_input_path = os.path.join(temp_dir, uploaded_file.name)
            out_dir = os.path.join(temp_dir, "KetQua")
            os.makedirs(out_dir, exist_ok=True)

            with open(temp_input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            try:
                # Khởi tạo App và truyền CẢ các tùy chọn nâng cao vào
                app = DTMIXApp(
                    filepath=temp_input_path, out_dir=out_dir,
                    so=so, truong=truong, kythi=kythi, 
                    namhoc=namhoc, monthi=monthi, thoigian=thoigian,
                    youngmix=youngmix_mode, ma_de_str=made_input
                )
                
                # Gán các tùy chọn nâng cao từ Tab 2 vào lõi
                app.continuous_numbering_var.val = continuous_numbering
                app.master_fix_var.val = master_fix
                app.keep_muc_var.val = keep_muc
                
                # Bắt đầu trộn
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
    st.success("✅ Trộn đề hoàn tất! Bạn có thể tải file về bên dưới.")
    st.download_button(
        label="⬇️ TẢI BỘ ĐỀ ĐÃ TRỘN VÀ ĐÁP ÁN (.ZIP)",
        data=st.session_state['download_data'],
        file_name="Bo_De_DTMIX_Online.zip",
        mime="application/zip",
        type="primary"
    )

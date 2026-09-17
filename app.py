import streamlit as st
import os
import tempfile
import shutil
from dtmix_core import DTMIXApp 

st.set_page_config(page_title="DTMIX Online", layout="wide")
st.title("🚀 Phần Mềm Trộn Đề Trắc Nghiệm - DTMIX Online")

# Tạo 2 cột để nhập thông tin
col1, col2 = st.columns(2)
with col1:
    so = st.text_input("Sở GD&ĐT / Phòng", "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH")
    truong = st.text_input("Tên Trường", "TRƯỜNG THPT DƯƠNG BẠCH MAI")
    kythi = st.text_input("Tên Kỳ Thi", "KIỂM TRA HỌC KỲ II")
    namhoc = st.text_input("Năm Học", "NĂM HỌC 2025 - 2026")
with col2:
    monthi = st.text_input("Môn Thi", "Môn: HÓA HỌC")
    thoigian = st.text_input("Thời Gian Làm Bài", "Thời gian làm bài: 45 phút")
    made_input = st.text_input("Nhập mã đề (Cách nhau dấu phẩy)", "101, 102, 103, 104")
    youngmix_mode = st.checkbox("Bật chế độ Youngmix (<g3>, <g2>.../Tiếng Anh)", value=False)

uploaded_file = st.file_uploader("Tải file đề gốc (.docx)", type=["docx"])

# Xử lý logic khi bấm nút trộn đề
if uploaded_file and st.button("🚀 TRỘN ĐỀ VÀ XUẤT FILE", type="primary"):
    with st.spinner("Đang xử lý đề, vui lòng chờ trong giây lát..."):
        temp_dir = tempfile.mkdtemp()
        temp_input_path = os.path.join(temp_dir, uploaded_file.name)
        out_dir = os.path.join(temp_dir, "KetQua")
        os.makedirs(out_dir, exist_ok=True)

        with open(temp_input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            app = DTMIXApp(
                filepath=temp_input_path,
                out_dir=out_dir,
                so=so, truong=truong, kythi=kythi, 
                namhoc=namhoc, monthi=monthi, thoigian=thoigian,
                youngmix=youngmix_mode,
                ma_de_str=made_input
            )
            
            app.process_web()

            zip_path = os.path.join(temp_dir, "De_Da_Tron")
            shutil.make_archive(zip_path, 'zip', out_dir)

            # Lưu file vào session_state để nút Download không bị mất khi refresh trang
            with open(zip_path + ".zip", "rb") as f:
                st.session_state['download_data'] = f.read()
                st.session_state['success'] = True
                
        except Exception as e:
            st.error(f"Có lỗi xảy ra trong quá trình trộn: {e}")

# Đưa nút Tải về ra ngoài. Nó sẽ hiện ra khi tiến trình báo 'success'
if st.session_state.get('success'):
    st.success("✅ Trộn đề thành công! Hãy bấm nút bên dưới để lưu file về máy.")
    st.download_button(
        label="⬇️ BẤM VÀO ĐÂY ĐỂ TẢI BỘ ĐỀ (.ZIP)",
        data=st.session_state['download_data'],
        file_name="Bo_De_DTMIX.zip",
        mime="application/zip"
    )
import os
from datetime import date

import pandas as pd
import streamlit as st


EXCEL_FILE = "ho_so_nhan_vien.xlsx"


def load_data():
    """Đọc dữ liệu nhân viên từ file Excel (nếu có)."""
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
        except Exception:
            # Nếu file bị lỗi định dạng thì tạo mới
            df = pd.DataFrame()
    else:
        df = pd.DataFrame()
    return df


def append_employee_to_excel(record: dict):
    """Thêm 1 bản ghi nhân viên vào cuối file Excel."""
    df_existing = load_data()

    df_new = pd.DataFrame([record])
    if df_existing.empty:
        df_final = df_new
    else:
        # Căn chỉnh cột để tránh lỗi nếu thêm trường mới
        df_final = pd.concat([df_existing, df_new], ignore_index=True)

    # Ghi lại ra file Excel
    df_final.to_excel(EXCEL_FILE, index=False)


def setup_page():
    st.set_page_config(
        page_title="QUẢN LÝ NHÂN VIÊN",
        page_icon="👨‍💼",
        layout="centered",
    )

    # CSS giao diện nền đen, chữ sáng
    dark_css = """
        <style>
        body {
            background-color: #111111;
            color: #ffffff;
        }
        .stApp {
            background-color: #111111;
            color: #ffffff;
        }
        header, .st-emotion-cache-18ni7ap, .st-emotion-cache-1avcm0n {
            background-color: #111111 !important;
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        .title-box {
            padding: 1rem 1.5rem;
            border-radius: 0.5rem;
            background: linear-gradient(135deg, #1f1f1f, #2a2a2a);
            border: 1px solid #333333;
            color: #ffffff;
            text-align: center;
        }
        .title-box h1 {
            font-size: 1.8rem;
            margin-bottom: 0.25rem;
            color: #ffffff;
            font-weight: 700;
        }
        .title-box p {
            margin: 0;
            font-size: 0.9rem;
            color: #f5f5f5;
            font-weight: 600;
        }
        .field-box {
            padding: 1rem 1.25rem;
            border-radius: 0.5rem;
            background-color: #181818;
            border: 1px solid #333333;
            margin-bottom: 1rem;
        }
        /* Nhãn (label) của các ô nhập liệu */
        label, .stMarkdown, .data-table-title {
            color: #ffffff !important;
            font-weight: 600 !important;
        }
        .stTextInput > div > div > input,
        .stNumberInput input,
        .stDateInput input,
        .stSelectbox > div > div > select,
        .stTextArea textarea {
            background-color: #101010 !important;
            color: #f0f0f0 !important;
            border-radius: 0.4rem;
            border: 1px solid #444444;
        }
        .stTextInput > div > div > input:focus,
        .stNumberInput input:focus,
        .stDateInput input:focus,
        .stSelectbox > div > div > select:focus,
        .stTextArea textarea:focus {
            border-color: #6c63ff !important;
            box-shadow: 0 0 0 1px #6c63ff33;
        }
        .stButton > button {
            background: linear-gradient(135deg, #6c63ff, #4a3fe4);
            color: #ffffff;
            border-radius: 999px;
            border: none;
            padding: 0.5rem 1.5rem;
            font-weight: 600;
        }
        .stButton > button:hover {
            background: linear-gradient(135deg, #7d74ff, #5b50ff);
        }
        .success-box {
            padding: 0.75rem 1rem;
            border-radius: 0.5rem;
            background-color: #220909;
            border: 1px solid #ff4d4d;
            color: #ff4d4d;
            font-size: 0.9rem;
            font-weight: 700;
        }
        .data-table-title {
            margin-top: 1.5rem;
            margin-bottom: 0.25rem;
            font-weight: 600;
        }
        </style>
    """
    st.markdown(dark_css, unsafe_allow_html=True)


def main():
    setup_page()

    st.markdown(
        """
        <div class="title-box">
            <h1>Nhập liệu hồ sơ nhân viên</h1>
            <p>Lưu trữ hồ sơ trực tiếp vào file Excel trên máy của bạn</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")

    # Form nhập liệu
    with st.form("employee_form"):
        st.markdown('<div class="field-box">', unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        with col1:
            ma_nv = st.text_input("Mã nhân viên")
            ho_ten = st.text_input("Họ và tên")
            ngay_sinh = st.date_input("Ngày sinh", value=date(1990, 1, 1))
            gioi_tinh = st.selectbox("Giới tính", ["Nam", "Nữ", "Khác"])

        with col2:
            phong_ban = st.text_input("Phòng ban")
            chuc_vu = st.text_input("Chức vụ")
            so_dien_thoai = st.text_input("Số điện thoại")
            email = st.text_input("Email")

        dia_chi = st.text_area("Địa chỉ", height=80)
        ngay_vao_lam = st.date_input("Ngày vào làm", value=date.today())
        luong_co_ban = st.number_input(
            "Lương cơ bản (VNĐ)",
            min_value=0.0,
            step=100000.0,
            format="%.0f",
        )

        st.markdown("</div>", unsafe_allow_html=True)

        submitted = st.form_submit_button("Lưu hồ sơ vào Excel")

    if submitted:
        # Kiểm tra các trường bắt buộc
        required_fields = {
            "Mã nhân viên": ma_nv,
            "Họ và tên": ho_ten,
            "Phòng ban": phong_ban,
        }
        missing = [k for k, v in required_fields.items() if str(v).strip() == ""]

        if missing:
            st.error(
                "Vui lòng nhập đầy đủ thông tin bắt buộc: "
                + ", ".join(missing)
            )
        else:
            record = {
                "Mã nhân viên": ma_nv,
                "Họ và tên": ho_ten,
                "Ngày sinh": ngay_sinh,
                "Giới tính": gioi_tinh,
                "Phòng ban": phong_ban,
                "Chức vụ": chuc_vu,
                "Số điện thoại": so_dien_thoai,
                "Email": email,
                "Địa chỉ": dia_chi,
                "Ngày vào làm": ngay_vao_lam,
                "Lương cơ bản": luong_co_ban,
            }

            try:
                append_employee_to_excel(record)
                st.markdown(
                    '<div class="success-box"><span style="color:#ff4d4d; font-weight:700;">✅ Đã lưu '
                    f'file <strong>{EXCEL_FILE}</strong> trong thư mục hiện tại.</span></div>',
                    unsafe_allow_html=True,
                )
            except Exception as e:
                st.error(f"Không thể ghi vào file Excel: {e}")

    # Hiển thị dữ liệu hiện có trong file Excel (nếu có)
    df_current = load_data()
    if not df_current.empty:
        st.markdown(
            '<p class="data-table-title">Danh sách hồ sơ nhân viên hiện tại:</p>',
            unsafe_allow_html=True,
        )
        st.dataframe(df_current, use_container_width=True)


if __name__ == "__main__":
    main()



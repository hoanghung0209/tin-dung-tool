import os
import io
import json
import re
from datetime import datetime
import pandas as pd
import streamlit as st

# Import từ các file core mà chúng ta đã xây dựng
from config import NOICAP_OPTIONS, LOAN_TYPE_OPTIONS, LONG_TEXT_HINTS
from core.utils import (
    vi_title_name, parse_cccd_payload, validate_cccd_12_digits,
    var_name, is_cccd_desc, is_noicap_desc, is_loan_type_desc, 
    is_name_desc, is_dob_desc, is_issue_desc, is_addr_desc, is_gender_desc,
    format_money_vi
)
from core.scanner import decode_qr_offline
from core.document import read_mapping, generate_word_document

# --- CẤU HÌNH GIAO DIỆN WEB ---
st.set_page_config(page_title="Hồ Sơ Tín Dụng", page_icon="🏦", layout="wide")
st.title("🏦 HỆ THỐNG KHỞI TẠO HỒ SƠ TÍN DỤNG (WEB)")
st.markdown("**HOÀNG VIỆT HƯNG - QTDND PHÙNG HƯNG**")

# Khởi tạo bộ nhớ tạm (Session State) để không bị mất chữ khi web tải lại
if "form_data" not in st.session_state:
    st.session_state.form_data = {}

# ==========================================
# SIDEBAR: KHU VỰC TẢI DỮ LIỆU & QUÉT QR
# ==========================================
with st.sidebar:
    st.header("⚙️ CẤU HÌNH ĐẦU VÀO")
    excel_file = st.file_uploader("1. Tải file Data (Excel)", type=["xlsx", "xls"])
    docx_file = st.file_uploader("2. Tải file Mẫu (Word)", type=["docx"])
    
    st.divider()
    st.header("📸 QUÉT CCCD TỰ ĐỘNG")
    qr_file = st.file_uploader("Tải ảnh CCCD để tự động điền", type=["png", "jpg", "jpeg"])
    if qr_file and st.button("Tiến hành quét mã", type="primary"):
        with open("temp_qr.jpg", "wb") as f:
            f.write(qr_file.getbuffer())
        try:
            info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
            if not validate_cccd_12_digits(info.get("CCCD", "")):
                st.warning("CCCD quét được không đủ 12 số, hãy kiểm tra lại!")
            
            # Đổ dữ liệu quét được vào bộ nhớ tạm để tự hiện lên form
            for key, val in info.items():
                st.session_state.form_data[f"qr_{key}"] = val
            st.success("Quét thành công! Dữ liệu đã được nạp vào form.")
        except Exception as e:
            st.error(f"Không thể đọc mã QR: {e}")

# ==========================================
# KHU VỰC CHÍNH: RENDER FORM TỪ EXCEL
# ==========================================
if excel_file and docx_file:
    # Lưu file mẫu ra nháp để xử lý
    with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
    with open("temp_template.docx", "wb") as f: f.write(docx_file.getbuffer())
    
    try:
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
    except Exception as e:
        st.error(f"Lỗi đọc file Excel: {e}")
        st.stop()

    # Xây dựng hệ thống Tab
    unique_tabs = list(dict.fromkeys(tabs_dict.values()))
    st_tabs = st.tabs(unique_tabs)
    tab_objs = dict(zip(unique_tabs, st_tabs))

    idx = 1
    current_row_cols = None
    col_idx = 0

    st.markdown("---")
    st.subheader("📝 NHẬP LIỆU CHI TIẾT HỒ SƠ")

    # Vòng lặp vẽ giao diện Tetris-Grid 2 cột
    for ph, ds in mapping.items():
        d = str(ds).lower()
        tn = tabs_dict.get(ph, "Thông tin chung")
        is_tall = any(k in d for k in LONG_TEXT_HINTS)
        
        # Tiêm dữ liệu từ QR (nếu có)
        default_val = st.session_state.form_data.get(ph, "")
        if is_cccd_desc(d): default_val = st.session_state.form_data.get("qr_CCCD", default_val)
        elif is_name_desc(d): default_val = vi_title_name(st.session_state.form_data.get("qr_Ho_va_ten", default_val))
        elif is_dob_desc(d): default_val = st.session_state.form_data.get("qr_Ngay_thang_nam_sinh", default_val)
        elif is_addr_desc(d): default_val = st.session_state.form_data.get("qr_Dia_chi", default_val)
        elif is_issue_desc(d): default_val = st.session_state.form_data.get("qr_Ngay_cap_CCCD", default_val)
        elif is_gender_desc(d): default_val = st.session_state.form_data.get("qr_Gioi_tinh", default_val)

        label_text = f"{idx}. {ds}"

        with tab_objs[tn]:
            if is_tall:
                # Textbox dài: Rớt dòng, chiếm trọn màn hình
                st.session_state.form_data[ph] = st.text_area(label_text, value=default_val, height=100, key=f"widget_{ph}")
                current_row_cols = None 
                col_idx = 0
            else:
                # Các ô ngắn: Xếp zic-zắc 2 cột
                if current_row_cols is None or col_idx == 2:
                    current_row_cols = st.columns(2)
                    col_idx = 0
                
                with current_row_cols[col_idx]:
                    if is_noicap_desc(d) or is_loan_type_desc(d):
                        options = NOICAP_OPTIONS if is_noicap_desc(d) else LOAN_TYPE_OPTIONS
                        st.session_state.form_data[ph] = st.selectbox(label_text, options=[""] + options, index=0, key=f"widget_{ph}")
                    else:
                        ft_val = str(field_types.get(ph, "")).lower()
                        if ft_val.startswith("dropdown:"):
                            opts = [o.strip() for o in ft_val.split(":")[1].split(",")]
                            st.session_state.form_data[ph] = st.selectbox(label_text, options=[""] + opts, index=0, key=f"widget_{ph}")
                        else:
                            st.session_state.form_data[ph] = st.text_input(label_text, value=default_val, key=f"widget_{ph}")
                col_idx += 1
        idx += 1

    st.markdown("---")
    
    # Nút Xuất File
    if st.button("🚀 TẠO HỒ SƠ (WORD)", type="primary", use_container_width=True):
        # 1. Thu thập và tự động format tiền tệ trước khi xuất
        ctx = {}
        for ph, ds in mapping.items():
            raw_val = str(st.session_state.form_data.get(ph, "")).strip()
            
            # Format tự động cho CCCD và Tiền tệ
            d = str(ds).lower()
            if is_cccd_desc(d): 
                raw_val = "".join(ch for ch in raw_val if ch.isdigit())
            
            # Mắt thần định dạng tiền (Gõ 100000 -> Tự xuất 100.000)
            money_kws = ["doanh thu", "thu nhập", "tổng thu", "lợi nhuận", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "tài sản", "dư nợ", "lãi", "giá"]
            is_money_field = "money" in str(field_types.get(ph, "")).lower() or any(kw in d for kw in money_kws)
            if is_money_field and raw_val.isdigit():
                raw_val = format_money_vi(raw_val)

            ctx[var_name(ph)] = raw_val

        # 2. Sinh mã Auto-ID
        name = ctx.get(var_name(next((p for p, d in mapping.items() if is_name_desc(d.lower())), "")), "")
        if name:
            initials = "".join([word[0].upper() for word in name.split() if word])
            auto_id = f"{initials}-{datetime.now().strftime('%d%m%y')}-001"
            for ph, ds in mapping.items():
                if "mã" in ds.lower() and ("hồ sơ" in ds.lower() or "hscv" in ds.lower()):
                    ctx[var_name(ph)] = auto_id

        # 3. Tạo file Word trên RAM (Không lưu xuống cứng)
        try:
            output_stream = io.BytesIO()
            generate_word_document("temp_template.docx", output_stream, ctx)
            output_stream.seek(0)
            
            member_name = safe_filename(name) or "HoSoTinDung"
            file_name = f"{member_name}_{datetime.now().strftime('%H%M')}.docx"
            
            st.success("🎉 Tạo hồ sơ thành công! Bấm nút bên dưới để tải về máy.")
            st.download_button(
                label="📥 TẢI XUỐNG FILE WORD",
                data=output_stream,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
        except Exception as e:
            st.error(f"Lỗi tạo file: {e}")

else:
    st.info("👈 Hãy tải lên file Data (Excel) và file Mẫu (Word) ở thanh Sidebar bên trái để bắt đầu.")
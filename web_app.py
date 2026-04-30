import os
import io
import json
import re
import logging
from datetime import datetime
import pandas as pd
import streamlit as st
from docxtpl import RichText

# Import từ các file core
from config import NOICAP_OPTIONS, LOAN_TYPE_OPTIONS, LONG_TEXT_HINTS
from core.utils import (
    vi_title_name, parse_cccd_payload, validate_cccd_12_digits,
    var_name, is_cccd_desc, is_noicap_desc, is_loan_type_desc, 
    is_name_desc, is_dob_desc, is_issue_desc, is_addr_desc, is_gender_desc,
    format_money_vi, parse_number_vi
)
from core.scanner import decode_qr_offline
from core.document import read_mapping, generate_word_document

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")

# --- CSS LÀM ĐẸP GIAO DIỆN ---
st.markdown("""
    <style>
    .stTextInput input, .stTextArea textarea, .stSelectbox select {
        border: 1px solid #2980B9 !important;
        border-radius: 5px !important;
    }
    div[data-testid="stExpander"] {
        border: 1px solid #27AE60 !important;
        background-color: #F4F9F4;
    }
    .stButton button {
        border-radius: 20px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- KHỞI TẠO DỮ LIỆU ---
if "mapping" not in st.session_state:
    st.session_state.mapping = {}
if "field_types" not in st.session_state:
    st.session_state.field_types = {}
if "tabs_dict" not in st.session_state:
    st.session_state.tabs_dict = {}

# ==========================================
# CÁC HÀM XỬ LÝ DỮ LIỆU TỰ ĐỘNG
# ==========================================

def fill_qr_to_form(info):
    """Điền dữ liệu từ CCCD vào các ô tương ứng"""
    for ph, ds in st.session_state.mapping.items():
        d = str(ds).lower()
        key = f"input_{ph}"
        if is_cccd_desc(d): st.session_state[key] = info.get("CCCD", "")
        elif is_name_desc(d): st.session_state[key] = vi_title_name(info.get("Ho_va_ten", ""))
        elif is_dob_desc(d): st.session_state[key] = info.get("Ngay_thang_nam_sinh", "")
        elif is_addr_desc(d): st.session_state[key] = info.get("Dia_chi", "")
        elif is_issue_desc(d): st.session_state[key] = info.get("Ngay_cap_CCCD", "")
        elif is_gender_desc(d): st.session_state[key] = info.get("Gioi_tinh", "")

def fill_plan_to_form(ma_pa, df_pa_data):
    """Điền dữ liệu phương án vào các ô TLCP, HMTN..."""
    # Đã sửa lỗi .str.lower() tại đây
    mask = df_pa_data['Mã phương án'].astype(str).str.strip().str.lower() == str(ma_pa).strip().lower()
    plan_rows = df_pa_data[mask]
    
    df_cp = plan_rows[plan_rows['Loại'].astype(str).str.contains('chi|phí|cp', case=False, na=False)]
    df_tn = plan_rows[plan_rows['Loại'].astype(str).str.contains('lợi|nhuận|thu|nhập|ln|tn', case=False, na=False)]

    for ph in st.session_state.mapping.keys():
        k_var = var_name(ph).lower()
        key = f"input_{ph}"
        
        # Xử lý Chi phí
        for i, (_, row) in enumerate(df_cp.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.')
            if k_var == f"tlcp{i}": st.session_state[key] = rate
            if k_var == f"hmcp{i}": st.session_state[key] = str(row['Hạng mục'])
            if k_var == f"ndcp{i}": st.session_state[key] = str(row['Nội dung chi tiết'])

        # Xử lý Thu nhập
        for i, (_, row) in enumerate(df_tn.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.')
            if k_var == f"tltn{i}": st.session_state[key] = rate
            if k_var == f"hmtn{i}": st.session_state[key] = str(row['Hạng mục'])
            if k_var == f"ndtn{i}": st.session_state[key] = str(row['Nội dung chi tiết'])

# ==========================================
# SIDEBAR: UPLOAD & SCANNER
# ==========================================
with st.sidebar:
    st.title("🏦 CẤU HÌNH HỆ THỐNG")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. File Mẫu Word", type=["docx"])
    pa_file = st.file_uploader("3. File Phương án", type=["xlsx"])
    
    st.divider()
    st.header("📸 QUÉT MÃ CCCD")
    qr_img = st.file_uploader("Tải ảnh chụp CCCD", type=["png", "jpg", "jpeg"])
    if qr_img and st.button("TIẾN HÀNH QUÉT MÃ", type="primary", use_container_width=True):
        if not st.session_state.mapping:
            st.error("Hãy tải file Data.xlsx lên trước khi quét!")
        else:
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                fill_qr_to_form(info)
                st.success("Đã quét và điền dữ liệu thành công!")
            except Exception as e: st.error(f"Lỗi: {e}")

# ==========================================
# CHƯƠNG TRÌNH CHÍNH
# ==========================================
if excel_file and docx_file:
    # Nạp dữ liệu Excel vào session
    with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
    with open("temp_template.docx", "wb") as f: f.write(docx_file.getbuffer())
    
    m, ft, tb = read_mapping("temp_data.xlsx")
    st.session_state.mapping = m
    st.session_state.field_types = ft
    st.session_state.tabs_dict = tb

    # Dropdown Phương án
    if pa_file:
        df_pa_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_pa_data = pd.read_excel(pa_file, sheet_name="data")
        pa_map = dict(zip(df_pa_list['Tên PA'], df_pa_list['Mã PA']))
        
        col_pa1, col_pa2 = st.columns([3, 1])
        with col_pa1:
            selected_pa_name = st.selectbox("⚡ ĐIỀN TỰ ĐỘNG PHƯƠNG ÁN:", ["-- Chọn phương án --"] + list(pa_map.keys()))
        with col_pa2:
            if st.button("ÁP DỤNG", type="primary", use_container_width=True):
                if selected_pa_name != "-- Chọn phương án --":
                    fill_plan_to_form(pa_map[selected_pa_name], df_pa_data)
                    st.toast(f"Đã áp dụng phương án: {selected_pa_name}")

    # Khởi tạo Tabs
    unique_tabs = list(dict.fromkeys(st.session_state.tabs_dict.values()))
    st_tabs = st.tabs(unique_tabs)
    tab_map = dict(zip(unique_tabs, st_tabs))
    
    # Render Form với Grid 2 cột
    tab_occ = {t: set() for t in unique_tabs}
    tab_row = {t: 0 for t in unique_tabs}

    for ph, ds in st.session_state.mapping.items():
        d = str(ds).lower()
        tn = st.session_state.tabs_dict.get(ph, "Thông tin chung")
        key = f"input_{ph}"
        
        # Đảm bảo key luôn tồn tại trong session state để không bị lỗi value
        if key not in st.session_state: st.session_state[key] = ""
        
        with tab_map[tn]:
            r, c = tab_row[tn], 0
            while (r, c) in tab_occ[tn]:
                c += 1
                if c > 1: c=0; r+=1
            tab_row[tn] = r
            
            is_tall = any(k in d for k in ["địa chỉ", "nội dung", "mục đích", "phương án", "tài sản"])
            rs = 2 if is_tall else 1
            for _r in range(r, r + rs): tab_occ[tn].add((_r, c))

            if c == 0: cols = st.columns(2)
            
            with cols[c]:
                if is_tall:
                    st.text_area(ds, key=key, height=110)
                elif is_noicap_desc(d) or is_loan_type_desc(d):
                    opts = NOICAP_OPTIONS if is_noicap_desc(d) else LOAN_TYPE_OPTIONS
                    st.selectbox(ds, [""] + opts, key=key)
                else:
                    st.text_input(ds, key=key)

    # NÚT XUẤT FILE WORD
    st.divider()
    if st.button("🚀 XUẤT HỒ SƠ WORD TỔNG HỢP", type="primary", use_container_width=True):
        ctx = {}
        for ph, ds in st.session_state.mapping.items():
            val = str(st.session_state.get(f"input_{ph}", "")).strip()
            
            # Auto-format tiền tệ
            money_kws = ["doanh thu", "thu nhập", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "lãi", "hạn mức"]
            if any(k in ds.lower() for k in money_kws) and val.replace(".","").isdigit():
                val = format_money_vi(val.replace(".",""))
            
            # Xử lý xuống dòng cho các ô nhiều tài sản
            if "\n" in val:
                rt = RichText()
                lines = val.split("\n")
                for i, line in enumerate(lines):
                    rt.add(line)
                    if i < len(lines) - 1: rt.add_line_break()
                ctx[var_name(ph)] = rt
            else:
                ctx[var_name(ph)] = val

        # Sinh mã Auto-ID
        name_key = next((p for p, d in st.session_state.mapping.items() if is_name_desc(d.lower())), None)
        if name_key:
            name = st.session_state.get(f"input_{name_key}", "")
            if name:
                initials = "".join([w[0].upper() for word in name.split() if w])
                auto_id = f"{initials}-{datetime.now().strftime('%d%m%y')}-001"
                for ph, ds in st.session_state.mapping.items():
                    if "mã" in ds.lower() and "hồ sơ" in ds.lower(): ctx[var_name(ph)] = auto_id

        # Tạo file
        out = io.BytesIO()
        generate_word_document("temp_template.docx", out, ctx)
        out.seek(0)
        st.success("Đã tạo hồ sơ hoàn tất!")
        st.download_button("📥 TẢI FILE WORD (.DOCX)", data=out, file_name=f"HSCV_{datetime.now().strftime('%H%M')}.docx")

else:
    st.info("💡 HƯỚNG DẪN: Tải file Data.xlsx và Mẫu Word ở Sidebar bên trái để bắt đầu nhập liệu.")

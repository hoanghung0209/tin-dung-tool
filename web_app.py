import os
import io
import json
import re
import logging
from datetime import datetime
import pandas as pd
import streamlit as st

# Import từ các file core
from config import NOICAP_OPTIONS, LOAN_TYPE_OPTIONS, LONG_TEXT_HINTS
from core.utils import (
    vi_title_name, parse_cccd_payload, validate_cccd_12_digits,
    var_name, is_cccd_desc, is_noicap_desc, is_loan_type_desc, 
    is_name_desc, is_dob_desc, is_issue_desc, is_addr_desc, is_gender_desc,
    format_money_vi, parse_number_vi, parse_percent_vi
)
from core.scanner import decode_qr_offline
from core.document import read_mapping, generate_word_document

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")

# --- CSS TÙY CHỈNH ĐỂ GIAO DIỆN SẮC NÉT HƠN ---
st.markdown("""
    <style>
    .stTextInput input, .stTextArea textarea, .stSelectbox select {
        border: 1px solid #979DA2 !important;
    }
    [data-testid="stExpander"] {
        border: 1px solid #2980B9 !important;
        background-color: #F8F9F9;
    }
    </style>
    """, unsafe_allow_width=True)

# --- KHỞI TẠO SESSION STATE (BỘ NHỚ TẠM) ---
if "form_data" not in st.session_state:
    st.session_state.form_data = {}
if "last_plan" not in st.session_state:
    st.session_state.last_plan = None

# ==========================================
# SIDEBAR: CẤU HÌNH & FILE
# ==========================================
with st.sidebar:
    st.title("⚙️ HỆ THỐNG LÕI")
    excel_file = st.file_uploader("1. File Data (Excel)", type=["xlsx"])
    docx_file = st.file_uploader("2. Mẫu Word (.docx)", type=["docx"])
    pa_file = st.file_uploader("3. Danh mục Phương án (Excel)", type=["xlsx"])
    
    st.divider()
    st.header("📸 QUÉT CCCD")
    qr_file = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
    if qr_file and st.button("🚀 TIẾN HÀNH QUÉT", type="primary"):
        with open("temp_qr.jpg", "wb") as f: f.write(qr_file.getbuffer())
        try:
            info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
            st.session_state.form_data["qr_CCCD"] = info.get("CCCD", "")
            st.session_state.form_data["qr_Ho_va_ten"] = info.get("Ho_va_ten", "")
            st.session_state.form_data["qr_Ngay_thang_nam_sinh"] = info.get("Ngay_thang_nam_sinh", "")
            st.session_state.form_data["qr_Dia_chi"] = info.get("Dia_chi", "")
            st.session_state.form_data["qr_Ngay_cap_CCCD"] = info.get("Ngay_cap_CCCD", "")
            st.session_state.form_data["qr_Gioi_tinh"] = info.get("Gioi_tinh", "")
            st.success("Đã nạp dữ liệu QR!")
            st.rerun()
        except Exception as e: st.error(f"Lỗi: {e}")

# ==========================================
# CHỨC NĂNG TỰ ĐỘNG FILL PHƯƠNG ÁN
# ==========================================
def apply_plan_data(selected_pa, df_pa_data, mapping):
    try:
        mask = df_pa_data['Mã phương án'].astype(str).str.strip().lower() == selected_pa.lower()
        plan_rows = df_pa_data[mask]
        
        # Reset các ô tài chính cũ
        for k in list(st.session_state.form_data.keys()):
            if any(x in k.lower() for x in ["cp", "tn", "hm", "nd"]):
                st.session_state.form_data[k] = ""

        df_cp = plan_rows[plan_rows['Loại'].astype(str).str.contains('chi|phí|cp', case=False, na=False)]
        df_tn = plan_rows[plan_rows['Loại'].astype(str).str.contains('lợi|nhuận|thu|nhập|ln|tn', case=False, na=False)]

        # Điền Tỷ lệ Chi phí
        for i, (_, row) in enumerate(df_cp.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.')
            for ph in mapping.keys():
                k_var = var_name(ph).lower()
                if k_var == f"tlcp{i}": st.session_state[f"w_{ph}"] = rate
                if k_var == f"hmcp{i}": st.session_state[f"w_{ph}"] = str(row['Hạng mục'])

        # Điền Tỷ lệ Thu nhập
        for i, (_, row) in enumerate(df_tn.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.')
            for ph in mapping.keys():
                k_var = var_name(ph).lower()
                if k_var == f"tltn{i}": st.session_state[f"w_{ph}"] = rate
                if k_var == f"hmtn{i}": st.session_state[f"w_{ph}"] = str(row['Hạng mục'])
    except: pass

# ==========================================
# HIỂN THỊ FORM CHÍNH
# ==========================================
if excel_file and docx_file:
    with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
    with open("temp_template.docx", "wb") as f: f.write(docx_file.getbuffer())
    
    mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
    
    # Khu vực chọn Phương án
    if pa_file:
        df_pa_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_pa_data = pd.read_excel(pa_file, sheet_name="data")
        pa_options = dict(zip(df_pa_list['Tên PA'], df_pa_list['Mã PA']))
        
        selected_plan_name = st.selectbox("⚡ CHỌN PHƯƠNG ÁN VAY VỐN:", ["-- Chọn phương án --"] + list(pa_options.keys()))
        
        if selected_plan_name != "-- Chọn phương án --" and selected_plan_name != st.session_state.last_plan:
            apply_plan_data(pa_options[selected_plan_name], df_pa_data, mapping)
            st.session_state.last_plan = selected_plan_name
            st.rerun()

    # Tạo Tabs
    unique_tabs = list(dict.fromkeys(tabs_dict.values()))
    st_tabs = st.tabs(unique_tabs)
    tab_map = dict(zip(unique_tabs, st_tabs))
    
    # Thuật toán Grid 2 cột
    tab_occupancy = {t: set() for t in unique_tabs}
    tab_row = {t: 0 for t in unique_tabs}

    for ph, ds in mapping.items():
        d = str(ds).lower()
        tn = tabs_dict.get(ph, "Thông tin chung")
        is_tall = any(k in d for k in ["địa chỉ", "nội dung", "mục đích", "phương án", "tài sản"])
        
        # Xử lý dữ liệu mặc định từ QR hoặc Session
        val = st.session_state.get(f"w_{ph}", "")
        if is_cccd_desc(d): val = st.session_state.form_data.get("qr_CCCD", val)
        elif is_name_desc(d): val = vi_title_name(st.session_state.form_data.get("qr_Ho_va_ten", val))
        elif is_dob_desc(d): val = st.session_state.form_data.get("qr_Ngay_thang_nam_sinh", val)
        elif is_addr_desc(d): val = st.session_state.form_data.get("qr_Dia_chi", val)

        with tab_map[tn]:
            # Tìm vị trí trống trong lưới
            r, c = tab_row[tn], 0
            while (r, c) in tab_occupancy[tn]:
                c += 1
                if c > 1: c=0; r+=1
            tab_row[tn] = r
            
            rs = 2 if is_tall else 1
            for _r in range(r, r + rs): tab_occupancy[tn].add((_r, c))

            # Render UI
            if c == 0: cols = st.columns(2) # Chỉ tạo column mới khi ở cột đầu
            
            with st.container(): # Dùng container để bọc các widget cùng cột
                if is_tall:
                    st.session_state[f"w_{ph}"] = st.text_area(f"{ds}", value=val, height=110, key=f"input_{ph}")
                elif is_noicap_desc(d) or is_loan_type_desc(d):
                    opts = NOICAP_OPTIONS if is_noicap_desc(d) else LOAN_TYPE_OPTIONS
                    st.session_state[f"w_{ph}"] = st.selectbox(f"{ds}", [""] + opts, key=f"input_{ph}")
                else:
                    st.session_state[f"w_{ph}"] = st.text_input(f"{ds}", value=val, key=f"input_{ph}")

    # ==========================================
    # NÚT XUẤT FILE & LOGIC TÀI CHÍNH
    # ==========================================
    st.divider()
    if st.button("🚀 XUẤT HỒ SƠ WORD", type="primary", use_container_width=True):
        ctx = {}
        # Thu thập dữ liệu và định dạng tiền tệ
        for ph, ds in mapping.items():
            raw = str(st.session_state.get(f"input_{ph}", "")).strip()
            d_low = ds.lower()
            
            # Format tiền tự động cho các trường tài chính
            money_kws = ["doanh thu", "thu nhập", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "lãi"]
            if any(kw in d_low for kw in money_kws) and raw.replace(".","").isdigit():
                raw = format_money_vi(raw.replace(".",""))
            
            ctx[var_name(ph)] = raw

        # Sinh Auto-ID
        name_key = next((p for p, d in mapping.items() if is_name_desc(d.lower())), None)
        if name_key:
            name = st.session_state.get(f"input_{name_key}", "")
            initials = "".join([w[0].upper() for word in name.split() if w])
            auto_id = f"{initials}-{datetime.now().strftime('%d%m%y')}-001"
            for ph, ds in mapping.items():
                if "mã" in ds.lower() and "hồ sơ" in ds.lower(): ctx[var_name(ph)] = auto_id

        # Render Word
        output = io.BytesIO()
        generate_word_document("temp_template.docx", output, ctx)
        output.seek(0)
        st.success("Đã tạo file thành công!")
        st.download_button("📥 TẢI FILE WORD VỀ MÁY", data=output, file_name=f"HoSo_{datetime.now().strftime('%H%M')}.docx")

else:
    st.info("💡 Vui lòng tải đủ 2 file Data.xlsx và HSCV.docx ở cột bên trái để bắt đầu làm việc.")

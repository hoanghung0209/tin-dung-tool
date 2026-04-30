import os
import io
import re
import pandas as pd
from datetime import datetime
import streamlit as st
from docxtpl import DocxTemplate, RichText

# Import core
from config import NOICAP_OPTIONS, LOAN_TYPE_OPTIONS, LONG_TEXT_HINTS
from core.utils import (
    vi_title_name, parse_cccd_payload, validate_cccd_12_digits,
    var_name, is_cccd_desc, is_noicap_desc, is_loan_type_desc, 
    is_name_desc, is_dob_desc, is_issue_desc, is_addr_desc, is_gender_desc,
    format_money_vi, parse_number_vi
)
from core.scanner import decode_qr_offline
from core.document import read_mapping

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")

# ==========================================
# HÀM BỔ TRỢ TÀI CHÍNH (FIX TRIỆT ĐỂ)
# ==========================================
def get_num(val):
    if not val: return 0
    digits = "".join(c for c in str(val) if c.isdigit())
    return int(digits) if digits else 0

def run_financial_formulas(mapping):
    """Logic: Vay + Tự có = Tổng vốn | LS Năm / 12 = LS Tháng"""
    # 1. Tính Tổng Vốn
    vay_k = next((p for p in mapping if var_name(p).lower() in ["tienvay", "sotienvay", "sovonvay"]), None)
    tc_k = next((p for p in mapping if var_name(p).lower() in ["vontuco", "von_tu_co"]), None)
    tvdt_k = next((p for p in mapping if var_name(p).lower() in ["tvdt", "tongvon", "tongvondautu"]), None)
    
    if vay_k and tc_k and tvdt_k:
        total = get_num(st.session_state.get(vay_k, 0)) + get_num(st.session_state.get(tc_k, 0))
        if total > 0: st.session_state[tvdt_k] = f"{total:,.0f}".replace(",", ".")

    # 2. Tính Lãi Suất
    lsn_k = next((p for p in mapping if var_name(p).lower() in ["laisuatnam", "ls_nam"]), None)
    lst_k = next((p for p in mapping if var_name(p).lower() in ["laisuatthang", "ls_thang"]), None)
    if lsn_k and lst_k:
        val_lsn = str(st.session_state.get(lsn_k, "0")).replace(",", ".")
        try:
            num_lsn = float(val_lsn)
            if num_lsn > 0: st.session_state[lst_k] = str(round(num_lsn / 12, 4)).replace(".", ",")
        except: pass

def calculate_plan_money(mapping):
    """Nhân tỷ lệ % với Tổng vốn để ra số tiền từng mục"""
    tvdt_k = next((p for p in mapping if var_name(p).lower() in ["tvdt", "tongvon", "tongvondautu"]), None)
    total_val = get_num(st.session_state.get(tvdt_k, 0))
    if total_val <= 0: return

    for ph in mapping.keys():
        k_var = var_name(ph).lower()
        # Tìm các ô Số tiền Chi phí/Thu nhập (cp1, tn1, stcp1, sttn1...)
        m = re.match(r"^(?:stcp|sttn|cp|tn|stdt|dt)(\d+)$", k_var)
        if m:
            idx = m.group(1)
            type_prefix = "tlcp" if "cp" in k_var else "tltn"
            rate_k = next((p for p in mapping if var_name(p).lower() in [f"{type_prefix}{idx}", f"t{type_prefix[2:]}{idx}"]), None)
            if rate_k and st.session_state.get(rate_k):
                try:
                    rate = float(str(st.session_state[rate_k]).replace(",", "."))
                    st.session_state[ph] = f"{int(total_val * rate / 100):,.0f}".replace(",", ".")
                except: pass

# ==========================================
# BỘ LỌC QUÉT QR (4 NHÓM ĐÍCH DANH)
# ==========================================
def is_target_match_strict(target, d_low):
    if target == "Tất cả": return True
    t_low = target.lower()
    if t_low not in d_low: return False
    # Chống ghi đè cho Thành viên
    if t_low == "thành viên":
        if any(ex in d_low for ex in ["đồng vay", "ủy quyền", "vợ", "chồng"]): return False
    if t_low == "người đồng vay vốn 1" and "2" in d_low: return False
    if t_low == "người ủy quyền 1" and "2" in d_low: return False
    return True

def process_field_change(ph, d_lower, field_types, mapping):
    val = str(st.session_state.get(ph, ""))
    # Format tiền
    if any(k in d_lower for k in ["doanh thu", "thu nhập", "chi phí", "số tiền", "vốn", "giá trị"]):
        digits = "".join(c for c in val if c.isdigit())
        if digits: st.session_state[ph] = f"{int(digits):,}".replace(",", ".")
    # Format tên
    if is_name_desc(d_lower) and val: st.session_state[ph] = vi_title_name(val)
    # Chạy công thức
    run_financial_formulas(mapping)
    calculate_plan_money(mapping)

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
with st.sidebar:
    st.header("⚙️ CẤU HÌNH")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    pa_file = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    if excel_file:
        with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        for ph in mapping: # Khởi tạo session state để không mất dữ liệu khi chuyển tab
            if ph not in st.session_state: st.session_state[ph] = ""
            
        st.divider()
        st.header("📸 QUÉT CCCD")
        target = st.selectbox("Đối tượng:", ["Thành viên", "Người đồng vay vốn 1", "Người ủy quyền 1", "Người ủy quyền 2"])
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        if qr_img and st.button("🚀 TIẾN HÀNH QUÉT", type="primary", use_container_width=True):
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                for ph, ds in mapping.items():
                    if is_target_match_strict(target, ds.lower()):
                        if is_cccd_desc(ds.lower()): st.session_state[ph] = info.get("CCCD", "")
                        elif is_name_desc(ds.lower()): st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif is_dob_desc(ds.lower()): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif is_addr_desc(ds.lower()): st.session_state[ph] = info.get("Dia_chi", "")
                st.success(f"Xong: {target}"); st.rerun()
            except Exception as e: st.error(f"Lỗi: {e}")

if excel_file and docx_file:
    # 1. Chọn Phương án
    if pa_file:
        df_pa_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_pa_data = pd.read_excel(pa_file, sheet_name="data")
        pa_map = dict(zip(df_pa_list['Tên PA'], df_pa_list['Mã PA']))
        cp1, cp2 = st.columns([3, 1])
        with cp1: sel_pa = st.selectbox("⚡ PHƯƠNG ÁN:", ["-- Trống --"] + list(pa_map.keys()))
        with cp2:
            st.write(""); 
            if st.button("ÁP DỤNG", type="primary", use_container_width=True):
                if sel_pa != "-- Trống --":
                    ma = pa_map[sel_pa]
                    rows = df_pa_data[df_pa_data['Mã phương án'].astype(str).str.strip().str.lower() == str(ma).lower()]
                    df_cp = rows[rows['Loại'].astype(str).str.contains('chi|phí|cp', case=False, na=False)]
                    df_tn = rows[rows['Loại'].astype(str).str.contains('thu|nhập|tn|dt|doanh', case=False, na=False)]
                    for ph, ds in mapping.items():
                        kv = var_name(ph).lower()
                        if "tên phương án" in ds.lower(): st.session_state[ph] = sel_pa
                        # Fill Chi phí
                        for i, (_, r) in enumerate(df_cp.iterrows(), 1):
                            if re.match(fr"^(?:tlcp|tcp){i}$", kv): st.session_state[ph] = str(round(float(r['Tỉ lệ'])*100,2))
                            if re.match(fr"^(?:hmcp|hm_cp){i}$", kv): st.session_state[ph] = str(r['Hạng mục'])
                            if re.match(fr"^(?:ndcp|nd_cp){i}$", kv): st.session_state[ph] = str(r['Nội dung chi tiết'])
                        # Fill Thu nhập
                        for i, (_, r) in enumerate(df_tn.iterrows(), 1):
                            if re.match(fr"^(?:tltn|ttn){i}$", kv): st.session_state[ph] = str(round(float(r['Tỉ lệ'])*100,2))
                            if re.match(fr"^(?:hmtn|hm_tn){i}$", kv): st.session_state[ph] = str(r['Hạng mục'])
                            if re.match(fr"^(?:ndtn|nd_tn){i}$", kv): st.session_state[ph] = str(r['Nội dung chi tiết'])
                    calculate_plan_money(mapping); st.rerun()

    # 2. Hiển thị Form theo Tabs
    st.divider()
    unique_tabs = list(dict.fromkeys(tabs_dict.values()))
    st_tabs = st.tabs(unique_tabs)
    for i, tab_name in enumerate(unique_tabs):
        with st_tabs[i]:
            f_in_tab = [(ph, ds) for ph, ds in mapping.items() if tabs_dict.get(ph) == tab_name]
            c1, c2 = st.columns(2)
            for j, (ph, ds) in enumerate(f_in_tab):
                target_col = c1 if j % 2 == 0 else c2
                with target_col:
                    kw = {"ph": ph, "d_lower": ds.lower(), "field_types": field_types, "mapping": mapping}
                    if any(k in ds.lower() for k in LONG_TEXT_HINTS):
                        st.text_area(ds, key=ph, height=120, on_change=process_field_change, kwargs=kw)
                    elif is_noicap_desc(ds.lower()) or is_loan_type_desc(ds.lower()):
                        st.selectbox(ds, [""] + (NOICAP_OPTIONS if is_noicap_desc(ds.lower()) else LOAN_TYPE_OPTIONS), key=ph, on_change=process_field_change, kwargs=kw)
                    else:
                        st.text_input(ds, key=ph, on_change=process_field_change, kwargs=kw)

    # 3. Xuất Word
    if st.button("🚀 XUẤT HỒ SƠ WORD", type="primary", use_container_width=True):
        ctx = {}
        for ph, ds in mapping.items():
            val = str(st.session_state.get(ph, "")).strip()
            if "\n" in val:
                rt = RichText()
                for line in val.split("\n"): rt.add(line); rt.add_line_break()
                ctx[var_name(ph)] = rt
            else: ctx[var_name(ph)] = val
        from core.document import generate_word_document
        out = io.BytesIO()
        generate_word_document("temp_template.docx", out, ctx); out.seek(0)
        st.download_button("📥 TẢI FILE WORD", data=out, file_name=f"HSCV_{datetime.now().strftime('%H%M')}.docx", type="primary")

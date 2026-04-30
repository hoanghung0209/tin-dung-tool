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

# ==========================================
# HÀM TÍNH TOÁN TỰ ĐỘNG (CORE CALCULATION)
# ==========================================
def run_financial_logic(mapping):
    """Tính toán Lãi suất và Tổng vốn đầu tư"""
    # 1. Tính Tổng vốn đầu tư = Tiền vay + Vốn tự có
    vay_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["tienvay", "sotienvay", "sovonthexin"]), None)
    tu_co_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["vontuco", "von_tu_co"]), None)
    tvdt_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["tvdt", "tongvon", "tongvondautu"]), None)

    if vay_key and tu_co_key and tvdt_key:
        val_vay = parse_number_vi(st.session_state[vay_key])
        val_tu_co = parse_number_vi(st.session_state[tu_co_key])
        if val_vay > 0 or val_tu_co > 0:
            st.session_state[tvdt_key] = format_money_vi(val_vay + val_tu_co)

    # 2. Tính Lãi suất tháng = Lãi suất năm / 12
    ls_nam_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["laisuatnam", "ls_nam"]), None)
    ls_thang_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["laisuatthang", "ls_thang"]), None)

    if ls_nam_key and ls_thang_key:
        val_ls_nam = str(st.session_state[ls_nam_key]).replace(",", ".")
        if val_ls_nam and re.match(r"^\d*\.?\d*$", val_ls_nam):
            val_thang = round(float(val_ls_nam) / 12, 4)
            st.session_state[ls_thang_key] = str(val_thang).replace(".", ",")

def process_field_change(ph, d_lower, field_types, mapping):
    """Xử lý định dạng và kích hoạt tính toán khi rời ô nhập liệu"""
    val = str(st.session_state[ph])
    digits = "".join(c for c in val if c.isdigit())
    ft = str(field_types.get(ph, "")).lower()

    # Định dạng tiền tệ
    money_kws = ["doanh thu", "thu nhập", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "lãi", "hạn mức"]
    if ("money" in ft or any(kw in d_lower for kw in money_kws)) and digits:
        st.session_state[ph] = f"{int(digits):,}".replace(",", ".")
        # Dịch số thành chữ (Spell)
        if "spell:" in ft:
            from num2words import num2words
            target = re.split(r"^spell\s*:", ft, flags=re.IGNORECASE)[1].strip()
            if target in st.session_state:
                txt = num2words(int(digits), lang="vi").replace("-", " ")
                st.session_state[target] = txt[0].upper() + txt[1:] + " đồng chẵn."

    # Kích hoạt tính toán tài chính
    run_financial_logic(mapping)

def is_target_match_strict(target, d_low):
    """Bộ lọc đối tượng QR: Chống ghi đè chéo"""
    if target == "Tất cả": return True
    d_low = d_low.lower()
    t_low = target.lower()
    
    if t_low not in d_low: return False
    
    # Nếu là Thành viên, loại trừ các ô của người đi kèm
    if t_low == "thành viên":
        for ex in ["đồng vay", "ủy quyền", "vợ", "chồng", "bảo lãnh"]:
            if ex in d_low: return False
    # Nếu là Người ủy quyền 1, loại trừ Người ủy quyền 2
    if t_low == "người ủy quyền 1" and "người ủy quyền 2" in d_low: return False
    
    return True

# ==========================================
# UI & APP FLOW
# ==========================================
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")
st.title("🏦 HỆ THỐNG KHỞI TẠO HỒ SƠ TÍN DỤNG (BẢN FIX)")

with st.sidebar:
    st.header("⚙️ CẤU HÌNH")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    pa_file = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    st.divider()
    st.header("📸 QUÉT CCCD ĐÍCH DANH")
    if excel_file:
        with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        
        target = st.selectbox("Đối tượng quét:", ["Thành viên", "Đồng vay vốn", "Người ủy quyền 1", "Người ủy quyền 2", "Tất cả"])
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        
        if qr_img and st.button("🚀 QUÉT & ĐIỀN", type="primary", use_container_width=True):
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                for ph, ds in mapping.items():
                    d_low = ds.lower()
                    if is_target_match_strict(target, d_low):
                        if is_cccd_desc(d_low): st.session_state[ph] = info.get("CCCD", "")
                        elif is_name_desc(d_low): st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif is_dob_desc(d_low): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif is_addr_desc(d_low): st.session_state[ph] = info.get("Dia_chi", "")
                        elif is_issue_desc(d_low): st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                st.success(f"Đã nạp xong cho: {target}"); st.rerun()
            except Exception as e: st.error(f"Lỗi: {e}")

if excel_file and docx_file:
    # Khởi tạo keys
    for ph in mapping.keys():
        if ph not in st.session_state: st.session_state[ph] = ""

    # Tab Rendering
    unique_tabs = list(dict.fromkeys(tabs_dict.values()))
    st_tabs = st.tabs(unique_tabs)
    tab_obj = dict(zip(unique_tabs, st_tabs))

    for tab_name in unique_tabs:
        with tab_obj[tab_name]:
            fields = [(ph, ds) for ph, ds in mapping.items() if tabs_dict.get(ph) == tab_name]
            col_l, col_r = st.columns(2)
            cols = [col_l, col_r]; c_idx = 0
            for ph, ds in fields:
                d_low = str(ds).lower()
                kwargs = {"ph": ph, "d_lower": d_low, "field_types": field_types, "mapping": mapping}
                with cols[c_idx]:
                    if is_noicap_desc(d_low) or is_loan_type_desc(d_low):
                        st.selectbox(ds, [""] + (NOICAP_OPTIONS if is_noicap_desc(d_low) else LOAN_TYPE_OPTIONS), key=ph, on_change=process_field_change, kwargs=kwargs)
                    elif any(k in d_low for k in LONG_TEXT_HINTS):
                        st.text_area(ds, key=ph, height=120, on_change=process_field_change, kwargs=kwargs)
                    else:
                        st.text_input(ds, key=ph, on_change=process_field_change, kwargs=kwargs)
                c_idx = 1 - c_idx

    # Nút Xuất File
    st.divider()
    if st.button("🚀 XUẤT HỒ SƠ WORD", type="primary", use_container_width=True):
        ctx = {}
        for ph, ds in mapping.items():
            val = str(st.session_state[ph]).strip()
            if "\n" in val:
                rt = RichText()
                for i, line in enumerate(val.split("\n")):
                    rt.add(line); rt.add_line_break() if i < len(val.split("\n"))-1 else None
                ctx[var_name(ph)] = rt
            else: ctx[var_name(ph)] = val
        
        from core.document import generate_word_document
        out = io.BytesIO()
        generate_word_document("temp_template.docx", out, ctx); out.seek(0)
        st.download_button("📥 TẢI FILE WORD", data=out, file_name=f"HSCV_{datetime.now().strftime('%H%M')}.docx", type="primary")

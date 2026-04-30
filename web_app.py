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
# CÔNG CỤ TÍNH TOÁN & ĐỊNH DẠNG (FIX LỖI)
# ==========================================
def number_to_vn_money(num_str):
    try:
        from num2words import num2words
        txt = num2words(int(num_str), lang="vi").replace("-", " ")
        return txt[0].upper() + txt[1:] + " đồng chẵn."
    except: return ""

def calculate_total_and_plan(mapping):
    """Tính Tổng vốn đầu tư và Rót tiền phương án (Sửa lỗi logic)"""
    # 1. Tính Tổng Vốn
    vay_key = next((p for p in mapping if var_name(p).lower() in ["tienvay", "sotienvay", "tongvay", "sotienxinvay"]), None)
    tc_key = next((p for p in mapping if var_name(p).lower() in ["vontuco", "von_tu_co", "tuco"]), None)
    tvdt_key = next((p for p in mapping if var_name(p).lower() in ["tvdt", "tongvon", "tongvondautu", "tongmucdautu"]), None)
    
    total_val = 0
    if vay_key and tc_key and tvdt_key:
        v_vay = parse_number_vi(str(st.session_state.get(vay_key, "0")))
        v_tc = parse_number_vi(str(st.session_state.get(tc_key, "0")))
        if v_vay > 0 or v_tc > 0:
            total_val = v_vay + v_tc
            st.session_state[tvdt_key] = f"{int(total_val):,}".replace(",", ".")
    
    if total_val == 0 and tvdt_key:
        total_val = parse_number_vi(str(st.session_state.get(tvdt_key, "0")))

    # 2. Nhân tiền Phương án
    if total_val > 0:
        for ph in mapping.keys():
            k_var = var_name(ph).lower()
            
            # Tiền Chi phí
            m_cp = re.match(r"^(?:stcp|st_cp|sotiencp|cp|chiphi|tien_cp)(\d+)$", k_var)
            if m_cp:
                idx = m_cp.group(1)
                rate_keys = [p for p in mapping if re.match(fr"^(?:tlcp|tcp|t\.cp|tylecp){idx}$", var_name(p).lower())]
                if rate_keys and st.session_state.get(rate_keys[0]):
                    try:
                        rate = float(str(st.session_state[rate_keys[0]]).replace(",", "."))
                        st.session_state[ph] = f"{int(total_val * rate / 100):,}".replace(",", ".")
                    except: pass
            
            # Tiền Thu nhập (Sửa lỗi không hiển thị)
            m_tn = re.match(r"^(?:sttn|st_tn|sotientn|tn|thunhap|tien_tn)(\d+)$", k_var)
            if m_tn:
                idx = m_tn.group(1)
                rate_keys = [p for p in mapping if re.match(fr"^(?:tltn|ttn|t\.tn|tyletn){idx}$", var_name(p).lower())]
                if rate_keys and st.session_state.get(rate_keys[0]):
                    try:
                        rate = float(str(st.session_state[rate_keys[0]]).replace(",", "."))
                        st.session_state[ph] = f"{int(total_val * rate / 100):,}".replace(",", ".")
                    except: pass

# ==========================================
# CALLBACKS: XỬ LÝ SỰ KIỆN GÕ PHÍM (ON_CHANGE)
# ==========================================
def process_field_change(ph, d_lower, field_types, mapping):
    val = str(st.session_state[ph])
    ft = str(field_types.get(ph, "")).lower()
    
    is_rate = "percent" in ft or any(k in d_lower for k in ["lãi suất", "tỷ lệ", "laisuat"])
    is_money = "money" in ft or "spell:" in ft or any(k in d_lower for k in ["doanh thu", "thu nhập", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "hạn mức"])

    # 1. XỬ LÝ LÃI SUẤT (Bảo tồn dấu phẩy)
    if is_rate:
        clean_val = "".join(c for c in val if c.isdigit() or c in ",.")
        if clean_val:
            st.session_state[ph] = clean_val # Giữ nguyên định dạng 12,6
            if "năm" in d_lower or "lsnam" in var_name(ph).lower() or "laisuatnam" in var_name(ph).lower():
                thang_key = next((p for p, d in mapping.items() if "tháng" in d.lower() and ("lãi suất" in d.lower() or "ls" in var_name(p).lower())), None)
                if thang_key:
                    try:
                        num_val = float(clean_val.replace(",", "."))
                        st.session_state[thang_key] = str(round(num_val / 12, 4)).replace(".", ",")
                    except: pass

    # 2. XỬ LÝ TIỀN TỆ (Chỉ lấy số, format chấm phẩy)
    elif is_money:
        digits = "".join(c for c in val if c.isdigit())
        if digits:
            st.session_state[ph] = f"{int(digits):,}".replace(",", ".")
            if "spell:" in ft:
                target = re.split(r"^spell\s*:", ft, flags=re.IGNORECASE)[1].strip()
                if target in st.session_state: st.session_state[target] = number_to_vn_money(digits)
    
    # 3. VIẾT HOA TÊN & TẠO AUTO-ID
    elif is_name_desc(d_lower) and val:
        st.session_state[ph] = vi_title_name(val)
        initials = "".join([w[0].upper() for w in val.split() if w])
        auto_id = f"{initials}-{datetime.now().strftime('%d%m%y')}-001"
        for p, d in mapping.items():
            if "mã" in d.lower() and "hồ sơ" in d.lower(): st.session_state[p] = auto_id

    # KÍCH HOẠT TÍNH TỔNG VỐN & TIỀN PHƯƠNG ÁN
    calculate_total_and_plan(mapping)

# ==========================================
# CALLBACK: ÁP DỤNG PHƯƠNG ÁN (FULL REGEX)
# ==========================================
def apply_plan_callback(selected_pa, ma_pa, df_pa_data, mapping):
    mask = df_pa_data['Mã phương án'].astype(str).str.strip().str.lower() == str(ma_pa).strip().lower()
    plan_rows = df_pa_data[mask]
    
    df_cp = plan_rows[plan_rows['Loại'].astype(str).str.contains('chi|phí|cp', case=False, na=False)]
    df_tn = plan_rows[plan_rows['Loại'].astype(str).str.contains('lợi|nhuận|thu|nhập|ln|tn', case=False, na=False)]

    for ph, ds in mapping.items():
        k_var = var_name(ph).lower()
        if "tên phương án" in ds.lower() or k_var in ["tenphuongan", "ten_pa"]:
            st.session_state[ph] = selected_pa

        # Rót Chi phí
        for i, (_, row) in enumerate(df_cp.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
            hm = str(row.get('Hạng mục', ''))
            nd = str(row.get('Nội dung chi tiết', ''))
            
            if re.match(fr"^(?:tlcp|tcp|t\.cp|tylecp){i}$", k_var): st.session_state[ph] = rate
            if re.match(fr"^(?:hmcp|hm_cp|hangmuccp){i}$", k_var): st.session_state[ph] = hm
            if re.match(fr"^(?:ndcp|nd_cp|noidungcp|nd_chiphi|chiphi|chi_phi){i}$", k_var): 
                st.session_state[ph] = nd if not hm else (f"{hm}: {nd}" if k_var.startswith("chi") else nd)

        # Rót Thu nhập (Sửa lỗi mù chữ)
        for i, (_, row) in enumerate(df_tn.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
            hm = str(row.get('Hạng mục', ''))
            nd = str(row.get('Nội dung chi tiết', ''))
            
            if re.match(fr"^(?:tltn|ttn|t\.tn|tyletn){i}$", k_var): st.session_state[ph] = rate
            if re.match(fr"^(?:hmtn|hm_tn|hangmuctn){i}$", k_var): st.session_state[ph] = hm
            if re.match(fr"^(?:ndtn|nd_tn|noidungtn|nd_thunhap|thunhap|thu_nhap){i}$", k_var): 
                st.session_state[ph] = nd if not hm else (f"{hm}: {nd}" if k_var.startswith("thu") else nd)

    # Tính toán lại số tiền ngay sau khi nạp tỷ lệ
    calculate_total_and_plan(mapping)

# ==========================================
# ĐÍCH DANH BỘ LỌC QUÉT QR
# ==========================================
def is_target_match_strict(target, d_low):
    if target == "Tất cả": return True
    t_low = target.lower()
    if t_low not in d_low: return False
    
    # Ràng buộc chống ghi đè cho Thành viên
    if t_low == "thành viên":
        if any(ex in d_low for ex in ["đồng vay", "ủy quyền", "vợ", "chồng", "bảo lãnh", "thụ hưởng"]): 
            return False
    # Ràng buộc Ủy quyền 1 không ăn vào Ủy quyền 2
    if t_low == "người ủy quyền 1" and "người ủy quyền 2" in d_low: return False
    return True

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")
st.markdown("<style>div[data-baseweb='input'] > div { border: 1px solid #2980B9; border-radius:4px;}</style>", unsafe_allow_html=True)

st.title("🏦 HỆ THỐNG KHỞI TẠO HỒ SƠ TÍN DỤNG")
st.caption("PHIÊN BẢN CHỐNG GHI ĐÈ & TÍNH TOÁN LIÊN KẾT - QTDND PHÙNG HƯNG")

with st.sidebar:
    st.header("⚙️ TẢI FILE CẤU HÌNH")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    pa_file = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    st.divider()
    st.header("📸 QUÉT CCCD ĐÍCH DANH")
    
    if excel_file:
        with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        
        # Danh sách cố định (Chuẩn)
        subjects = ["Tất cả", "Thành viên", "Đồng vay vốn", "Vợ thành viên", "Chồng thành viên", "Người ủy quyền 1", "Người ủy quyền 2"]
        target_person = st.selectbox("Quét thông tin cho:", subjects)
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        
        if qr_img and st.button("🚀 TIẾN HÀNH QUÉT", type="primary", use_container_width=True):
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                for ph, ds in mapping.items():
                    d_low = ds.lower()
                    if is_target_match_strict(target_person, d_low):
                        if is_cccd_desc(d_low): st.session_state[ph] = info.get("CCCD", "")
                        elif is_name_desc(d_low): st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif is_dob_desc(d_low): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif is_addr_desc(d_low): st.session_state[ph] = info.get("Dia_chi", "")
                        elif is_issue_desc(d_low): st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                st.success(f"Đã nạp chính xác cho: {target_person}")
                st.rerun()
            except Exception as e: st.error(f"Lỗi: {e}")

if excel_file and docx_file:
    for ph in mapping.keys():
        if ph not in st.session_state: st.session_state[ph] = ""

    # KHU VỰC PHƯƠNG ÁN TỰ ĐỘNG
    if pa_file:
        df_pa_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_pa_data = pd.read_excel(pa_file, sheet_name="data")
        pa_map = dict(zip(df_pa_list['Tên PA'], df_pa_list['Mã PA']))
        
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            sel_pa = st.selectbox("⚡ CHỌN PHƯƠNG ÁN VAY VỐN:", ["-- Trống --"] + list(pa_map.keys()))
        with col_p2:
            st.write("")
            if st.button("ÁP DỤNG", type="primary", use_container_width=True):
                if sel_pa != "-- Trống --":
                    apply_plan_callback(sel_pa, pa_map[sel_pa], df_pa_data, mapping)
                    st.rerun()

    st.divider()
    # HIỂN THỊ CÁC TAB
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
                is_tall = any(k in d_low for k in LONG_TEXT_HINTS)
                kwargs = {"ph": ph, "d_lower": d_low, "field_types": field_types, "mapping": mapping}
                with cols[c_idx]:
                    if is_noicap_desc(d_low) or is_loan_type_desc(d_low):
                        st.selectbox(ds, [""] + (NOICAP_OPTIONS if is_noicap_desc(d_low) else LOAN_TYPE_OPTIONS), key=ph, on_change=process_field_change, kwargs=kwargs)
                    elif is_tall:
                        st.text_area(ds, key=ph, height=120, on_change=process_field_change, kwargs=kwargs)
                    else:
                        st.text_input(ds, key=ph, on_change=process_field_change, kwargs=kwargs)
                c_idx = 1 - c_idx

    # XUẤT FILE
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
            
        out = io.BytesIO()
        from core.document import generate_word_document
        generate_word_document("temp_template.docx", out, ctx); out.seek(0)
        
        name_key = next((p for p, d in mapping.items() if is_name_desc(d.lower())), None)
        member_name = "HoSo"
        if name_key and st.session_state[name_key]:
            import unicodedata
            member_name = unicodedata.normalize('NFKD', st.session_state[name_key]).encode('ASCII', 'ignore').decode('utf-8')
            member_name = "".join(c for c in member_name if c.isalnum() or c in [' ', '-', '_']).replace(" ", "_")

        st.download_button("📥 TẢI FILE WORD", data=out, file_name=f"{member_name}_{datetime.now().strftime('%H%M')}.docx", type="primary")

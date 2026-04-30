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
# CÔNG CỤ TÀI CHÍNH MẠNH MẼ
# ==========================================
def number_to_vn_money(num_str):
    try:
        from num2words import num2words
        txt = num2words(int(num_str), lang="vi").replace("-", " ")
        return txt[0].upper() + txt[1:] + " đồng chẵn."
    except: return ""

def calculate_amounts_live(mapping):
    """Tính toán tự động: Cộng vốn & Nhân tỷ lệ ra tiền"""
    # 1. Tự động tính Tổng vốn đầu tư
    vay_keys = [p for p in mapping if var_name(p).lower() in ["tienvay", "sotienvay", "tongvay", "vay", "sotienxinvay"]]
    tc_keys = [p for p in mapping if var_name(p).lower() in ["vontuco", "von_tu_co", "tuco"]]
    tvdt_keys = [p for p in mapping if var_name(p).lower() in ["tvdt", "tongvon", "tongvondautu", "tongmucdautu", "tongchiphi"]]
    
    val_vay = parse_number_vi(st.session_state[vay_keys[0]]) if vay_keys and st.session_state.get(vay_keys[0]) else 0
    val_tc = parse_number_vi(st.session_state[tc_keys[0]]) if tc_keys and st.session_state.get(tc_keys[0]) else 0
    
    total_val = 0
    if tvdt_keys:
        if val_vay > 0 or val_tc > 0:
            total_val = val_vay + val_tc
            st.session_state[tvdt_keys[0]] = format_money_vi(total_val)
        else:
            total_val = parse_number_vi(st.session_state[tvdt_keys[0]]) if st.session_state.get(tvdt_keys[0]) else 0

    if total_val <= 0: return

    # 2. Tự động tính Số tiền từng Hạng mục Phương án
    for ph in mapping.keys():
        k_var = var_name(ph).lower()
        
        # Số tiền Chi phí (Ví dụ: stcp1, cp1...)
        m_cp = re.match(r"^(?:stcp|st_cp|sotiencp|tien_cp|tiencp|sotien_cp|cp)(\d+)$", k_var)
        if m_cp:
            idx = m_cp.group(1)
            # Dò tìm ô Tỷ lệ tương ứng (tlcp1, tcp1...)
            rate_keys = [p for p in mapping if re.match(fr"^(?:tlcp|tcp|t\.cp|tylecp|ty_le_cp|tl_cp){idx}$", var_name(p).lower())]
            if rate_keys and st.session_state.get(rate_keys[0]):
                try:
                    rate = float(str(st.session_state[rate_keys[0]]).replace(",", "."))
                    st.session_state[ph] = format_money_vi(total_val * (rate / 100.0))
                except: pass

        # Số tiền Thu nhập (Ví dụ: sttn1, tn1...)
        m_tn = re.match(r"^(?:sttn|st_tn|sotientn|tien_tn|tientn|sotien_tn|tn)(\d+)$", k_var)
        if m_tn:
            idx = m_tn.group(1)
            rate_keys = [p for p in mapping if re.match(fr"^(?:tltn|ttn|t\.tn|tyletn|ty_le_tn|tl_tn){idx}$", var_name(p).lower())]
            if rate_keys and st.session_state.get(rate_keys[0]):
                try:
                    rate = float(str(st.session_state[rate_keys[0]]).replace(",", "."))
                    st.session_state[ph] = format_money_vi(total_val * (rate / 100.0))
                except: pass

# ==========================================
# CALLBACK: SỰ KIỆN KHI NHẬP XONG 1 Ô
# ==========================================
def process_field_change(ph, d_lower, field_types, mapping):
    val = str(st.session_state[ph])
    digits = "".join(c for c in val if c.isdigit())
    ft = str(field_types.get(ph, "")).lower()
    
    # Kích hoạt Live Format Tiền & Dịch chữ
    money_kws = ["doanh thu", "thu nhập", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "lãi", "hạn mức"]
    if ("money" in ft or any(kw in d_lower for kw in money_kws)) and digits:
        st.session_state[ph] = f"{int(digits):,}".replace(",", ".")
        if "spell:" in ft:
            target = re.split(r"^spell\s*:", ft, flags=re.IGNORECASE)[1].strip()
            if target in st.session_state: st.session_state[target] = number_to_vn_money(digits)

    # Kích hoạt Live Tính Lãi Suất (Năm <-> Tháng)
    if var_name(ph).lower() in ["laisuatnam", "ls_nam"]:
        ls_thang_key = next((p for p in mapping if var_name(p).lower() in ["laisuatthang", "ls_thang"]), None)
        if ls_thang_key and re.match(r"^\d*\.?\d*$", val.replace(",", ".")):
            st.session_state[ls_thang_key] = str(round(float(val.replace(",", ".")) / 12, 4)).replace(".", ",")

    # Kích hoạt Auto-ID
    if is_name_desc(d_lower) and val:
        st.session_state[ph] = vi_title_name(val)
        initials = "".join([w[0].upper() for w in val.split() if w])
        for p, d in mapping.items():
            if "mã" in d.lower() and "hồ sơ" in d.lower(): st.session_state[p] = f"{initials}-{datetime.now().strftime('%d%m%y')}-001"

    # Định dạng ngày
    if "date" in ft and len(digits) >= 8:
        st.session_state[ph] = f"{digits[:2]}/{digits[2:4]}/{digits[4:8]}"

    # CHỐT CHẶN: Mỗi khi nhập bất kỳ ô nào, hệ thống tự động kiểm tra xem có cần nhân số tiền Phương án không
    calculate_amounts_live(mapping)

# ==========================================
# CALLBACK: ÁP DỤNG PHƯƠNG ÁN (REGEX MỞ RỘNG)
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

        # Rót Chi phí (Hỗ trợ TẤT CẢ các kiểu viết biến Excel của bạn)
        for i, (_, row) in enumerate(df_cp.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
            hm = str(row.get('Hạng mục', ''))
            nd = str(row.get('Nội dung chi tiết', ''))
            if re.match(fr"^(?:tlcp|tcp|t\.cp|tylecp|ty_le_cp|tl_cp){i}$", k_var): st.session_state[ph] = rate
            if re.match(fr"^(?:hmcp|hm_cp|hangmuccp|hang_muc_cp){i}$", k_var): st.session_state[ph] = hm
            if re.match(fr"^(?:ndcp|nd_cp|noidungcp|nd_chiphi|noidung_cp){i}$", k_var): st.session_state[ph] = nd
            if re.match(fr"^(?:chiphi|chi_phi|cp_nd){i}$", k_var): st.session_state[ph] = f"{hm}: {nd}" if hm else nd

        # Rót Thu nhập
        for i, (_, row) in enumerate(df_tn.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
            hm = str(row.get('Hạng mục', ''))
            nd = str(row.get('Nội dung chi tiết', ''))
            if re.match(fr"^(?:tltn|ttn|t\.tn|tyletn|ty_le_tn|tl_tn){i}$", k_var): st.session_state[ph] = rate
            if re.match(fr"^(?:hmtn|hm_tn|hangmuctn|hang_muc_tn){i}$", k_var): st.session_state[ph] = hm
            if re.match(fr"^(?:ndtn|nd_tn|noidungtn|nd_thunhap|noidung_thunhap){i}$", k_var): st.session_state[ph] = nd
            if re.match(fr"^(?:thunhap|thu_nhap|tn_nd){i}$", k_var): st.session_state[ph] = f"{hm}: {nd}" if hm else nd

    calculate_amounts_live(mapping)

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")
st.markdown("<style>div[data-baseweb='input'] > div { border: 1px solid #2980B9 !important; }</style>", unsafe_allow_html=True)
st.title("🏦 HỆ THỐNG KHỞI TẠO HỒ SƠ TÍN DỤNG")

with st.sidebar:
    st.header("⚙️ TẢI FILE CẤU HÌNH")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    pa_file = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    st.divider()
    st.header("🎯 QUÉT CCCD CÁCH LY THEO TAB")
    
    if excel_file:
        with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        
        # GIẢI PHÁP CHỐNG GHI ĐÈ: Chọn Tab để rót dữ liệu
        unique_tabs = list(dict.fromkeys(tabs_dict.values()))
        target_tab = st.selectbox("Rót dữ liệu CCCD vào Tab nào?", ["-- Chọn Tab --"] + unique_tabs)
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        
        if qr_img and st.button("🚀 TIẾN HÀNH QUÉT", type="primary", use_container_width=True):
            if target_tab == "-- Chọn Tab --":
                st.error("❗ Vui lòng chọn Tab bạn muốn điền dữ liệu trước!")
            else:
                with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
                try:
                    info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                    for ph, ds in mapping.items():
                        # CHỈ RÓT DỮ LIỆU VÀO ĐÚNG TAB ĐƯỢC CHỌN, BẢO VỆ 100% CÁC TAB KHÁC
                        if tabs_dict.get(ph) == target_tab:
                            d_low = ds.lower()
                            if is_cccd_desc(d_low): st.session_state[ph] = info.get("CCCD", "")
                            elif is_name_desc(d_low): st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                            elif is_dob_desc(d_low): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                            elif is_addr_desc(d_low): st.session_state[ph] = info.get("Dia_chi", "")
                            elif is_issue_desc(d_low): st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                            elif is_gender_desc(d_low): st.session_state[ph] = info.get("Gioi_tinh", "")
                    st.success(f"Đã nạp CCCD an toàn vào Tab: {target_tab}")
                    st.rerun()
                except Exception as e: st.error(f"Lỗi: {e}")

if excel_file and docx_file:
    for ph in mapping.keys():
        if ph not in st.session_state: st.session_state[ph] = ""

    if pa_file:
        df_pa_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_pa_data = pd.read_excel(pa_file, sheet_name="data")
        pa_map = dict(zip(df_pa_list['Tên PA'], df_pa_list['Mã PA']))
        
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            sel_pa = st.selectbox("⚡ CHỌN PHƯƠNG ÁN:", ["-- Trống --"] + list(pa_map.keys()))
        with col_p2:
            st.write("")
            if st.button("ÁP DỤNG PHƯƠNG ÁN", type="primary", use_container_width=True):
                if sel_pa != "-- Trống --":
                    apply_plan_callback(sel_pa, pa_map[sel_pa], df_pa_data, mapping)
                    st.rerun()

    st.divider()
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

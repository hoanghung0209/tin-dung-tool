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
# CÔNG CỤ TRỢ LÝ TÀI CHÍNH & NGÔN NGỮ
# ==========================================
def number_to_vn_money(num_str):
    try:
        from num2words import num2words
        txt = num2words(int(num_str), lang="vi").replace("-", " ")
        return txt[0].upper() + txt[1:] + " đồng chẵn."
    except: return ""

def number_to_vn_decimal(text_val):
    try:
        from num2words import num2words
        if not text_val: return ""
        parts = text_val.split(",")
        int_txt = num2words(int(parts[0]) if parts[0] else 0, lang="vi").replace("-", " ")
        if len(parts) < 2: return int_txt[0].upper() + int_txt[1:] + " mét vuông."
        
        dec_part = parts[1]
        leading_zeros = "".join(["không " for char in dec_part if char == '0' and not dec_part[:dec_part.find(char)].replace('0','') ])
        dec_num = int(dec_part)
        dec_txt = leading_zeros + (num2words(dec_num, lang="vi").replace("-", " ") if dec_num > 0 else "")
        final_txt = f"{int_txt} phẩy {dec_txt}".strip()
        return final_txt[0].upper() + final_txt[1:] + " mét vuông."
    except: return ""

def calculate_amounts_live(mapping):
    """Tự động tính lại toàn bộ số tiền dựa trên tỷ lệ và tổng vốn"""
    tvdt_key = next((ph for ph, ds in mapping.items() if var_name(ph).lower() in ["tvdt", "tongvon", "tongvondautu"]), None)
    if not tvdt_key or not st.session_state.get(tvdt_key): return
    
    total_val = parse_number_vi(st.session_state[tvdt_key])
    if total_val <= 0: return

    for ph, ds in mapping.items():
        k_var = var_name(ph).lower()
        # Tính chi phí
        m_cp = re.match(r"^(?:cp|st_cp)(\d+)$", k_var)
        if m_cp:
            idx = m_cp.group(1)
            rate_key = next((p for p, d in mapping.items() if var_name(p).lower() in [f"tlcp{idx}", f"tcp{idx}", f"t.cp{idx}"]), None)
            if rate_key and st.session_state.get(rate_key):
                rate = float(str(st.session_state[rate_key]).replace(",", "."))
                st.session_state[ph] = format_money_vi(total_val * (rate / 100.0))

        # Tính thu nhập
        m_tn = re.match(r"^(?:tn|st_tn)(\d+)$", k_var)
        if m_tn:
            idx = m_tn.group(1)
            rate_key = next((p for p, d in mapping.items() if var_name(p).lower() in [f"tltn{idx}", f"ttn{idx}", f"t.tn{idx}"]), None)
            if rate_key and st.session_state.get(rate_key):
                rate = float(str(st.session_state[rate_key]).replace(",", "."))
                st.session_state[ph] = format_money_vi(total_val * (rate / 100.0))

# ==========================================
# LOGIC CHỐNG GHI ĐÈ QR & AUTO-FILL
# ==========================================
def is_target_match_strict(target, d_low):
    if target == "Tất cả": return True
    if target.lower() not in d_low: return False
    
    if target == "Thành viên":
        for ex in ["đồng vay", "ủy quyền", "vợ", "chồng", "thụ hưởng", "bảo lãnh"]:
            if ex in d_low: return False
    
    if target == "Người ủy quyền 1" and "người ủy quyền 2" in d_low: return False
    return True

def process_field_change(ph, d_lower, field_types, mapping):
    val = str(st.session_state[ph])
    digits = "".join(c for c in val if c.isdigit())
    ft = str(field_types.get(ph, "")).lower()
    
    money_kws = ["doanh thu", "thu nhập", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "lãi", "hạn mức"]
    if ("money" in ft or any(kw in d_lower for kw in money_kws)) and digits:
        st.session_state[ph] = f"{int(digits):,}".replace(",", ".")
        if "spell:" in ft:
            target = re.split(r"^spell\s*:", ft, flags=re.IGNORECASE)[1].strip()
            if target in st.session_state: st.session_state[target] = number_to_vn_money(digits)
        
        # Nếu nhập Vốn Vay hoặc Tự Có -> Tính Tổng vốn -> Kéo theo toàn bộ file phương án
        if var_name(ph).lower() in ["tienvay", "sotienvay", "vontuco", "von_tu_co", "tvdt", "tongvon"]:
            # Tự cộng Tổng Vốn
            vay_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["tienvay", "sotienvay"]), None)
            tu_co_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["vontuco", "von_tu_co"]), None)
            tvdt_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["tvdt", "tongvon", "tongvondautu"]), None)
            if vay_key and tu_co_key and tvdt_key:
                val_vay = parse_number_vi(st.session_state.get(vay_key, "0"))
                val_tc = parse_number_vi(st.session_state.get(tu_co_key, "0"))
                if val_vay > 0 or val_tc > 0: st.session_state[tvdt_key] = format_money_vi(val_vay + val_tc)
            
            # Kích hoạt tính lại tiền chi tiết
            calculate_amounts_live(mapping)
            
    # Auto tính lãi tháng
    elif var_name(ph).lower() in ["laisuatnam", "ls_nam"]:
        ls_thang_key = next((p for p, d in mapping.items() if var_name(p).lower() in ["laisuatthang", "ls_thang"]), None)
        if ls_thang_key and re.match(r"^\d*\.?\d*$", val.replace(",", ".")):
            st.session_state[ls_thang_key] = str(round(float(val.replace(",", ".")) / 12, 4)).replace(".", ",")

    elif ("area" in ft or "spell_area:" in ft) and val:
        st.session_state[ph] = val.replace(".", ",")
        if "spell_area:" in ft:
            target = re.split(r"^spell_area\s*:", ft, flags=re.IGNORECASE)[1].strip()
            if target in st.session_state: st.session_state[target] = number_to_vn_decimal(val.replace(".", ","))

    elif is_name_desc(d_lower) and val:
        st.session_state[ph] = vi_title_name(val)

# ==========================================
# CẤU HÌNH GIAO DIỆN
# ==========================================
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")
st.markdown("<style>div[data-baseweb='input'] > div { border: 1px solid #2980B9 !important; }</style>", unsafe_allow_html=True)

st.title("🏦 HỆ THỐNG KHỞI TẠO HỒ SƠ TÍN DỤNG")
st.caption("PHIÊN BẢN CHỐNG GHI ĐÈ & TỰ ĐỘNG FULL TÍNH TOÁN - QTDND PHÙNG HƯNG")

with st.sidebar:
    st.header("⚙️ CẤU HÌNH")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    pa_file = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    st.divider()
    st.header("📸 QUÉT CCCD")
    if excel_file:
        with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        
        target_person = st.selectbox("Quét cho ai?", ["Thành viên", "Đồng vay vốn", "Người ủy quyền 1", "Người ủy quyền 2", "Tất cả"])
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
                        elif is_gender_desc(d_low): st.session_state[ph] = info.get("Gioi_tinh", "")
                st.success(f"Đã nạp xong: {target_person}"); st.rerun()
            except Exception as e: st.error(f"Lỗi: {e}")

if excel_file and docx_file:
    for ph in mapping.keys():
        if ph not in st.session_state: st.session_state[ph] = ""

    # KHÔI PHỤC CHỨC NĂNG PHƯƠNG ÁN FULL
    if pa_file:
        df_pa_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_pa_data = pd.read_excel(pa_file, sheet_name="data")
        pa_map = dict(zip(df_pa_list['Tên PA'], df_pa_list['Mã PA']))
        
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            sel_pa = st.selectbox("⚡ CHỌN PHƯƠNG ÁN VAY VỐN:", ["-- Trống --"] + list(pa_map.keys()))
        with col_p2:
            st.write("")
            if st.button("ÁP DỤNG PHƯƠNG ÁN", type="primary", use_container_width=True):
                if sel_pa != "-- Trống --":
                    ma_pa = pa_map[sel_pa]
                    plan_rows = df_pa_data[df_pa_data['Mã phương án'].astype(str).str.strip().str.lower() == str(ma_pa).lower()]
                    df_cp = plan_rows[plan_rows['Loại'].astype(str).str.contains('chi|phí|cp', case=False, na=False)]
                    df_tn = plan_rows[plan_rows['Loại'].astype(str).str.contains('lợi|nhuận|thu|nhập|ln|tn', case=False, na=False)]
                    
                    for ph, ds in mapping.items():
                        k_var = var_name(ph).lower()
                        if "tên phương án" in ds.lower(): st.session_state[ph] = sel_pa
                        
                        # Rót toàn bộ Hạng mục & Nội dung chi tiết (Chi phí)
                        for i, (_, row) in enumerate(df_cp.iterrows(), 1):
                            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
                            hm = str(row.get('Hạng mục', ''))
                            nd = str(row.get('Nội dung chi tiết', ''))
                            if re.match(fr"^(?:t\.cp|tcp|tlcp){i}$", k_var) and rate: st.session_state[ph] = rate
                            if re.match(fr"^(?:hm_cp|hmcp){i}$", k_var) and hm: st.session_state[ph] = hm
                            if re.match(fr"^(?:nd_cp|nd_chiphi|ndcp){i}$", k_var) and nd: st.session_state[ph] = nd
                            if re.match(fr"^(?:chiphi|noidung_cp){i}$", k_var): st.session_state[ph] = f"{hm}: {nd}" if hm else nd
                            
                        # Rót toàn bộ Hạng mục & Nội dung chi tiết (Thu nhập)
                        for i, (_, row) in enumerate(df_tn.iterrows(), 1):
                            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
                            hm = str(row.get('Hạng mục', ''))
                            nd = str(row.get('Nội dung chi tiết', ''))
                            if re.match(fr"^(?:t\.tn|ttn|tltn){i}$", k_var) and rate: st.session_state[ph] = rate
                            if re.match(fr"^(?:hm_tn|hmtn){i}$", k_var) and hm: st.session_state[ph] = hm
                            if re.match(fr"^(?:nd_tn|nd_thunhap|ndtn){i}$", k_var) and nd: st.session_state[ph] = nd
                    
                    # Cập nhật số tiền nếu Tổng vốn đã được nhập trước đó
                    calculate_amounts_live(mapping)
                    st.rerun()

    st.divider()
    # Vẽ Form
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

    # Xuất File
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

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
    is_name_desc, is_dob_desc, is_issue_desc, is_addr_desc, is_gender_desc
)
from core.scanner import decode_qr_offline
from core.document import read_mapping

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")
st.markdown("""
    <style>
    div[data-baseweb="input"] > div { border: 1px solid #2980B9; border-radius: 4px; }
    div[data-baseweb="textarea"] > div { border: 1px solid #2980B9; border-radius: 4px; }
    div[data-baseweb="select"] > div { border: 1px solid #2980B9; border-radius: 4px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. HÀM XỬ LÝ TIỀN TỆ, LÃI SUẤT & DIỆN TÍCH
# ==========================================
def parse_money(val):
    d = "".join(c for c in str(val) if c.isdigit())
    return int(d) if d else 0

def format_money(val):
    return f"{int(val):,}".replace(",", ".")

def parse_rate(val):
    c = "".join(x for x in str(val) if x.isdigit() or x in ",.")
    if not c: return 0.0
    c = c.replace(",", ".")
    try: return float(c)
    except: return 0.0

def format_rate(val):
    s = f"{val:.4f}".rstrip("0").rstrip(".")
    if s.endswith(","): s = s[:-1]
    if s == "": s = "0"
    return s.replace(".", ",")

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

# ==========================================
# 2. HÀM TÍNH TOÁN CỐT LÕI (VỐN, LÃI, PHƯƠNG ÁN)
# ==========================================
def calculate_globals(mapping, field_types):
    """Tính toán mọi thông số tài chính và tự động điền"""
    # Lấy các key cấu hình
    vay_keys = [p for p in mapping if var_name(p).lower().replace("_", "") in ["tienvay", "sotienvay", "sovonvay", "sotienxinvay", "mucvay", "sovonxinvay", "sovay"]]
    tc_keys = [p for p in mapping if var_name(p).lower().replace("_", "") in ["vontuco", "tuco", "vonthamgia", "sovontuco"]]
    tvdt_keys = [p for p in mapping if var_name(p).lower().replace("_", "") in ["tvdt", "tongvon", "tongvondautu", "tongmucdautu", "tongchiphi", "tongsovon", "tongsovondautu"]]
    
    lsnam_keys = [p for p in mapping if "năm" in str(mapping[p]).lower() and ("ls" in var_name(p).lower() or "lãi" in str(mapping[p]).lower())]
    lsthang_keys = [p for p in mapping if "tháng" in str(mapping[p]).lower() and ("ls" in var_name(p).lower() or "lãi" in str(mapping[p]).lower())]

    # Tính Tổng vốn
    vay_val = max([parse_money(st.session_state.get(k, "0")) for k in vay_keys]) if vay_keys else 0
    tuco_val = max([parse_money(st.session_state.get(k, "0")) for k in tc_keys]) if tc_keys else 0
    total_von = vay_val + tuco_val
    
    if tvdt_keys:
        tvdt_ph = tvdt_keys[0]
        if total_von > 0:
            st.session_state[tvdt_ph] = format_money(total_von)
        else:
            total_von = parse_money(st.session_state.get(tvdt_ph, "0"))

    # Tính Lãi tháng
    if lsnam_keys and lsthang_keys:
        lsnam_val = parse_rate(st.session_state.get(lsnam_keys[0], "0"))
        if lsnam_val > 0:
            st.session_state[lsthang_keys[0]] = format_rate(lsnam_val / 12.0)

    # Rót tiền theo Tỷ lệ Phương án
    if total_von > 0:
        for ph, ds in mapping.items():
            kv = var_name(ph).lower().replace("_", "")
            
            # Chi phí
            m_cp = re.match(r"^(?:stcp|sotiencp|cp|chiphi|tien|sotien)(\d+)$", kv)
            if m_cp and "tỉ lệ" not in ds.lower() and "tỷ lệ" not in ds.lower():
                idx = m_cp.group(1)
                rate_ph = next((p for p in mapping if re.match(fr"^(?:tlcp|tcp|tylecp|t){idx}$", var_name(p).lower().replace("_", ""))), None)
                if rate_ph:
                    rate = parse_rate(st.session_state.get(rate_ph, "0"))
                    if rate > 0: st.session_state[ph] = format_money(total_von * rate / 100)

            # Thu nhập
            m_tn = re.match(r"^(?:sttn|sotientn|tn|thunhap|dt|stdt|doanhthu|tien|sotien)(\d+)$", kv)
            if m_tn and "tỉ lệ" not in ds.lower() and "tỷ lệ" not in ds.lower() and not m_cp:
                idx = m_tn.group(1)
                rate_ph = next((p for p in mapping if re.match(fr"^(?:tltn|ttn|tyletn|tldt|tyledt){idx}$", var_name(p).lower().replace("_", ""))), None)
                if rate_ph:
                    rate = parse_rate(st.session_state.get(rate_ph, "0"))
                    if rate > 0: st.session_state[ph] = format_money(total_von * rate / 100)

    # Đọc số thành chữ (Spell)
    for ph, ds in mapping.items():
        ft = str(field_types.get(ph, "")).strip().lower()
        if "spell:" in ft:
            target_ph = ft.split("spell:")[1].strip()
            if target_ph in st.session_state:
                money_val = parse_money(st.session_state.get(ph, ""))
                if money_val > 0:
                    try:
                        from num2words import num2words
                        txt = num2words(money_val, lang="vi").replace("-", " ")
                        st.session_state[target_ph] = txt[0].upper() + txt[1:] + " đồng chẵn."
                    except: pass

# ==========================================
# 3. SỰ KIỆN GÕ PHÍM (ON_CHANGE)
# ==========================================
def handle_change(ph, ds, field_types, mapping):
    val = str(st.session_state.get(ph, ""))
    d_low = ds.lower()
    ft = str(field_types.get(ph, "")).lower()
    
    # Format Lãi suất (Bảo tồn dấu phẩy)
    if "lãi suất" in d_low or "lãi" in d_low or "tỷ lệ" in d_low or "ls" in var_name(ph).lower():
        clean_val = "".join(c for c in val if c.isdigit() or c in ",.")
        st.session_state[ph] = clean_val
    
    # Format Tiền tệ
    elif "money" in ft or any(k in d_low for k in ["số tiền", "vốn", "thu nhập", "chi phí", "giá trị", "doanh thu", "hạn mức"]):
        digits = "".join(c for c in val if c.isdigit())
        if digits: st.session_state[ph] = f"{int(digits):,}".replace(",", ".")
            
    # Format Diện tích đất ra chữ
    elif "area" in ft or "spell_area:" in ft:
        clean_val = "".join(c for c in val if c.isdigit() or c in ",.")
        st.session_state[ph] = clean_val
        if "spell_area:" in ft:
            target_ph = ft.split("spell_area:")[1].strip()
            if target_ph in st.session_state:
                st.session_state[target_ph] = number_to_vn_decimal(clean_val)
                
    # Format Tên và Auto-ID
    elif is_name_desc(d_low) and val:
        st.session_state[ph] = vi_title_name(val)
        initials = "".join([w[0].upper() for w in val.split() if w])
        for p, d in mapping.items():
            if "mã" in d.lower() and "hồ sơ" in d.lower():
                st.session_state[p] = f"{initials}-{datetime.now().strftime('%d%m%y')}-001"
                
    # Kích hoạt toàn bộ tính toán
    calculate_globals(mapping, field_types)

# ==========================================
# 4. CHỐNG GHI ĐÈ CCCD TẦNG CAO
# ==========================================
def get_group(ds):
    d = ds.lower()
    if "thụ hưởng" in d: return "Người thụ hưởng"
    if "ủy quyền 2" in d: return "Người ủy quyền 2"
    if "ủy quyền 1" in d or "ủy quyền" in d: return "Người ủy quyền 1"
    if "đồng vay 2" in d: return "Người đồng vay vốn 2"
    if "đồng vay" in d or "cùng vay" in d: return "Người đồng vay vốn 1"
    if any(x in d for x in ["vợ", "chồng", "bảo lãnh"]): return "Khác"
    return "Thành viên"

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
with st.sidebar:
    st.header("⚙️ CẤU HÌNH ĐẦU VÀO")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    pa_file = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    if excel_file:
        with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        
        # Khởi tạo khóa Session State
        for ph in mapping:
            if ph not in st.session_state: st.session_state[ph] = ""
            
        st.divider()
        st.header("📸 QUÉT CCCD ĐÍCH DANH")
        # Thêm Người thụ hưởng vào danh sách
        subjects = ["Thành viên", "Người đồng vay vốn 1", "Người ủy quyền 1", "Người ủy quyền 2", "Người thụ hưởng"]
        target = st.selectbox("Chọn đối tượng:", subjects)
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        
        if qr_img and st.button("🚀 TIẾN HÀNH QUÉT", type="primary", use_container_width=True):
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                for ph, ds in mapping.items():
                    # CHỐNG GHI ĐÈ BẰNG TƯỜNG LỬA GET_GROUP
                    if get_group(ds) == target:
                        if is_cccd_desc(ds.lower()): st.session_state[ph] = info.get("CCCD", "")
                        elif is_name_desc(ds.lower()): st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif is_dob_desc(ds.lower()): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif is_addr_desc(ds.lower()): st.session_state[ph] = info.get("Dia_chi", "")
                        elif is_issue_desc(ds.lower()): st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                        elif is_gender_desc(ds.lower()): st.session_state[ph] = info.get("Gioi_tinh", "")
                st.success(f"Nạp xong cho: {target}")
                st.rerun()
            except Exception as e: st.error(f"Lỗi: {e}")

if excel_file and docx_file:
    # 1. FILL PHƯƠNG ÁN (GỒM CẢ THU NHẬP)
    if pa_file:
        df_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_data = pd.read_excel(pa_file, sheet_name="data")
        pa_map = dict(zip(df_list['Tên PA'], df_list['Mã PA']))
        
        c1, c2 = st.columns([3, 1])
        with c1: sel_pa = st.selectbox("⚡ CHỌN PHƯƠNG ÁN:", ["-- Trống --"] + list(pa_map.keys()))
        with c2:
            st.write("")
            if st.button("ÁP DỤNG", type="primary", use_container_width=True):
                if sel_pa != "-- Trống --":
                    ma = pa_map[sel_pa]
                    rows = df_data[df_data['Mã phương án'].astype(str).str.strip().str.lower() == str(ma).lower()]
                    df_cp = rows[rows['Loại'].astype(str).str.contains('chi|phí|cp', case=False, na=False)].reset_index()
                    df_tn = rows[rows['Loại'].astype(str).str.contains('thu|nhập|tn|dt|doanh', case=False, na=False)].reset_index()
                    
                    for ph, ds in mapping.items():
                        kv = var_name(ph).lower().replace("_", "")
                        if "tenphuongan" in kv or "tenpa" in kv: st.session_state[ph] = sel_pa
                        
                        # Nội dung Chi phí
                        for i, r in df_cp.iterrows():
                            idx = str(i + 1)
                            rate = format_rate(float(r['Tỉ lệ']) * 100) if pd.notna(r.get('Tỉ lệ')) else ""
                            hm = str(r.get('Hạng mục', '')).strip() if pd.notna(r.get('Hạng mục')) else ""
                            nd = str(r.get('Nội dung chi tiết', '')).strip() if pd.notna(r.get('Nội dung chi tiết')) else ""
                            
                            if re.match(fr"^(?:tlcp|tcp|tylecp){idx}$", kv): st.session_state[ph] = rate
                            if re.match(fr"^(?:hmcp|hangmuccp){idx}$", kv): st.session_state[ph] = hm
                            if re.match(fr"^(?:ndcp|noidungcp|ndchiphi|noidungchiphi|nd){idx}$", kv): st.session_state[ph] = nd
                            if re.match(fr"^(?:chiphi){idx}$", kv): st.session_state[ph] = nd if not hm else f"{hm}: {nd}"

                        # Nội dung Thu nhập
                        for i, r in df_tn.iterrows():
                            idx = str(i + 1)
                            rate = format_rate(float(r['Tỉ lệ']) * 100) if pd.notna(r.get('Tỉ lệ')) else ""
                            hm = str(r.get('Hạng mục', '')).strip() if pd.notna(r.get('Hạng mục')) else ""
                            nd = str(r.get('Nội dung chi tiết', '')).strip() if pd.notna(r.get('Nội dung chi tiết')) else ""
                            
                            if re.match(fr"^(?:tltn|ttn|tyletn|tldt|tyledt){idx}$", kv): st.session_state[ph] = rate
                            if re.match(fr"^(?:hmtn|hangmuctn|hmdt|hangmucdt){idx}$", kv): st.session_state[ph] = hm
                            if re.match(fr"^(?:ndtn|noidungtn|nddt|noidungdt|ndthunhap|noidungthunhap){idx}$", kv): st.session_state[ph] = nd
                            if re.match(fr"^(?:thunhap|doanhthu){idx}$", kv): st.session_state[ph] = nd if not hm else f"{hm}: {nd}"
                    
                    calculate_globals(mapping, field_types)
                    st.rerun()

    # 2. HIỂN THỊ CÁC TAB 
    st.divider()
    unique_tabs = list(dict.fromkeys(tabs_dict.values()))
    st_tabs = st.tabs(unique_tabs)
    
    for i, t_name in enumerate(unique_tabs):
        with st_tabs[i]:
            fields = [(ph, ds) for ph, ds in mapping.items() if tabs_dict.get(ph) == t_name]
            c1, c2 = st.columns(2)
            for j, (ph, ds) in enumerate(fields):
                with c1 if j % 2 == 0 else c2:
                    ft = str(field_types.get(ph, "")).strip()
                    is_tall = any(k in ds.lower() for k in LONG_TEXT_HINTS)
                    kw = {"ph": ph, "ds": ds, "field_types": field_types, "mapping": mapping}
                    
                    # Giải quyết Dropdown (Loại cho vay, Thời hạn vay...)
                    if ft.lower().startswith("dropdown"):
                        raw_opts = ft.split(":", 1)[1]
                        # Chia lựa chọn bằng dấu phẩy hoặc dấu gạch chéo
                        if "," in raw_opts: opts = [o.strip() for o in raw_opts.split(",")]
                        elif "/" in raw_opts: opts = [o.strip() for o in raw_opts.split("/")]
                        else: opts = [raw_opts.strip()]
                        
                        st.selectbox(ds, [""] + opts, key=ph, on_change=handle_change, kwargs=kw)
                    elif is_noicap_desc(ds.lower()):
                        st.selectbox(ds, [""] + NOICAP_OPTIONS, key=ph, on_change=handle_change, kwargs=kw)
                    elif is_loan_type_desc(ds.lower()):
                        st.selectbox(ds, [""] + LOAN_TYPE_OPTIONS, key=ph, on_change=handle_change, kwargs=kw)
                    elif is_tall:
                        st.text_area(ds, key=ph, height=120, on_change=handle_change, kwargs=kw)
                    else:
                        st.text_input(ds, key=ph, on_change=handle_change, kwargs=kw)

    # 3. XUẤT FILE WORD
    st.divider()
    if st.button("🚀 XUẤT HỒ SƠ WORD", type="primary", use_container_width=True):
        calculate_globals(mapping, field_types)
        ctx = {}
        for ph, ds in mapping.items():
            val = str(st.session_state.get(ph, "")).strip()
            if "\n" in val:
                rt = RichText()
                for line in val.split("\n"): rt.add(line); rt.add_line_break()
                ctx[var_name(ph)] = rt
            else: ctx[var_name(ph)] = val
            
        out = io.BytesIO()
        from core.document import generate_word_document
        generate_word_document("temp_template.docx", out, ctx); out.seek(0)
        
        name_key = next((p for p, d in mapping.items() if is_name_desc(d.lower())), None)
        member_name = "HoSo"
        if name_key and st.session_state.get(name_key):
            import unicodedata
            member_name = unicodedata.normalize('NFKD', st.session_state[name_key]).encode('ASCII', 'ignore').decode('utf-8')
            member_name = "".join(c for c in member_name if c.isalnum() or c in [' ', '-', '_']).replace(" ", "_")

        st.download_button("📥 TẢI FILE WORD", data=out, file_name=f"{member_name}_{datetime.now().strftime('%H%M')}.docx", type="primary")

import os
import io
import re
import pandas as pd
from datetime import datetime
import streamlit as st
from docxtpl import DocxTemplate, RichText

# --- MODULE CẤU HÌNH VÀ TIỆN ÍCH ---
from config import NOICAP_OPTIONS, LOAN_TYPE_OPTIONS, LONG_TEXT_HINTS
from core.utils import (
    vi_title_name, parse_cccd_payload, validate_cccd_12_digits,
    var_name, is_cccd_desc, is_noicap_desc, is_loan_type_desc, 
    is_name_desc, is_dob_desc, is_issue_desc, is_addr_desc, is_gender_desc
)
from core.scanner import decode_qr_offline
from core.document import read_mapping

st.set_page_config(page_title="Hồ Sơ Tín Dụng Pro", page_icon="🏦", layout="wide")
st.markdown("""
    <style>
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, div[data-baseweb="select"] > div { 
        border: 1px solid #2980B9; border-radius: 4px; 
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# MODULE 1: XỬ LÝ DỮ LIỆU & TOÁN HỌC
# ==========================================
def extract_number(val):
    """Trích xuất số nguyên từ chuỗi (loại bỏ mọi ký tự thừa)"""
    digits = "".join(c for c in str(val) if c.isdigit())
    return int(digits) if digits else 0

def extract_float(val):
    """Trích xuất số thập phân (dành cho lãi suất, tỷ lệ)"""
    c = "".join(x for x in str(val) if x.isdigit() or x in ",.")
    return float(c.replace(",", ".")) if c else 0.0

def format_currency(val):
    return f"{int(val):,}".replace(",", ".")

def format_percent(val):
    s = f"{val:.4f}".rstrip("0").rstrip(".")
    return "0" if not s else s.replace(".", ",")

def num_to_text_area(text_val):
    """Dịch diện tích (có dấu phẩy) ra chữ"""
    try:
        from num2words import num2words
        if not text_val: return ""
        parts = str(text_val).split(",")
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
# MODULE 2: ĐỘNG CƠ TÍNH TOÁN TÀI CHÍNH (CORE ENGINE)
# ==========================================
def execute_financial_engine(mapping, field_types):
    """Trái tim của hệ thống: Tính Tổng vốn, Lãi suất, và rót tiền Phương án"""
    tvdt_key = ls_thang_key = None
    vay_val = tuco_val = ls_nam_val = 0.0

    # Bước 1: Quét tìm dữ liệu mồi
    for ph, ds in mapping.items():
        kv = var_name(ph).lower().replace("_", "")
        val = str(st.session_state.get(ph, "")).strip()

        if kv in ["tienvay", "sotienvay", "sovonvay", "sotienxinvay", "mucvay", "sovay"]: vay_val = extract_number(val)
        elif kv in ["vontuco", "tuco", "vonthamgia", "sovontuco"]: tuco_val = extract_number(val)
        elif kv in ["tvdt", "tongvon", "tongvondautu", "tongmucdautu", "tongchiphi"]: tvdt_key = ph
        elif ("ls" in kv or "laisuat" in kv) and "năm" in ds.lower(): ls_nam_val = extract_float(val)
        elif ("ls" in kv or "laisuat" in kv) and "tháng" in ds.lower(): ls_thang_key = ph

    # Bước 2: Xử lý Vốn & Lãi suất
    total_capital = vay_val + tuco_val
    if total_capital > 0 and tvdt_key:
        st.session_state[tvdt_key] = format_currency(total_capital)
    elif tvdt_key:
        total_capital = extract_number(st.session_state.get(tvdt_key, "0"))

    if ls_nam_val > 0 and ls_thang_key:
        st.session_state[ls_thang_key] = format_percent(ls_nam_val / 12.0)

    # Bước 3: Rót tiền cho Phương Án
    if total_capital > 0:
        for ph, ds in mapping.items():
            kv = var_name(ph).lower().replace("_", "")
            
            # Tính Chi phí
            m_cp = re.match(r"^(?:stcp|sotiencp|cp|chiphi|tien|sotien)(\d+)$", kv)
            if m_cp and "lệ" not in ds.lower():
                idx = m_cp.group(1)
                rate_ph = next((p for p in mapping if re.match(fr"^(?:tlcp|tcp|tylecp|t){idx}$", var_name(p).lower().replace("_", ""))), None)
                if rate_ph and extract_float(st.session_state.get(rate_ph, "0")) > 0:
                    st.session_state[ph] = format_currency(total_capital * extract_float(st.session_state[rate_ph]) / 100)

            # Tính Thu nhập
            m_tn = re.match(r"^(?:sttn|sotientn|tn|thunhap|dt|stdt|doanhthu|tien|sotien)(\d+)$", kv)
            if m_tn and "lệ" not in ds.lower() and not m_cp:
                idx = m_tn.group(1)
                rate_ph = next((p for p in mapping if re.match(fr"^(?:tltn|ttn|tyletn|tldt|tyledt){idx}$", var_name(p).lower().replace("_", ""))), None)
                if rate_ph and extract_float(st.session_state.get(rate_ph, "0")) > 0:
                    st.session_state[ph] = format_currency(total_capital * extract_float(st.session_state[rate_ph]) / 100)

    # Bước 4: Dịch tiền ra chữ
    for ph, ds in mapping.items():
        ft = str(field_types.get(ph, "")).strip().lower()
        if "spell:" in ft:
            target = ft.split("spell:")[1].strip()
            money_val = extract_number(st.session_state.get(ph, "0"))
            if target in st.session_state and money_val > 0:
                try:
                    from num2words import num2words
                    txt = num2words(money_val, lang="vi").replace("-", " ")
                    st.session_state[target] = txt[0].upper() + txt[1:] + " đồng chẵn."
                except: pass

# ==========================================
# MODULE 3: CALLBAKS & BỘ LỌC CÁCH LY
# ==========================================
def field_callback(ph, ds, field_types, mapping):
    """Lắng nghe mọi thao tác nhập liệu để làm sạch và định dạng"""
    val = str(st.session_state.get(ph, ""))
    d_low = ds.lower()
    ft = str(field_types.get(ph, "")).lower()
    
    if "lãi suất" in d_low or "lãi" in d_low or "tỷ lệ" in d_low or "ls" in var_name(ph).lower():
        st.session_state[ph] = "".join(c for c in val if c.isdigit() or c in ",.")
    elif "money" in ft or any(k in d_low for k in ["số tiền", "vốn", "thu nhập", "chi phí", "giá trị", "doanh thu"]):
        digits = extract_number(val)
        if digits > 0: st.session_state[ph] = format_currency(digits)
    elif "area" in ft or "spell_area:" in ft:
        clean_val = "".join(c for c in val if c.isdigit() or c in ",.")
        st.session_state[ph] = clean_val
        if "spell_area:" in ft:
            target = ft.split("spell_area:")[1].strip()
            if target in st.session_state: st.session_state[target] = num_to_text_area(clean_val)
    elif is_name_desc(d_low) and val:
        st.session_state[ph] = vi_title_name(val)
        
    # Kích hoạt động cơ tính toán
    execute_financial_engine(mapping, field_types)

def firewall_qr_target(ph, ds):
    """Bức tường lửa: Cách ly hoàn toàn các vùng dữ liệu CCCD"""
    text = f"{ds.lower()} | {var_name(ph).lower()}"
    if any(x in text for x in ["thụ hưởng", "thuhuong", "_th"]): return "Người thụ hưởng"
    if any(x in text for x in ["ủy quyền 2", "uyquyen2", "uq2"]): return "Người ủy quyền 2"
    if any(x in text for x in ["ủy quyền 1", "ủy quyền", "uyquyen1", "uq1"]): return "Người ủy quyền 1"
    if any(x in text for x in ["đồng vay 2", "dongvay2", "dv2"]): return "Người đồng vay vốn 2"
    if any(x in text for x in ["đồng vay", "cùng vay", "dongvay", "dv1", "_dv"]): return "Người đồng vay vốn 1"
    if any(x in text for x in ["vợ", "vo", "chồng", "chong", "bảo lãnh"]): return "Khác"
    return "Thành viên"

# ==========================================
# MODULE 4: GIAO DIỆN & TƯƠNG TÁC WEB
# ==========================================
with st.sidebar:
    st.header("⚙️ CẤU HÌNH HỆ THỐNG")
    file_data = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    file_word = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    file_plan = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    if file_data:
        with open("temp_data.xlsx", "wb") as f: f.write(file_data.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        st.session_state.field_types = field_types
        
        # Khởi tạo khóa cứng Session State
        for ph in mapping:
            if ph not in st.session_state: st.session_state[ph] = ""
            
        st.divider()
        st.header("📸 QUÉT CCCD BẢO MẬT")
        subjects = ["Thành viên", "Người đồng vay vốn 1", "Người ủy quyền 1", "Người ủy quyền 2", "Người thụ hưởng"]
        target = st.selectbox("Chọn đích danh:", subjects)
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        
        if qr_img and st.button("🚀 XỬ LÝ QUÉT", type="primary", use_container_width=True):
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                for ph, ds in mapping.items():
                    if firewall_qr_target(ph, ds) == target:
                        d_low = ds.lower()
                        if is_cccd_desc(d_low): st.session_state[ph] = info.get("CCCD", "")
                        elif is_name_desc(d_low): st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif is_dob_desc(d_low): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif is_addr_desc(d_low): st.session_state[ph] = info.get("Dia_chi", "")
                        elif is_issue_desc(d_low): st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                        elif is_gender_desc(d_low): st.session_state[ph] = info.get("Gioi_tinh", "")
                st.success(f"Nạp an toàn cho: {target}")
                st.rerun()
            except Exception as e: st.error(f"Lỗi đọc thẻ: {e}")

if file_data and file_word:
    # --- ÁP DỤNG PHƯƠNG ÁN ---
    if file_plan:
        df_list = pd.read_excel(file_plan, sheet_name="phuongan")
        df_data = pd.read_excel(file_plan, sheet_name="data")
        pa_map = dict(zip(df_list['Tên PA'], df_list['Mã PA']))
        
        c1, c2 = st.columns([3, 1])
        with c1: sel_pa = st.selectbox("⚡ ĐIỀN TỰ ĐỘNG PHƯƠNG ÁN:", ["-- Trống --"] + list(pa_map.keys()))
        with c2:
            st.write("")
            if st.button("ÁP DỤNG", type="primary", use_container_width=True):
                if sel_pa != "-- Trống --":
                    ma = pa_map[sel_pa]
                    rows = df_data[df_data['Mã phương án'].astype(str).str.strip().str.lower() == str(ma).lower()]
                    df_cp = rows[rows['Loại'].astype(str).str.contains('chi|phí|cp', case=False, na=False)].reset_index()
                    df_tn = rows[rows['Loại'].astype(str).str.contains('thu|doanh|lợi|tn|dt', case=False, na=False)].reset_index()
                    
                    for ph, ds in mapping.items():
                        kv = var_name(ph).lower().replace("_", "")
                        if "tenphuongan" in kv or "tenpa" in kv: st.session_state[ph] = sel_pa
                        
                        # Rót Chi phí
                        for i, r in df_cp.iterrows():
                            idx = str(i + 1)
                            rate = format_percent(float(r['Tỉ lệ']) * 100) if pd.notna(r.get('Tỉ lệ')) else ""
                            hm = str(r.get('Hạng mục', '')).strip() if pd.notna(r.get('Hạng mục')) else ""
                            nd = str(r.get('Nội dung chi tiết', '')).strip() if pd.notna(r.get('Nội dung chi tiết')) else ""
                            
                            if re.match(fr"^(?:tlcp|tcp|tylecp){idx}$", kv): st.session_state[ph] = rate
                            if re.match(fr"^(?:hmcp|hangmuccp){idx}$", kv): st.session_state[ph] = hm
                            if re.match(fr"^(?:ndcp|noidungcp|ndchiphi|nd){idx}$", kv): st.session_state[ph] = nd
                            if re.match(fr"^(?:chiphi){idx}$", kv): st.session_state[ph] = nd if not hm else f"{hm}: {nd}"

                        # Rót Thu nhập
                        for i, r in df_tn.iterrows():
                            idx = str(i + 1)
                            rate = format_percent(float(r['Tỉ lệ']) * 100) if pd.notna(r.get('Tỉ lệ')) else ""
                            hm = str(r.get('Hạng mục', '')).strip() if pd.notna(r.get('Hạng mục')) else ""
                            nd = str(r.get('Nội dung chi tiết', '')).strip() if pd.notna(r.get('Nội dung chi tiết')) else ""
                            
                            if re.match(fr"^(?:tltn|ttn|tyletn|tldt){idx}$", kv): st.session_state[ph] = rate
                            if re.match(fr"^(?:hmtn|hangmuctn|hmdt){idx}$", kv): st.session_state[ph] = hm
                            if re.match(fr"^(?:ndtn|noidungtn|nddt|noidungdt|ndthunhap){idx}$", kv): st.session_state[ph] = nd
                            if re.match(fr"^(?:thunhap|doanhthu){idx}$", kv): st.session_state[ph] = nd if not hm else f"{hm}: {nd}"
                    
                    execute_financial_engine(mapping, field_types)
                    st.rerun()

    # --- RENDER FORM ---
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
                    
                    if ft.lower().startswith("dropdown"):
                        raw = ft.split(":", 1)[1]
                        opts = [o.strip() for o in raw.split(",")] if "," in raw else [o.strip() for o in raw.split("/")] if "/" in raw else [raw.strip()]
                        st.selectbox(ds, [""] + opts, key=ph, on_change=field_callback, kwargs=kw)
                    elif is_noicap_desc(ds.lower()):
                        st.selectbox(ds, [""] + NOICAP_OPTIONS, key=ph, on_change=field_callback, kwargs=kw)
                    elif is_loan_type_desc(ds.lower()):
                        st.selectbox(ds, [""] + LOAN_TYPE_OPTIONS, key=ph, on_change=field_callback, kwargs=kw)
                    elif is_tall:
                        st.text_area(ds, key=ph, height=120, on_change=field_callback, kwargs=kw)
                    else:
                        st.text_input(ds, key=ph, on_change=field_callback, kwargs=kw)

    # --- XUẤT HỒ SƠ WORD ---
    st.divider()
    if st.button("🚀 XUẤT HỒ SƠ WORD TỔNG HỢP", type="primary", use_container_width=True):
        execute_financial_engine(mapping, field_types)
        ctx = {}
        for ph, ds in mapping.items():
            val = str(st.session_state.get(ph, "")).strip()
            # Xử lý xuống dòng mượt mà cho Word (RichText)
            if "\n" in val:
                rt = RichText()
                lines = val.split("\n")
                for i, line in enumerate(lines):
                    rt.add(line)
                    if i < len(lines) - 1: rt.add_line_break()
                ctx[var_name(ph)] = rt
            else:
                ctx[var_name(ph)] = val
            
        out = io.BytesIO()
        from core.document import generate_word_document
        generate_word_document("temp_template.docx", out, ctx)
        out.seek(0)
        
        name_key = next((p for p, d in mapping.items() if is_name_desc(d.lower())), None)
        member_name = "HoSo"
        if name_key and st.session_state.get(name_key):
            import unicodedata
            member_name = unicodedata.normalize('NFKD', st.session_state[name_key]).encode('ASCII', 'ignore').decode('utf-8')
            member_name = "".join(c for c in member_name if c.isalnum() or c in [' ', '-', '_']).replace(" ", "_")

        st.download_button("📥 TẢI FILE WORD (.DOCX)", data=out, file_name=f"{member_name}_{datetime.now().strftime('%H%M')}.docx", type="primary")

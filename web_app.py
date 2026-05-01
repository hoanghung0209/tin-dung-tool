import os
import io
import re
import pandas as pd
from datetime import datetime
import streamlit as st
from docxtpl import DocxTemplate, RichText

# --- MODULE CẤU HÌNH ---
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
    div[data-baseweb="input"] > div { border: 1px solid #2980B9; border-radius: 4px; }
    div[data-baseweb="textarea"] > div { border: 1px solid #2980B9; border-radius: 4px; }
    div[data-baseweb="select"] > div { border: 1px solid #2980B9; border-radius: 4px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. HÀM XỬ LÝ TIỀN TỆ & TOÁN HỌC
# ==========================================
def extract_number(val):
    if not val: return 0
    digits = "".join(c for c in str(val) if c.isdigit())
    return int(digits) if digits else 0

def extract_float(val):
    if not val: return 0.0
    c = str(val).replace(",", ".")
    c = "".join(x for x in c if x.isdigit() or x == ".")
    try: return float(c)
    except: return 0.0

def format_currency(val):
    return f"{int(val):,}".replace(",", ".")

def format_percent(val):
    s = f"{val:.4f}".rstrip("0").rstrip(".")
    return "0" if not s else s.replace(".", ",")

# ==========================================
# 2. ĐỘNG CƠ TÍNH TOÁN TÀI CHÍNH
# ==========================================
def execute_financial_engine(mapping, field_types):
    tvdt_keys = []
    ls_thang_keys = []
    vay_val = 0
    tuco_val = 0
    ls_nam_val = 0.0

    # 1. Tìm các biến số quan trọng
    for ph, ds in mapping.items():
        kv = var_name(ph).lower().replace("_", "")
        d_low = ds.lower()
        val = str(st.session_state.get(ph, "")).strip()

        if any(x in kv for x in ["vay", "xinvay", "mucvay"]) and not any(x in kv for x in ["tong", "ls", "lai", "thoi", "ngay", "loai"]): 
            vay_val = max(vay_val, extract_number(val))
        elif any(x in kv for x in ["tuco", "vonthamgia", "vontc"]): 
            tuco_val = max(tuco_val, extract_number(val))
        elif any(x in kv for x in ["tongvon", "tvdt", "tongmucdautu", "tongchiphi", "tongsovon"]): 
            tvdt_keys.append(ph)
        elif ("ls" in kv or "laisuat" in kv) and "năm" in d_low: 
            ls_nam_val = max(ls_nam_val, extract_float(val))
        elif ("ls" in kv or "laisuat" in kv) and "tháng" in d_low: 
            ls_thang_keys.append(ph)

    # 2. Tính Tổng Vốn & Lãi tháng
    total_capital = vay_val + tuco_val
    
    if total_capital > 0:
        for k in tvdt_keys:
            st.session_state[k] = format_currency(total_capital)
    else:
        for k in tvdt_keys:
            total_capital = max(total_capital, extract_number(st.session_state.get(k, "0")))

    if ls_nam_val > 0:
        for k in ls_thang_keys:
            st.session_state[k] = format_percent(ls_nam_val / 12.0)

    # 3. Rót tiền Phương án
    if total_capital > 0:
        for ph, ds in mapping.items():
            kv = var_name(ph).lower().replace("_", "")
            m = re.search(r'(\d+)$', kv)
            if not m or "lệ" in ds.lower(): continue
            
            idx = m.group(1)
            base = kv[:-len(idx)]
            
            if base in ["stcp", "sotiencp", "cp", "chiphi", "tien", "sotien", "sttn", "sotientn", "tn", "thunhap", "dt", "stdt", "doanhthu"]:
                prefix = "tlcp" if any(x in base for x in ["cp", "chiphi"]) else "tltn" if any(x in base for x in ["tn", "thunhap", "dt", "doanhthu"]) else None
                if not prefix: 
                    prefix = "tlcp" if "chi phí" in ds.lower() else "tltn" if "thu nhập" in ds.lower() else None
                
                if prefix:
                    rate_ph = next((p for p in mapping if var_name(p).lower().replace("_", "") in [f"{prefix}{idx}", f"t{prefix[2:]}{idx}"]), None)
                    if rate_ph:
                        rate = extract_float(st.session_state.get(rate_ph, "0"))
                        if rate > 0: st.session_state[ph] = format_currency(total_capital * rate / 100)

    # 4. Spell (Dịch số ra chữ)
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
# 3. SỰ KIỆN GÕ PHÍM & BỘ LỌC CCCD
# ==========================================
def field_callback(ph, ds, field_types, mapping):
    val = str(st.session_state.get(ph, ""))
    d_low = ds.lower()
    
    if "lãi suất" in d_low or "lãi" in d_low or "tỷ lệ" in d_low or "ls" in var_name(ph).lower():
        st.session_state[ph] = "".join(c for c in val if c.isdigit() or c in ",.")
    elif any(k in d_low for k in ["số tiền", "vốn", "thu nhập", "chi phí", "giá trị", "doanh thu"]):
        digits = "".join(c for c in val if c.isdigit())
        if digits: st.session_state[ph] = format_currency(digits)
    elif "họ tên" in d_low or "họ và tên" in d_low:
        st.session_state[ph] = vi_title_name(val)
        
    execute_financial_engine(mapping, field_types)

def get_cccd_field_type(ds):
    """KHÓA TỬ HUYỆT: Chặn đứng các trường đất đai, sổ ruộng khoán"""
    d = ds.lower()
    if any(x in d for x in ["sổ", "đất", "gcn", "giấy chứng nhận", "ruộng", "xe", "tsđb", "tài sản"]): return None
    
    if any(x in d for x in ["ngày cấp", "cấp ngày", "ngay cap"]): return "issue_date"
    if any(x in d for x in ["cccd", "cmnd", "căn cước", "chứng minh", "định danh"]): return "cccd"
    if any(x in d for x in ["họ tên", "họ và tên", "người"]): return "name"
    if any(x in d for x in ["ngày sinh", "sinh ngày", "năm sinh"]): return "dob"
    if any(x in d for x in ["địa chỉ", "thường trú", "cư trú", "hộ khẩu", "nơi ở"]): return "address"
    if any(x in d for x in ["giới tính", "nam/nữ"]): return "gender"
    return None

def firewall_qr_target(ph, ds):
    """BẢO VỆ PHÂN VÙNG DỮ LIỆU"""
    text = f"{ds.lower()} | {var_name(ph).lower()}"
    if any(x in text for x in ["thụ hưởng", "thuhuong", "_th", "nhận tiền"]): return "Người thụ hưởng"
    if any(x in text for x in ["ủy quyền 2", "uyquyen2", "uq2"]): return "Người ủy quyền 2"
    if any(x in text for x in ["ủy quyền 1", "ủy quyền", "uyquyen1", "uq1"]): return "Người ủy quyền 1"
    if any(x in text for x in ["đồng vay 2", "dongvay2", "dv2"]): return "Người đồng vay vốn 2"
    if any(x in text for x in ["đồng vay", "cùng vay", "dongvay", "dv1", "_dv", "dv_"]): return "Người đồng vay vốn 1"
    if any(x in text for x in ["vợ", "vo", "chồng", "chong", "bảo lãnh"]): return "Khác"
    return "Thành viên"

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
with st.sidebar:
    st.header("⚙️ CẤU HÌNH HỆ THỐNG")
    file_data = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    file_word = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    file_plan = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])
    
    if file_data:
        with open("temp_data.xlsx", "wb") as f: f.write(file_data.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        
        for ph in mapping:
            if ph not in st.session_state: st.session_state[ph] = ""
            
        st.divider()
        st.header("📸 QUÉT CCCD ĐÍCH DANH")
        subjects = ["Thành viên", "Người đồng vay vốn 1", "Người ủy quyền 1", "Người ủy quyền 2", "Người thụ hưởng"]
        target = st.selectbox("Chọn đối tượng:", subjects)
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        
        if qr_img and st.button("🚀 TIẾN HÀNH QUÉT", type="primary", use_container_width=True):
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                for ph, ds in mapping.items():
                    if firewall_qr_target(ph, ds) == target:
                        cccd_field = get_cccd_field_type(ds)
                        if cccd_field == "cccd": st.session_state[ph] = info.get("CCCD", "")
                        elif cccd_field == "name": st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif cccd_field == "dob": st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif cccd_field == "address": st.session_state[ph] = info.get("Dia_chi", "")
                        elif cccd_field == "issue_date": st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                        elif cccd_field == "gender": st.session_state[ph] = info.get("Gioi_tinh", "")
                
                execute_financial_engine(mapping, field_types)
                st.success(f"Nạp xong cho: {target}")
                st.rerun()
            except Exception as e: st.error(f"Lỗi: {e}")

if file_data and file_word:
    if file_plan:
        df_list = pd.read_excel(file_plan, sheet_name="phuongan")
        df_data = pd.read_excel(file_plan, sheet_name="data")
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
                    df_tn = rows[rows['Loại'].astype(str).str.contains('thu|doanh|lợi|tn|dt', case=False, na=False)].reset_index()
                    
                    for ph, ds in mapping.items():
                        kv = var_name(ph).lower().replace("_", "")
                        if "tenphuongan" in kv or "tenpa" in kv: st.session_state[ph] = sel_pa
                        
                        for i, r in df_cp.iterrows():
                            idx = str(i + 1)
                            rate = format_percent(float(r['Tỉ lệ']) * 100) if pd.notna(r.get('Tỉ lệ')) else ""
                            hm = str(r.get('Hạng mục', '')).strip() if pd.notna(r.get('Hạng mục')) else ""
                            nd = str(r.get('Nội dung chi tiết', '')).strip() if pd.notna(r.get('Nội dung chi tiết')) else ""
                            
                            if re.match(fr"^(?:tlcp|tcp|tylecp){idx}$", kv): st.session_state[ph] = rate
                            if re.match(fr"^(?:hmcp|hangmuccp){idx}$", kv): st.session_state[ph] = hm
                            if re.match(fr"^(?:ndcp|noidungcp|ndchiphi|nd){idx}$", kv): st.session_state[ph] = nd
                            if re.match(fr"^(?:chiphi){idx}$", kv): st.session_state[ph] = nd if not hm else f"{hm}: {nd}"

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

    st.divider()
    if st.button("🚀 XUẤT HỒ SƠ WORD", type="primary", use_container_width=True):
        execute_financial_engine(mapping, field_types)
        ctx = {}
        for ph, ds in mapping.items():
            val = str(st.session_state.get(ph, "")).strip()
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

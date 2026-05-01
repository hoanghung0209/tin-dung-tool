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
# 1. HÀM XỬ LÝ TIỀN TỆ & TÍNH TOÁN
# ==========================================
def parse_money(val):
    """Bóc tách mọi dấu chấm, phẩy để lấy số nguyên chuẩn"""
    if not val: return 0
    s = str(val).replace(".", "").replace(",", "").replace(" ", "")
    digits = "".join(c for c in s if c.isdigit())
    return int(digits) if digits else 0

def format_money(val):
    return f"{int(val):,}".replace(",", ".")

def parse_rate(val):
    c = "".join(x for x in str(val) if x.isdigit() or x in ",.")
    if not c: return 0.0
    return float(c.replace(",", "."))

def format_rate(val):
    s = f"{val:.4f}".rstrip("0").rstrip(".")
    return "0" if s == "" else s.replace(".", ",")

def calculate_globals(mapping, field_types):
    """Tính toán Lãi suất, Tổng vốn, và Số tiền phương án"""
    vay_val = tuco_val = tvdt_val = 0
    tvdt_ph = lsnam_ph = lsthang_ph = None
    lsnam_val = 0.0

    # 1. Quét tìm các ô quan trọng
    for ph, ds in mapping.items():
        kv = var_name(ph).lower().replace("_", "")
        val = str(st.session_state.get(ph, "")).strip()

        if kv in ["tienvay", "sotienvay", "sovonvay", "sotienxinvay", "mucvay", "sovay"]:
            vay_val = parse_money(val)
        elif kv in ["vontuco", "tuco", "vonthamgia", "sovontuco"]:
            tuco_val = parse_money(val)
        elif kv in ["tvdt", "tongvon", "tongvondautu", "tongmucdautu", "tongchiphi", "tongsovondautu"]:
            tvdt_ph = ph
            tvdt_val = parse_money(val)
        elif ("ls" in kv or "laisuat" in kv) and "năm" in ds.lower():
            lsnam_ph = ph
            lsnam_val = parse_rate(val)
        elif ("ls" in kv or "laisuat" in kv) and "tháng" in ds.lower():
            lsthang_ph = ph

    # 2. Tính Tổng vốn = Vay + Tự có
    total_von = vay_val + tuco_val
    if total_von > 0 and tvdt_ph:
        st.session_state[tvdt_ph] = format_money(total_von)
    elif tvdt_val > 0:
        total_von = tvdt_val

    # 3. Tính Lãi tháng
    if lsnam_val > 0 and lsthang_ph:
        st.session_state[lsthang_ph] = format_rate(lsnam_val / 12.0)

    # 4. Rót tiền cho Phương án (Dựa trên Tổng Vốn)
    if total_von > 0:
        for ph, ds in mapping.items():
            kv = var_name(ph).lower().replace("_", "")
            
            # Check Tiền Chi Phí
            m_cp = re.match(r"^(?:stcp|sotiencp|cp|chiphi|tien|sotien)(\d+)$", kv)
            if m_cp and "lệ" not in ds.lower():
                idx = m_cp.group(1)
                rate_ph = next((p for p in mapping if var_name(p).lower().replace("_", "") in [f"tlcp{idx}", f"tcp{idx}", f"tylecp{idx}", f"t{idx}"]), None)
                if rate_ph:
                    r_val = parse_rate(st.session_state.get(rate_ph, "0"))
                    if r_val > 0: st.session_state[ph] = format_money(total_von * r_val / 100)

            # Check Tiền Thu Nhập
            m_tn = re.match(r"^(?:sttn|sotientn|tn|thunhap|dt|stdt|doanhthu|tien|sotien)(\d+)$", kv)
            if m_tn and "lệ" not in ds.lower() and not m_cp:
                idx = m_tn.group(1)
                rate_ph = next((p for p in mapping if var_name(p).lower().replace("_", "") in [f"tltn{idx}", f"ttn{idx}", f"tyletn{idx}", f"tldt{idx}", f"tyledt{idx}"]), None)
                if rate_ph:
                    r_val = parse_rate(st.session_state.get(rate_ph, "0"))
                    if r_val > 0: st.session_state[ph] = format_money(total_von * r_val / 100)

    # 5. Dịch số ra chữ
    for ph, ds in mapping.items():
        ft = str(field_types.get(ph, "")).strip().lower()
        if "spell:" in ft:
            target_ph = ft.split("spell:")[1].strip()
            if target_ph in st.session_state:
                money_val = parse_money(st.session_state.get(ph, "0"))
                if money_val > 0:
                    try:
                        from num2words import num2words
                        txt = num2words(money_val, lang="vi").replace("-", " ")
                        st.session_state[target_ph] = txt[0].upper() + txt[1:] + " đồng chẵn."
                    except: pass

# ==========================================
# 2. XỬ LÝ SỰ KIỆN GÕ PHÍM
# ==========================================
def handle_change(ph, ds, field_types, mapping):
    val = str(st.session_state.get(ph, ""))
    d_low = ds.lower()
    
    if "lãi suất" in d_low or "lãi" in d_low or "tỷ lệ" in d_low or "ls" in var_name(ph).lower():
        st.session_state[ph] = "".join(c for c in val if c.isdigit() or c in ",.")
    elif any(k in d_low for k in ["số tiền", "vốn", "thu nhập", "chi phí", "giá trị", "doanh thu"]):
        digits = "".join(c for c in val if c.isdigit())
        if digits: st.session_state[ph] = format_money(digits)
    elif is_name_desc(d_low) and val:
        st.session_state[ph] = vi_title_name(val)
        
    calculate_globals(mapping, field_types)

# ==========================================
# 3. BỘ LỌC CÁCH LY CCCD (CHỐNG GHI ĐÈ 100%)
# ==========================================
def get_group(ph, ds):
    """
    KIỂM TRA CHÉO CẢ MÃ BIẾN LẪN MÔ TẢ ĐỂ CHỐT CHẶN VÙNG DỮ LIỆU
    """
    text = f"{ds.lower()} | {var_name(ph).lower()}"
    if any(x in text for x in ["thụ hưởng", "thuhuong", "_th"]): return "Người thụ hưởng"
    if any(x in text for x in ["ủy quyền 2", "uyquyen2", "uq2"]): return "Người ủy quyền 2"
    if any(x in text for x in ["ủy quyền 1", "ủy quyền", "uyquyen1", "uq1"]): return "Người ủy quyền 1"
    if any(x in text for x in ["đồng vay 2", "dongvay2", "dv2"]): return "Người đồng vay vốn 2"
    if any(x in text for x in ["đồng vay", "cùng vay", "dongvay", "dv1", "_dv"]): return "Người đồng vay vốn 1"
    if any(x in text for x in ["vợ", "vo", "chồng", "chong", "bảo lãnh"]): return "Khác"
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
        st.session_state.field_types = field_types
        
        # Khóa cứng Session
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
                    if get_group(ph, ds) == target:
                        if is_cccd_desc(ds.lower()): st.session_state[ph] = info.get("CCCD", "")
                        elif is_name_desc(ds.lower()): st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif is_dob_desc(ds.lower()): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif is_addr_desc(ds.lower()): st.session_state[ph] = info.get("Dia_chi", "")
                        elif is_issue_desc(ds.lower()): st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                        elif is_gender_desc(ds.lower()): st.session_state[ph] = info.get("Gioi_tinh", "")
                st.success(f"Đã nạp chính xác cho: {target}")
                st.rerun()
            except Exception as e: st.error(f"Lỗi: {e}")

if excel_file and docx_file:
    # 1. RÓT PHƯƠNG ÁN (FIX LỖI THIẾU THU NHẬP)
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
                        
                        # Rót Chi Phí
                        for i, r in df_cp.iterrows():
                            idx = str(i + 1)
                            if kv.endswith(idx):
                                base = kv[:-len(idx)]
                                rate = format_rate(float(r['Tỉ lệ']) * 100) if pd.notna(r.get('Tỉ lệ')) else ""
                                hm = str(r.get('Hạng mục', '')).strip() if pd.notna(r.get('Hạng mục')) else ""
                                nd = str(r.get('Nội dung chi tiết', '')).strip() if pd.notna(r.get('Nội dung chi tiết')) else ""
                                
                                if base in ["tlcp", "tcp", "tylecp"]: st.session_state[ph] = rate
                                elif base in ["hmcp", "hangmuccp"]: st.session_state[ph] = hm
                                elif base in ["ndcp", "noidungcp", "ndchiphi", "noidungchiphi", "nd"]: st.session_state[ph] = nd
                                elif base in ["chiphi", "cp"]: st.session_state[ph] = nd if not hm else f"{hm}: {nd}"

                        # Rót Thu Nhập
                        for i, r in df_tn.iterrows():
                            idx = str(i + 1)
                            if kv.endswith(idx):
                                base = kv[:-len(idx)]
                                rate = format_rate(float(r['Tỉ lệ']) * 100) if pd.notna(r.get('Tỉ lệ')) else ""
                                hm = str(r.get('Hạng mục', '')).strip() if pd.notna(r.get('Hạng mục')) else ""
                                nd = str(r.get('Nội dung chi tiết', '')).strip() if pd.notna(r.get('Nội dung chi tiết')) else ""
                                
                                if base in ["tltn", "ttn", "tyletn", "tldt", "tyledt"]: st.session_state[ph] = rate
                                elif base in ["hmtn", "hangmuctn", "hmdt", "hangmucdt"]: st.session_state[ph] = hm
                                elif base in ["ndtn", "noidungtn", "nddt", "noidungdt", "ndthunhap", "noidungthunhap"]: st.session_state[ph] = nd
                                elif base in ["thunhap", "doanhthu", "tn", "dt"]: st.session_state[ph] = nd if not hm else f"{hm}: {nd}"
                    
                    calculate_globals(mapping, field_types)
                    st.rerun()

    # 2. HIỂN THỊ TAB VÀ WIDGETS
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
                        opts = [o.strip() for o in ft.split(":", 1)[1].split(",")] if "," in ft else [o.strip() for o in ft.split(":", 1)[1].split("/")] if "/" in ft else [ft.split(":", 1)[1].strip()]
                        st.selectbox(ds, [""] + opts, key=ph, on_change=handle_change, kwargs=kw)
                    elif is_noicap_desc(ds.lower()):
                        st.selectbox(ds, [""] + NOICAP_OPTIONS, key=ph, on_change=handle_change, kwargs=kw)
                    elif is_loan_type_desc(ds.lower()):
                        st.selectbox(ds, [""] + LOAN_TYPE_OPTIONS, key=ph, on_change=handle_change, kwargs=kw)
                    elif is_tall:
                        st.text_area(ds, key=ph, height=120, on_change=handle_change, kwargs=kw)
                    else:
                        st.text_input(ds, key=ph, on_change=handle_change, kwargs=kw)

    # 3. XUẤT FILE
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

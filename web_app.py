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
    var_name
)
from core.scanner import decode_qr_offline
from core.document import read_mapping

st.set_page_config(page_title="Hồ Sơ Tín Dụng Pro", page_icon="🏦", layout="wide")
st.markdown("""
    <style>
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, div[data-baseweb="select"] > div { 
        border: 1px solid #2980B9; border-radius: 4px; 
    }
    .stExpander { border: 2px solid #e0e0e0; border-radius: 8px; margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. HÀM XỬ LÝ TIỀN TỆ & TOÁN HỌC CHUẨN
# ==========================================
def extract_number(val):
    if not val: return 0
    digits = "".join(c for c in str(val) if c.isdigit())
    return int(digits) if digits else 0

def extract_float(val):
    if not val: return 0.0
    c = "".join(x for x in str(val) if x.isdigit() or x in ",.")
    return float(c.replace(",", ".")) if c else 0.0

def format_currency(val):
    return f"{int(val):,}".replace(",", ".")

def format_percent(val):
    s = f"{val:.4f}".rstrip("0").rstrip(".")
    return "0" if not s else s.replace(".", ",")

def num_to_text_area(text_val):
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
# 2. ĐỘNG CƠ TÍNH TOÁN TÀI CHÍNH (SỬA LỖI TỔNG VỐN)
# ==========================================
def calculate_globals(mapping, field_types):
    tvdt_keys = []
    ls_thang_keys = []
    vay_val = 0
    tuco_val = 0
    ls_nam_val = 0.0

    # BƯỚC 1: Quét tìm các biến độc lập
    for ph, ds in mapping.items():
        kv = var_name(ph).lower().replace("_", "")
        d_low = ds.lower()
        val = str(st.session_state.get(ph, "")).strip()

        # Tiền vay
        if any(x in kv for x in ["vay", "xinvay", "mucvay"]) and not any(x in kv for x in ["tong", "ls", "lai", "thoi", "ngay", "loai"]): 
            vay_val = max(vay_val, extract_number(val))
        # Vốn tự có
        elif any(x in kv for x in ["tuco", "vonthamgia", "vontc"]): 
            tuco_val = max(tuco_val, extract_number(val))
        # Tổng vốn
        elif any(x in kv for x in ["tongvon", "tvdt", "tongmucdautu", "tongchiphi", "tongsovon"]): 
            tvdt_keys.append(ph)
        # Lãi suất
        elif ("ls" in kv or "laisuat" in kv) and "năm" in d_low: 
            ls_nam_val = max(ls_nam_val, extract_float(val))
        elif ("ls" in kv or "laisuat" in kv) and "tháng" in d_low: 
            ls_thang_keys.append(ph)

    # BƯỚC 2: Tính toán Lãi và Tổng vốn chuẩn xác
    total_capital = vay_val + tuco_val
    
    # Ghi đè vào các ô Tổng vốn
    if total_capital > 0:
        for k in tvdt_keys:
            st.session_state[k] = format_currency(total_capital)
    else:
        # Nếu chưa gõ Vay/Tự có nhưng đã gõ Tổng vốn
        for k in tvdt_keys:
            total_capital = max(total_capital, extract_number(st.session_state.get(k, "0")))

    # Tính lãi tháng
    if ls_nam_val > 0:
        for k in ls_thang_keys:
            st.session_state[k] = format_percent(ls_nam_val / 12.0)

    # BƯỚC 3: Rót tiền cho Phương án (Tự động nhân tỷ lệ)
    if total_capital > 0:
        for ph, ds in mapping.items():
            kv = var_name(ph).lower().replace("_", "")
            m = re.search(r'(\d+)$', kv)
            if not m or "lệ" in ds.lower(): continue
            
            idx = m.group(1)
            base = kv[:-len(idx)]
            
            if base in ["stcp", "sotiencp", "cp", "chiphi", "tien", "sotien", "sttn", "sotientn", "tn", "thunhap", "dt", "stdt", "doanhthu"]:
                # Tìm ô Tỷ lệ tương ứng
                prefix = "tlcp" if any(x in base for x in ["cp", "chiphi"]) else "tltn" if any(x in base for x in ["tn", "thunhap", "dt", "doanhthu"]) else None
                if not prefix: 
                    prefix = "tlcp" if "chi phí" in ds.lower() else "tltn" if "thu nhập" in ds.lower() else None
                
                if prefix:
                    rate_ph = next((p for p in mapping if var_name(p).lower().replace("_", "") in [f"{prefix}{idx}", f"t{prefix[2:]}{idx}"]), None)
                    if rate_ph:
                        rate = extract_float(st.session_state.get(rate_ph, "0"))
                        if rate > 0: st.session_state[ph] = format_currency(total_capital * rate / 100)

    # BƯỚC 4: Dịch tiền ra chữ
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
# 3. ĐIỀU HƯỚNG TƯƠNG TÁC (CALLBACKS)
# ==========================================
def field_callback(ph, ds, field_types, mapping):
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
    elif ("họ tên" in d_low or "họ và tên" in d_low) and val:
        st.session_state[ph] = vi_title_name(val)
        
    execute_financial_engine(mapping, field_types)

# ==========================================
# 4. CHỐT CHẶN BẢO MẬT CCCD (CỰC KỲ NGHIÊM NGẶT)
# ==========================================
def get_cccd_field_type(ds):
    """KHÓA TỬ HUYỆT: Từ chối mọi trường ngày cấp có chữ Sổ, Đất, Ruộng, GCN..."""
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
    """PHÂN VÙNG DỮ LIỆU: Cô lập tuyệt đối 5 đối tượng"""
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
        st.session_state.field_types = field_types
        
        # Khóa State
        for ph in mapping:
            if ph not in st.session_state: st.session_state[ph] = ""
            
        st.divider()
        st.header("📸 BẢNG QUÉT CCCD SONG SONG")
        subjects = ["Thành viên", "Người đồng vay vốn 1", "Người ủy quyền 1", "Người ủy quyền 2", "Người thụ hưởng"]
        
        for target in subjects:
            with st.expander(f"👤 {target}", expanded=False):
                qr_img = st.file_uploader(f"Ảnh CCCD", type=["png", "jpg", "jpeg"], key=f"qr_{target}")
                if qr_img and st.button(f"🚀 Quét dữ liệu", key=f"btn_{target}", use_container_width=True):
                    with open(f"temp_qr_{target}.jpg", "wb") as f: f.write(qr_img.getbuffer())
                    try:
                        info = parse_cccd_payload(decode_qr_offline(f"temp_qr_{target}.jpg"))
                        count = 0
                        for ph, ds in mapping.items():
                            # Kiểm tra chéo 2 lớp: Phân vùng (firewall) và Loại trường (get_cccd_field)
                            if firewall_qr_target(ph, ds) == target:
                                cccd_field = get_cccd_field_type(ds)
                                if cccd_field == "cccd": st.session_state[ph] = info.get("CCCD", "")
                                elif cccd_field == "name": st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                                elif cccd_field == "dob": st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                                elif cccd_field == "address": st.session_state[ph] = info.get("Dia_chi", "")
                                elif cccd_field == "issue_date": st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                                elif cccd_field == "gender": st.session_state[ph] = info.get("Gioi_tinh", "")
                                if cccd_field: count += 1
                        
                        execute_financial_engine(mapping, field_types)
                        st.success(f"Nạp an toàn {count} trường cho: {target}")
                        st.rerun()
                    except Exception as e: 
                        st.error(f"Lỗi đọc thẻ: {e}")

if file_data and file_word:
    # --- ÁP DỤNG PHƯƠNG ÁN (AUTO-FILL BẤT BẠI) ---
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
                        if "tenphuongan" in kv or "tenpa" in kv: 
                            st.session_state[ph] = sel_pa
                            continue
                        
                        m = re.search(r'(\d+)$', kv)
                        if not m: continue
                        
                        idx = int(m.group(1)) - 1
                        base = kv[:-len(m.group(1))]
                        
                        # Phân loại là Chi phí hay Thu nhập dựa trên Mã biến hoặc Mô tả
                        is_cp = any(x in base for x in ["cp", "chiphi"]) or "chi phí" in ds.lower()
                        is_tn = any(x in base for x in ["tn", "thunhap", "dt", "doanhthu"]) or "thu nhập" in ds.lower() or "doanh thu" in ds.lower()
                        
                        target_df = df_cp if is_cp else df_tn if is_tn else None
                        
                        if target_df is not None and 0 <= idx < len(target_df):
                            row = target_df.iloc[idx]
                            rate = format_percent(float(row['Tỉ lệ']) * 100) if pd.notna(row.get('Tỉ lệ')) else ""
                            hm = str(row.get('Hạng mục', '')).strip() if pd.notna(row.get('Hạng mục')) else ""
                            nd = str(row.get('Nội dung chi tiết', '')).strip() if pd.notna(row.get('Nội dung chi tiết')) else ""
                            
                            if any(x in base for x in ["tl", "tyle", "tcp", "ttn"]): st.session_state[ph] = rate
                            elif any(x in base for x in ["hm", "hangmuc"]): st.session_state[ph] = hm
                            elif any(x in base for x in ["nd", "noidung", "chitiet"]): st.session_state[ph] = nd
                            elif base in ["chiphi", "cp", "thunhap", "tn", "doanhthu", "dt"]: st.session_state[ph] = nd if not hm else f"{hm}: {nd}"
                    
                    execute_financial_engine(mapping, field_types)
                    st.rerun()

    # --- RENDER FORM TABS ---
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

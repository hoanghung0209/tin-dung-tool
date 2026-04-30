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

# ==========================================
# CÁC HÀM BỔ TRỢ & LOGIC NGHIỆP VỤ
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

def is_target_match(target, d_low):
    """Hàm chống ghi đè: Lọc chính xác đối tượng cần điền QR"""
    if target == "Tất cả": return True
    if target.lower() not in d_low: return False
    
    # Nếu là Thành viên chính, tuyệt đối không được ghi đè vào các trường của người khác
    if target == "Thành viên":
        exclusions = ["đồng vay", "ủy quyền", "vợ", "chồng", "bảo lãnh", "thụ hưởng"]
        if any(ex in d_low for ex in exclusions):
            return False
    return True

# ==========================================
# CALLBACKS: KÍCH HOẠT KHI NHẬP LIỆU XONG
# ==========================================
def process_field_change(ph, d_lower, field_types, mapping):
    val = str(st.session_state[ph])
    digits = "".join(c for c in val if c.isdigit())
    digits_and_comma = "".join(c for c in val if c.isdigit() or c == ",")

    ft = str(field_types.get(ph, "")).lower()
    money_kws = ["doanh thu", "thu nhập", "chi phí", "số tiền", "giá trị", "vốn", "định giá", "lãi", "hạn mức"]
    is_money = "money" in ft or "spell:" in ft or any(kw in d_lower for kw in money_kws)

    # Format Tiền
    if is_money and digits:
        st.session_state[ph] = f"{int(digits):,}".replace(",", ".")
        if "spell:" in ft:
            target = re.split(r"^spell\s*:", ft, flags=re.IGNORECASE)[1].strip()
            if target in st.session_state: st.session_state[target] = number_to_vn_money(digits)
    
    # Format Diện tích
    elif ("area" in ft or "spell_area:" in ft) and digits_and_comma:
        pts = digits_and_comma.split(",")
        fmt = f"{int(pts[0]):,}".replace(",", ".") if pts[0] else ""
        if len(pts) > 1: fmt += f",{pts[1]}"
        elif digits_and_comma.endswith(","): fmt += ","
        st.session_state[ph] = fmt

        if "spell_area:" in ft:
            target = re.split(r"^spell_area\s*:", ft, flags=re.IGNORECASE)[1].strip()
            if target in st.session_state: st.session_state[target] = number_to_vn_decimal(digits_and_comma)

    # Viết hoa Tên & Tạo Auto ID
    elif is_name_desc(d_lower) and val:
        st.session_state[ph] = vi_title_name(val)
        initials = "".join([w[0].upper() for w in val.split() if w])
        auto_id = f"{initials}-{datetime.now().strftime('%d%m%y')}-001"
        for p, d in mapping.items():
            if "mã" in d.lower() and "hồ sơ" in d.lower(): st.session_state[p] = auto_id

    # Format Ngày tháng
    elif "date" in ft and len(digits) >= 8:
        st.session_state[ph] = f"{digits[:2]}/{digits[2:4]}/{digits[4:8]}"

# ==========================================
# CALLBACK: ÁP DỤNG PHƯƠNG ÁN (SỬ DỤNG REGEX THÔNG MINH)
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

        # Điền Chi Phí
        for i, (_, row) in enumerate(df_cp.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
            hm = str(row['Hạng mục']) if pd.notna(row.get('Hạng mục')) else ""
            nd = str(row['Nội dung chi tiết']) if pd.notna(row.get('Nội dung chi tiết')) else ""
            
            if re.match(fr"^(?:t\.cp|tcp|tlcp){i}$", k_var) and rate: st.session_state[ph] = rate
            if re.match(fr"^(?:hm_cp|hmcp){i}$", k_var) and hm: st.session_state[ph] = hm
            if re.match(fr"^(?:nd_cp|nd_chiphi|ndcp){i}$", k_var) and nd: st.session_state[ph] = nd

        # Điền Thu Nhập
        for i, (_, row) in enumerate(df_tn.iterrows(), 1):
            rate = str(round(float(row['Tỉ lệ']) * 100, 2)).rstrip('0').rstrip('.') if pd.notna(row.get('Tỉ lệ')) else ""
            hm = str(row['Hạng mục']) if pd.notna(row.get('Hạng mục')) else ""
            nd = str(row['Nội dung chi tiết']) if pd.notna(row.get('Nội dung chi tiết')) else ""
            
            if re.match(fr"^(?:t\.tn|ttn|tltn){i}$", k_var) and rate: st.session_state[ph] = rate
            if re.match(fr"^(?:hm_tn|hmtn){i}$", k_var) and hm: st.session_state[ph] = hm
            if re.match(fr"^(?:nd_tn|nd_thunhap|ndtn){i}$", k_var) and nd: st.session_state[ph] = nd

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
st.set_page_config(page_title="Hồ Sơ Tín Dụng Online", page_icon="🏦", layout="wide")

st.markdown("""
    <style>
    div[data-baseweb="input"] > div { border: 1px solid #2980B9; border-radius: 5px; }
    div[data-baseweb="textarea"] > div { border: 1px solid #2980B9; border-radius: 5px; }
    div[data-baseweb="select"] > div { border: 1px solid #2980B9; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

st.title("🏦 HỆ THỐNG KHỞI TẠO HỒ SƠ TÍN DỤNG")
st.caption("HOÀNG VIỆT HƯNG - QTDND PHÙNG HƯNG")

with st.sidebar:
    st.header("⚙️ CẤU HÌNH ĐẦU VÀO")
    excel_file = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    docx_file = st.file_uploader("2. File Mẫu Word.docx", type=["docx"])
    pa_file = st.file_uploader("3. File Phương án (Tùy chọn)", type=["xlsx"])
    
    st.divider()
    st.header("📸 QUÉT CCCD THÔNG MINH")
    
    if excel_file:
        with open("temp_data.xlsx", "wb") as f: f.write(excel_file.getbuffer())
        mapping, field_types, tabs_dict = read_mapping("temp_data.xlsx")
        
        # Danh sách 4 đối tượng cố định theo yêu cầu của Hùng
        subjects = [
            "Tất cả",
            "Thành viên",
            "Đồng vay vốn",
            "Người ủy quyền 1",
            "Người ủy quyền 2"
        ]
        
        target_person = st.selectbox("Quét thông tin cho đối tượng nào?", subjects)
        qr_img = st.file_uploader("Tải ảnh CCCD", type=["png", "jpg", "jpeg"])
        
        if qr_img and st.button("🚀 TIẾN HÀNH QUÉT ĐÍCH DANH", type="primary", use_container_width=True):
            with open("temp_qr.jpg", "wb") as f: f.write(qr_img.getbuffer())
            try:
                info = parse_cccd_payload(decode_qr_offline("temp_qr.jpg"))
                for ph, ds in mapping.items():
                    d_low = str(ds).lower()
                    
                    # CƠ CHẾ CHỐNG GHI ĐÈ
                    if is_target_match(target_person, d_low):
                        if is_cccd_desc(d_low): st.session_state[ph] = info.get("CCCD", "")
                        elif is_name_desc(d_low): 
                            st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif is_dob_desc(d_low): st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif is_addr_desc(d_low): st.session_state[ph] = info.get("Dia_chi", "")
                        elif is_issue_desc(d_low): st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                        elif is_gender_desc(d_low): st.session_state[ph] = info.get("Gioi_tinh", "")
                
                st.success(f"Đã điền thành công dữ liệu cho: {target_person}")
                st.rerun()
            except Exception as e: st.error(f"Lỗi đọc QR: {e}")

if excel_file and docx_file:
    # 1. Khởi tạo mọi key
    for ph in mapping.keys():
        if ph not in st.session_state: st.session_state[ph] = ""

    # 2. Logic Phương Án
    if pa_file:
        df_pa_list = pd.read_excel(pa_file, sheet_name="phuongan")
        df_pa_data = pd.read_excel(pa_file, sheet_name="data")
        pa_map = dict(zip(df_pa_list['Tên PA'], df_pa_list['Mã PA']))
        
        col_pa1, col_pa2 = st.columns([3, 1])
        with col_pa1:
            selected_pa = st.selectbox("⚡ ĐIỀN TỰ ĐỘNG PHƯƠNG ÁN VAY VỐN:", ["-- Chọn phương án --"] + list(pa_map.keys()))
        with col_pa2:
            st.write("") 
            if st.button("ÁP DỤNG PHƯƠNG ÁN", type="primary", use_container_width=True):
                if selected_pa != "-- Chọn phương án --":
                    apply_plan_callback(selected_pa, pa_map[selected_pa], df_pa_data, mapping)
                    st.rerun()

    st.divider()

    # 3. Vẽ Form & Tabs
    unique_tabs = list(dict.fromkeys(tabs_dict.values()))
    st_tabs = st.tabs(unique_tabs)
    tab_obj = dict(zip(unique_tabs, st_tabs))

    for tab_name in unique_tabs:
        with tab_obj[tab_name]:
            fields_in_tab = [(ph, ds) for ph, ds in mapping.items() if tabs_dict.get(ph) == tab_name]
            col1, col2 = st.columns(2)
            cols = [col1, col2]
            c_idx = 0
            
            for ph, ds in fields_in_tab:
                d_lower = str(ds).lower()
                is_tall = any(k in d_lower for k in LONG_TEXT_HINTS)
                cb_kwargs = {"ph": ph, "d_lower": d_lower, "field_types": field_types, "mapping": mapping}
                
                with cols[c_idx]:
                    if is_noicap_desc(d_lower) or is_loan_type_desc(d_lower):
                        opts = NOICAP_OPTIONS if is_noicap_desc(d_lower) else LOAN_TYPE_OPTIONS
                        st.selectbox(ds, [""] + opts, key=ph, on_change=process_field_change, kwargs=cb_kwargs)
                    elif is_tall:
                        st.text_area(ds, key=ph, height=120, on_change=process_field_change, kwargs=cb_kwargs)
                    else:
                        ft_val = str(field_types.get(ph, "")).lower()
                        if ft_val.startswith("dropdown:"):
                            opts = [o.strip() for o in ft_val.split(":")[1].split(",")]
                            st.selectbox(ds, [""] + opts, key=ph, on_change=process_field_change, kwargs=cb_kwargs)
                        else:
                            st.text_input(ds, key=ph, on_change=process_field_change, kwargs=cb_kwargs)
                c_idx = 1 - c_idx

    # 4. Xuất File
    st.divider()
    if st.button("🚀 TẠO HỒ SƠ WORD TỔNG HỢP", type="primary", use_container_width=True):
        ctx = {}
        for ph, ds in mapping.items():
            val = str(st.session_state[ph]).strip()
            
            if "\n" in val:
                rt = RichText()
                lines = val.split("\n")
                for i, line in enumerate(lines):
                    rt.add(line)
                    if i < len(lines) - 1: rt.add_line_break()
                ctx[var_name(ph)] = rt
            else:
                ctx[var_name(ph)] = val

        with open("temp_template.docx", "wb") as f: f.write(docx_file.getbuffer())
        out_stream = io.BytesIO()
        generate_word_document("temp_template.docx", out_stream, ctx)
        out_stream.seek(0)
        
        name_key = next((p for p, d in mapping.items() if is_name_desc(d.lower())), None)
        member_name = "HoSo"
        if name_key and st.session_state[name_key]:
            import unicodedata
            member_name = unicodedata.normalize('NFKD', st.session_state[name_key]).encode('ASCII', 'ignore').decode('utf-8')
            member_name = "".join(c for c in member_name if c.isalnum() or c in [' ', '-', '_']).replace(" ", "_")

        st.success("Tạo file thành công! Bạn có thể tải xuống ngay.")
        st.download_button("📥 TẢI XUỐNG BẢN WORD", data=out_stream, file_name=f"{member_name}_{datetime.now().strftime('%H%M')}.docx", type="primary")

else:
    st.info("💡 HƯỚNG DẪN: Hãy Upload file Data và Mẫu Word ở thanh điều khiển bên trái để nạp ứng dụng.")

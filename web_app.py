from __future__ import annotations

import io
import os
import re
import tempfile
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st
from docxtpl import RichText

from config import LONG_TEXT_HINTS
from core.document import FieldDefinition, MappingBundle, generate_word_document, load_mapping_bundle
from core.scanner import decode_qr_offline
from core.utils import (
    area_to_text_vi,
    ascii_safe_filename,
    extract_spell_target,
    filter_percent_text,
    format_area_input_text,
    format_money_input_text,
    format_money_vi,
    format_percent_vi,
    is_name_desc,
    money_to_text_vi,
    parse_cccd_payload,
    parse_decimal_vi,
    parse_number_vi,
    parse_percent_vi,
    validate_cccd_12_digits,
    var_name,
    vi_title_name,
)


st.set_page_config(page_title="Hồ Sơ Tín Dụng Pro", page_icon="📄", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.3rem; padding-bottom: 2rem;}
    .stTextInput input, .stTextArea textarea {font-size: 14px;}
    .small-note {font-size: 0.9rem; color: #667085;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 1. LOAD FILES CÓ CACHE + TEMPFILE AN TOÀN THEO SESSION
# ============================================================


def _write_bytes_to_tempfile(file_bytes: bytes, suffix: str) -> str:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        tmp.write(file_bytes)
        tmp.flush()
        return tmp.name
    finally:
        tmp.close()


@st.cache_data(show_spinner=False)
def load_mapping_from_bytes(file_bytes: bytes) -> MappingBundle:
    tmp_path = _write_bytes_to_tempfile(file_bytes, ".xlsx")
    try:
        return load_mapping_bundle(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


@st.cache_data(show_spinner=False)
def load_plan_from_bytes(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    bio = io.BytesIO(file_bytes)
    df_list = pd.read_excel(bio, sheet_name="phuongan")
    bio.seek(0)
    df_data = pd.read_excel(bio, sheet_name="data")
    return df_list, df_data


def decode_qr_from_upload(file_bytes: bytes, suffix: str = ".jpg") -> str:
    tmp_path = _write_bytes_to_tempfile(file_bytes, suffix)
    try:
        return decode_qr_offline(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


# ============================================================
# 2. HÀM XỬ LÝ TIỀN TỆ, TOÁN HỌC, BẰNG CHỮ
# ============================================================


def extract_number(val: Any) -> int:
    return int(parse_number_vi(val))


def extract_float(val: Any) -> float:
    return parse_decimal_vi(val)


def is_money_field(description: str, field_type: str) -> bool:
    desc = (description or "").lower()
    ft = (field_type or "").lower()
    return "money" in ft or any(
        keyword in desc
        for keyword in ["số tiền", "vốn", "thu nhập", "chi phí", "giá trị", "doanh thu", "lãi phải trả"]
    )


def is_rate_field(description: str, placeholder: str) -> bool:
    desc = (description or "").lower()
    key = var_name(placeholder).lower()
    return "lãi suất" in desc or "tỷ lệ" in desc or "ls" in key or "lai" in key


def calculate_globals(mapping: dict[str, str], field_types: dict[str, str]) -> None:
    """Tính lại các trường phụ thuộc sau mỗi lần nhập liệu.

    Bao gồm:
    - Tổng vốn = vốn xin vay + vốn tự có.
    - Lãi suất tháng = lãi suất năm / 12.
    - Số tiền chi phí/thu nhập theo tỷ lệ phương án.
    - Tiền/diện tích bằng chữ.
    """
    tvdt_keys: list[str] = []
    ls_thang_keys: list[str] = []
    vay_val = 0
    tuco_val = 0
    ls_nam_val = 0.0

    # Bước 1: tìm biến độc lập.
    for ph, ds in mapping.items():
        key = var_name(ph).lower().replace("_", "")
        desc = (ds or "").lower()
        value = str(st.session_state.get(ph, "")).strip()

        if any(x in key for x in ["vay", "xinvay", "mucvay"]) and not any(
            x in key for x in ["tong", "ls", "lai", "thoi", "ngay", "loai"]
        ):
            vay_val = max(vay_val, extract_number(value))
        elif any(x in key for x in ["tuco", "vonthamgia", "vontc"]):
            tuco_val = max(tuco_val, extract_number(value))
        elif any(x in key for x in ["tongvon", "tvdt", "tongmucdautu", "tongchiphi", "tongsovon"]):
            tvdt_keys.append(ph)
        elif ("ls" in key or "laisuat" in key) and "năm" in desc:
            ls_nam_val = max(ls_nam_val, parse_percent_vi(value))
        elif ("ls" in key or "laisuat" in key) and "tháng" in desc:
            ls_thang_keys.append(ph)

    # Bước 2: tổng vốn và lãi tháng.
    total_capital = vay_val + tuco_val
    if total_capital > 0:
        for key in tvdt_keys:
            st.session_state[key] = format_money_vi(total_capital)
    else:
        for key in tvdt_keys:
            total_capital = max(total_capital, extract_number(st.session_state.get(key, "0")))

    if ls_nam_val > 0:
        for key in ls_thang_keys:
            st.session_state[key] = format_percent_vi(ls_nam_val / 12.0)

    # Bước 3: tự động nhân tỷ lệ phương án với tổng vốn.
    if total_capital > 0:
        normalized_keys = {var_name(p).lower().replace("_", ""): p for p in mapping}
        for ph, ds in mapping.items():
            key = var_name(ph).lower().replace("_", "")
            desc = (ds or "").lower()
            match = re.search(r"(\d+)$", key)
            if not match or "lệ" in desc:
                continue

            idx = match.group(1)
            base = key[: -len(idx)]
            amount_bases = [
                "stcp",
                "sotiencp",
                "cp",
                "chiphi",
                "tien",
                "sotien",
                "sttn",
                "sotientn",
                "tn",
                "thunhap",
                "dt",
                "stdt",
                "doanhthu",
            ]
            if base not in amount_bases:
                continue

            prefix = None
            if any(x in base for x in ["cp", "chiphi"]):
                prefix = "tlcp"
            elif any(x in base for x in ["tn", "thunhap", "dt", "doanhthu"]):
                prefix = "tltn"
            elif "chi phí" in desc:
                prefix = "tlcp"
            elif "thu nhập" in desc or "doanh thu" in desc:
                prefix = "tltn"

            if not prefix:
                continue

            candidate_keys = [f"{prefix}{idx}", f"t{prefix[2:]}{idx}"]
            rate_ph = next((normalized_keys[k] for k in candidate_keys if k in normalized_keys), None)
            if not rate_ph:
                continue

            rate = parse_percent_vi(st.session_state.get(rate_ph, "0"))
            if rate > 0:
                st.session_state[ph] = format_money_vi(total_capital * rate / 100)

    # Bước 4: đọc tiền/diện tích bằng chữ.
    for ph, field_type in field_types.items():
        ft = str(field_type or "").strip().lower()
        target = extract_spell_target(ft)
        if not target:
            continue

        target_placeholder = next((p for p in mapping if var_name(p) == target), target)
        if target_placeholder not in st.session_state and target not in st.session_state:
            continue

        target_key = target_placeholder if target_placeholder in st.session_state else target
        source_value = st.session_state.get(ph, "")

        try:
            if ft.startswith("spell_area"):
                text = area_to_text_vi(source_value)
            else:
                text = money_to_text_vi(source_value)
            if text:
                st.session_state[target_key] = text
        except Exception:
            # Không chặn app nếu thiếu num2words hoặc dữ liệu chưa hợp lệ.
            pass


# Alias giữ tương thích với các đoạn gọi cũ.
execute_financial_engine = calculate_globals


# ============================================================
# 3. CALLBACK NHẬP LIỆU
# ============================================================


def field_callback(ph: str, ds: str, field_types: dict[str, str], mapping: dict[str, str]) -> None:
    val = str(st.session_state.get(ph, ""))
    desc = (ds or "").lower()
    field_type = str(field_types.get(ph, "")).strip().lower()

    if is_rate_field(ds, ph):
        st.session_state[ph] = filter_percent_text(val)
    elif is_money_field(ds, field_type):
        st.session_state[ph] = format_money_input_text(val)
    elif "area" in field_type or field_type.startswith("spell_area:"):
        clean_val = format_area_input_text(val)
        st.session_state[ph] = clean_val
        target = extract_spell_target(field_type)
        if target:
            target_placeholder = next((p for p in mapping if var_name(p) == target), target)
            if target_placeholder in st.session_state or target in st.session_state:
                try:
                    st.session_state[target_placeholder if target_placeholder in st.session_state else target] = area_to_text_vi(clean_val)
                except Exception:
                    pass
    elif ("họ tên" in desc or "họ và tên" in desc) and val:
        st.session_state[ph] = vi_title_name(val)

    calculate_globals(mapping, field_types)


# ============================================================
# 4. PHÂN LOẠI FIELD CCCD + FIREWALL PHÂN VÙNG DỮ LIỆU
# ============================================================


def get_cccd_field_type(description: str) -> str | None:
    """Từ chối các trường tài sản/đất để tránh nhầm ngày cấp GCN với ngày cấp CCCD."""
    desc = (description or "").lower()
    if any(x in desc for x in ["sổ", "đất", "gcn", "giấy chứng nhận", "ruộng", "xe", "tsđb", "tài sản"]):
        return None
    if any(x in desc for x in ["ngày cấp", "cấp ngày", "ngay cap"]):
        return "issue_date"
    if any(x in desc for x in ["cccd", "cmnd", "căn cước", "chứng minh", "định danh"]):
        return "cccd"
    if any(x in desc for x in ["họ tên", "họ và tên", "người"]):
        return "name"
    if any(x in desc for x in ["ngày sinh", "sinh ngày", "năm sinh"]):
        return "dob"
    if any(x in desc for x in ["địa chỉ", "thường trú", "cư trú", "hộ khẩu", "nơi ở"]):
        return "address"
    if any(x in desc for x in ["giới tính", "nam/nữ"]):
        return "gender"
    return None


def firewall_qr_target(placeholder: str, description: str) -> str:
    text = f"{(description or '').lower()} | {var_name(placeholder).lower()}"
    if any(x in text for x in ["thụ hưởng", "thuhuong", "_th", "nhận tiền"]):
        return "Người thụ hưởng"
    if any(x in text for x in ["ủy quyền 2", "uyquyen2", "uq2"]):
        return "Người ủy quyền 2"
    if any(x in text for x in ["ủy quyền 1", "ủy quyền", "uyquyen1", "uq1"]):
        return "Người ủy quyền 1"
    if any(x in text for x in ["đồng vay 2", "dongvay2", "dv2"]):
        return "Người đồng vay vốn 2"
    if any(x in text for x in ["đồng vay", "cùng vay", "dongvay", "dv1", "_dv", "dv_"]):
        return "Người đồng vay vốn 1"
    if any(x in text for x in ["vợ", "vo", "chồng", "chong", "bảo lãnh"]):
        return "Khác"
    return "Thành viên"


# ============================================================
# 5. PHƯƠNG ÁN VAY VỐN
# ============================================================


def validate_plan_columns(df_list: pd.DataFrame, df_data: pd.DataFrame) -> list[str]:
    errors: list[str] = []
    required_list = {"Tên PA", "Mã PA"}
    required_data = {"Mã phương án", "Loại", "Tỉ lệ", "Hạng mục", "Nội dung chi tiết"}
    missing_list = required_list - set(df_list.columns)
    missing_data = required_data - set(df_data.columns)
    if missing_list:
        errors.append(f"Sheet 'phuongan' thiếu cột: {', '.join(sorted(missing_list))}")
    if missing_data:
        errors.append(f"Sheet 'data' thiếu cột: {', '.join(sorted(missing_data))}")
    return errors


def apply_plan_to_session(
    selected_plan_name: str,
    df_list: pd.DataFrame,
    df_data: pd.DataFrame,
    mapping: dict[str, str],
    field_types: dict[str, str],
) -> None:
    pa_map = dict(zip(df_list["Tên PA"], df_list["Mã PA"]))
    plan_code = pa_map[selected_plan_name]
    rows = df_data[df_data["Mã phương án"].astype(str).str.strip().str.lower() == str(plan_code).lower()]

    df_cp = rows[rows["Loại"].astype(str).str.contains("chi|phí|cp", case=False, na=False)].reset_index(drop=True)
    df_tn = rows[rows["Loại"].astype(str).str.contains("thu|doanh|lợi|tn|dt", case=False, na=False)].reset_index(drop=True)

    for ph, ds in mapping.items():
        key = var_name(ph).lower().replace("_", "")
        desc = (ds or "").lower()

        if "tenphuongan" in key or "tenpa" in key:
            st.session_state[ph] = selected_plan_name
            continue

        match = re.search(r"(\d+)$", key)
        if not match:
            continue

        idx = int(match.group(1)) - 1
        base = key[: -len(match.group(1))]
        is_cp = any(x in base for x in ["cp", "chiphi"]) or "chi phí" in desc
        is_tn = (
            any(x in base for x in ["tn", "thunhap", "dt", "doanhthu"])
            or "thu nhập" in desc
            or "doanh thu" in desc
        )
        target_df = df_cp if is_cp else df_tn if is_tn else None
        if target_df is None or not (0 <= idx < len(target_df)):
            continue

        row = target_df.iloc[idx]
        raw_rate = row.get("Tỉ lệ")
        if pd.notna(raw_rate):
            try:
                # File phương án thường lưu tỷ lệ dạng 0.1 = 10%.
                rate = format_percent_vi(float(raw_rate) * 100)
            except Exception:
                rate = str(raw_rate)
        else:
            rate = ""

        hm = str(row.get("Hạng mục", "")).strip() if pd.notna(row.get("Hạng mục")) else ""
        nd = str(row.get("Nội dung chi tiết", "")).strip() if pd.notna(row.get("Nội dung chi tiết")) else ""

        if any(x in base for x in ["tl", "tyle", "tcp", "ttn"]):
            st.session_state[ph] = rate
        elif any(x in base for x in ["hm", "hangmuc"]):
            st.session_state[ph] = hm
        elif any(x in base for x in ["nd", "noidung", "chitiet"]):
            st.session_state[ph] = nd
        elif base in ["chiphi", "cp", "thunhap", "tn", "doanhthu", "dt"]:
            st.session_state[ph] = nd if not hm else f"{hm}: {nd}"

    calculate_globals(mapping, field_types)


# ============================================================
# 6. RENDER FORM
# ============================================================


def init_session_fields(fields: list[FieldDefinition]) -> None:
    for field in fields:
        if field.placeholder not in st.session_state:
            st.session_state[field.placeholder] = ""


def render_field(field: FieldDefinition, mapping: dict[str, str], field_types: dict[str, str]) -> None:
    kwargs = {
        "ph": field.placeholder,
        "ds": field.description,
        "field_types": field_types,
        "mapping": mapping,
    }
    label = field.description or field.key
    disabled = field.readonly or not field.visible

    if field.options:
        current_value = str(st.session_state.get(field.placeholder, ""))
        options = [""] + field.options
        index = options.index(current_value) if current_value in options else 0
        st.selectbox(
            label,
            options,
            key=field.placeholder,
            index=index,
            disabled=disabled,
            on_change=field_callback,
            kwargs=kwargs,
        )
    elif field.long_text or any(k in (field.description or "").lower() for k in LONG_TEXT_HINTS):
        st.text_area(
            label,
            key=field.placeholder,
            height=120,
            disabled=disabled,
            on_change=field_callback,
            kwargs=kwargs,
        )
    else:
        st.text_input(
            label,
            key=field.placeholder,
            disabled=disabled,
            on_change=field_callback,
            kwargs=kwargs,
        )


def render_tabs(bundle: MappingBundle) -> None:
    mapping = bundle.mapping
    field_types = bundle.field_types
    tabs = bundle.tab_sequence or list(dict.fromkeys(bundle.tabs.values()))
    st_tabs = st.tabs(tabs)

    for tab_index, tab_name in enumerate(tabs):
        with st_tabs[tab_index]:
            fields = [field for field in bundle.fields if field.tab == tab_name and field.visible]
            fields.sort(key=lambda item: (item.order, item.field_row or 9999, item.field_col or 9999, item.description))

            if not fields:
                st.info("Tab này chưa có trường hiển thị.")
                continue

            # Nếu file Data.xlsx có metadata field_col thì ưu tiên chia cột theo metadata.
            if bundle.uses_layout_metadata and any(field.field_col for field in fields):
                max_col = max(field.field_col or 1 for field in fields)
                cols = st.columns(min(max_col, 4))
                for field in fields:
                    col_index = min((field.field_col or 1) - 1, len(cols) - 1)
                    with cols[col_index]:
                        render_field(field, mapping, field_types)
            else:
                c1, c2 = st.columns(2)
                for idx, field in enumerate(fields):
                    with c1 if idx % 2 == 0 else c2:
                        render_field(field, mapping, field_types)


# ============================================================
# 7. XUẤT WORD
# ============================================================


def build_docx_context(mapping: dict[str, str]) -> dict[str, Any]:
    ctx: dict[str, Any] = {}
    for ph in mapping:
        value = str(st.session_state.get(ph, "")).strip()
        key = var_name(ph)
        if "\n" in value:
            rich = RichText()
            lines = value.split("\n")
            for idx, line in enumerate(lines):
                rich.add(line)
                if idx < len(lines) - 1:
                    rich.add_line_break()
            ctx[key] = rich
        else:
            ctx[key] = value
    return ctx


def find_member_name(mapping: dict[str, str]) -> str:
    name_key = next((p for p, d in mapping.items() if is_name_desc(d)), None)
    if name_key and st.session_state.get(name_key):
        return ascii_safe_filename(str(st.session_state[name_key]), fallback="HoSo")
    return "HoSo"


def generate_docx_download(file_word_bytes: bytes, mapping: dict[str, str]) -> tuple[bytes, str]:
    calculate_globals(mapping, st.session_state.get("field_types", {}))
    context = build_docx_context(mapping)
    template_path = _write_bytes_to_tempfile(file_word_bytes, ".docx")
    output = io.BytesIO()
    try:
        generate_word_document(template_path, output, context)
    finally:
        try:
            os.unlink(template_path)
        except OSError:
            pass
    output.seek(0)
    member_name = find_member_name(mapping)
    file_name = f"{member_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    return output.getvalue(), file_name


# ============================================================
# 8. GIAO DIỆN CHÍNH
# ============================================================


st.title("📄 Hồ Sơ Tín Dụng Pro")
st.caption("Nhập hồ sơ vay vốn, quét CCCD, áp dụng phương án và xuất mẫu Word.")

with st.sidebar:
    st.header("⚙️ Cấu hình hệ thống")
    file_data = st.file_uploader("1. File Data.xlsx", type=["xlsx"])
    file_word = st.file_uploader("2. Mẫu Word.docx", type=["docx"])
    file_plan = st.file_uploader("3. Phương án.xlsx", type=["xlsx"])

if not file_data:
    st.info("Vui lòng tải lên File Data.xlsx để bắt đầu nhập hồ sơ.")
    st.stop()

try:
    bundle = load_mapping_from_bytes(file_data.getvalue())
except Exception as exc:
    st.error(f"Không đọc được File Data.xlsx: {exc}")
    st.stop()

mapping = bundle.mapping
field_types = bundle.field_types
st.session_state.field_types = field_types
init_session_fields(bundle.fields)
calculate_globals(mapping, field_types)

with st.sidebar:
    st.divider()
    st.header("🪪 Quét CCCD song song")
    subjects = [
        "Thành viên",
        "Người đồng vay vốn 1",
        "Người đồng vay vốn 2",
        "Người ủy quyền 1",
        "Người ủy quyền 2",
        "Người thụ hưởng",
    ]

    for target in subjects:
        with st.expander(target, expanded=False):
            qr_img = st.file_uploader(
                "Ảnh CCCD",
                type=["png", "jpg", "jpeg"],
                key=f"qr_{target}",
            )
            if qr_img and st.button("Quét dữ liệu", key=f"btn_{target}", use_container_width=True):
                try:
                    suffix = os.path.splitext(qr_img.name or "qr.jpg")[1] or ".jpg"
                    payload = decode_qr_from_upload(qr_img.getvalue(), suffix=suffix)
                    info = parse_cccd_payload(payload)
                    if info.get("CCCD") and not validate_cccd_12_digits(info.get("CCCD", "")):
                        st.warning("QR đọc được nhưng số CCCD không đủ 12 chữ số. Vui lòng kiểm tra lại.")

                    count = 0
                    for ph, ds in mapping.items():
                        if firewall_qr_target(ph, ds) != target:
                            continue
                        cccd_field = get_cccd_field_type(ds)
                        if cccd_field == "cccd":
                            st.session_state[ph] = info.get("CCCD", "")
                        elif cccd_field == "name":
                            st.session_state[ph] = vi_title_name(info.get("Ho_va_ten", ""))
                        elif cccd_field == "dob":
                            st.session_state[ph] = info.get("Ngay_thang_nam_sinh", "")
                        elif cccd_field == "address":
                            st.session_state[ph] = info.get("Dia_chi", "")
                        elif cccd_field == "issue_date":
                            st.session_state[ph] = info.get("Ngay_cap_CCCD", "")
                        elif cccd_field == "gender":
                            st.session_state[ph] = info.get("Gioi_tinh", "")
                        if cccd_field:
                            count += 1

                    calculate_globals(mapping, field_types)
                    st.success(f"Nạp an toàn {count} trường cho: {target}")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Lỗi đọc thẻ: {exc}")

# Phần áp dụng phương án.
if file_plan:
    st.divider()
    try:
        df_list, df_data = load_plan_from_bytes(file_plan.getvalue())
        plan_errors = validate_plan_columns(df_list, df_data)
        if plan_errors:
            for err in plan_errors:
                st.warning(err)
        else:
            pa_map = dict(zip(df_list["Tên PA"], df_list["Mã PA"]))
            c1, c2 = st.columns([3, 1])
            with c1:
                sel_pa = st.selectbox("⚡ Điền tự động phương án", ["-- Trống --"] + list(pa_map.keys()))
            with c2:
                st.write("")
                st.write("")
                if st.button("Áp dụng", type="primary", use_container_width=True):
                    if sel_pa != "-- Trống --":
                        apply_plan_to_session(sel_pa, df_list, df_data, mapping, field_types)
                        st.success(f"Đã áp dụng phương án: {sel_pa}")
                        st.rerun()
    except Exception as exc:
        st.error(f"Không đọc được file Phương án.xlsx: {exc}")

# Render form.
st.divider()
render_tabs(bundle)

# Xuất Word.
st.divider()
if not file_word:
    st.warning("Tải lên Mẫu Word.docx để xuất hồ sơ.")
else:
    if st.button("📤 XUẤT HỒ SƠ WORD TỔNG HỢP", type="primary", use_container_width=True):
        try:
            docx_bytes, output_name = generate_docx_download(file_word.getvalue(), mapping)
            st.download_button(
                "⬇️ TẢI FILE WORD (.DOCX)",
                data=docx_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Không xuất được Word: {exc}")

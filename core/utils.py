from __future__ import annotations

import math
import re
from typing import Iterable

from config import LONG_TEXT_HINTS


class OptionalDependencyError(RuntimeError):
    """Raised when a runtime-only dependency is missing."""


def safe_filename(s: str) -> str:
    return re.sub(r'[<>:"/\\|?*\n\r\t]', "_", (s or "").strip()) or "KetQua"


def var_name(placeholder: str) -> str:
    text = (placeholder or "").strip()
    match = re.match(r"^\{\{\s*(.*?)\s*\}\}$", text) or re.match(r"^\{\s*(.*?)\s*\}$", text)
    return (match.group(1).strip() if match else text)


def ddmmyyyy_to_vn(date_str: str) -> str:
    digits = re.sub(r"\D", "", str(date_str or ""))
    if len(digits) != 8:
        return str(date_str or "")
    return f"{digits[0:2]}/{digits[2:4]}/{digits[4:8]}"


def parse_cccd_payload(payload: str) -> dict:
    parts = [part.strip() for part in (payload or "").split("|")]
    if len(parts) < 7:
        raise ValueError("QR CCCD không đúng định dạng hoặc thiếu thông tin.")
    return {
        "CCCD": parts[0],
        "So_CMND_cu": parts[1],
        "Ho_va_ten": parts[2],
        "Ngay_thang_nam_sinh": ddmmyyyy_to_vn(parts[3]),
        "Gioi_tinh": parts[4],
        "Dia_chi": parts[5],
        "Ngay_cap_CCCD": ddmmyyyy_to_vn(parts[6]),
        "raw": payload,
    }


def normalize_person_key(desc: str) -> str:
    text = (desc or "").lower()
    text = re.sub(r"^\d+[\.\-\)\:]*\s*", "", text)
    text = text.replace("uỷ quyền", "ủy quyền")
    text = re.sub(r"\bđvv\b", "đồng vay vốn", text)
    text = re.sub(r"\buq\b", "ủy quyền", text)

    if "thành viên" in text:
        return "thành viên"

    match = re.search(r"(đồng\s*vay\s*vốn|ủy\s*quyền)\s*([0-9]+)", text)
    if match:
        return f"{match.group(1)} {match.group(2)}"

    remove_words = [
        "số thẻ cccd", "số cccd", "cccd", "số thẻ", "cmnd", "số cmnd",
        "ngày tháng năm sinh", "ngày sinh", "năm sinh",
        "họ và tên", "họ tên", "tên",
        "ngày cấp", "nơi cấp", "địa chỉ thường trú", "thường trú", "địa chỉ", "giới tính",
    ]
    for word in remove_words:
        text = text.replace(word, "")

    text = re.sub(r"\s+", " ", text).strip()
    return text or "default"


def is_issue_desc(desc: str) -> bool:
    desc_lower = (desc or "").lower()
    if "qsdđ" in desc_lower or "gcn" in desc_lower or "đất" in desc_lower:
        return False
    return "ngày cấp" in desc_lower


def is_noicap_desc(desc: str) -> bool:
    desc_lower = (desc or "").lower()
    if "qsdđ" in desc_lower or "gcn" in desc_lower or "đất" in desc_lower:
        return False
    return "nơi cấp" in desc_lower


def is_cccd_desc(desc: str) -> bool:
    desc_lower = (desc or "").lower()
    return ("cccd" in desc_lower or "căn cước" in desc_lower or "cmnd" in desc_lower) and not is_issue_desc(desc) and not is_noicap_desc(desc)


def is_name_desc(desc: str) -> bool:
    desc_lower = (desc or "").lower()
    if "phương án" in desc_lower or "tên pa" in desc_lower:
        return False
    return any(keyword in desc_lower for keyword in ["tên", "họ tên", "họ và tên"])


def is_dob_desc(desc: str) -> bool:
    desc_lower = (desc or "").lower()
    return "ngày sinh" in desc_lower or "năm sinh" in desc_lower


def is_addr_desc(desc: str) -> bool:
    desc_lower = (desc or "").lower()
    return "địa chỉ" in desc_lower or "thường trú" in desc_lower


def is_gender_desc(desc: str) -> bool:
    return "giới tính" in (desc or "").lower()


def is_loan_type_desc(desc: str) -> bool:
    return "loại cho vay" in (desc or "").lower()


def parse_number_vi(value: str) -> float:
    text = str(value or "").strip().replace(".", "").replace(",", "")
    text = re.sub(r"[^0-9\-]", "", text)
    return float(text) if text not in ("", "-") else 0.0


def parse_decimal_vi(value: str) -> float:
    text = str(value or "").strip().replace(".", "").replace(",", ".")
    text = re.sub(r"[^0-9\.\-]", "", text)
    try:
        return float(text)
    except Exception:
        return 0.0


def parse_percent_vi(value: str) -> float:
    text = str(value or "").strip().replace("%", "").replace(",", ".")
    text = re.sub(r"[^0-9\.\-]", "", text)
    try:
        return float(text)
    except Exception:
        return 0.0


def parse_rate_vi(value: str) -> float:
    return parse_percent_vi(value)


def format_money_vi(value: float | int) -> str:
    if value in (None, ""):
        return ""
    try:
        return f"{int(round(float(value))):,}".replace(",", ".")
    except Exception:
        return ""


def format_decimal_vi(value: float, decimals: int = 2) -> str:
    if value is None or (isinstance(value, float) and not math.isfinite(value)):
        return ""
    text = f"{float(value):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return text.rstrip("0").rstrip(",")


def vi_title_name(value: str) -> str:
    if not value:
        return ""
    return " ".join(word.capitalize() if word else "" for word in value.split(" "))


def validate_cccd_12_digits(value: str) -> bool:
    digits = "".join(char for char in str(value or "") if char.isdigit())
    return len(digits) == 12


def coerce_date_text(value: str) -> str:
    digits = re.sub(r"\D", "", str(value or ""))[:8]
    if len(digits) >= 5:
        return f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"
    if len(digits) >= 3:
        return f"{digits[:2]}/{digits[2:]}"
    return digits


def filter_number_text(value: str) -> str:
    return "".join(char for char in str(value or "") if char.isdigit() or char in ".,-")


def filter_percent_text(value: str) -> str:
    text = str(value or "").replace("%", "").replace(",", ".")
    return "".join(char for char in text if char.isdigit() or char == ".")


def format_money_input_text(value: str) -> str:
    digits = "".join(char for char in str(value or "") if char.isdigit())
    return f"{int(digits):,}".replace(",", ".") if digits else ""


def format_area_input_text(value: str) -> str:
    raw = str(value or "").replace(".", "")
    clean = "".join(char for char in raw if char.isdigit() or char == ",")
    if not clean:
        return ""
    parts = clean.split(",")
    int_part = f"{int(parts[0]):,}".replace(",", ".") if parts[0] else ""
    if len(parts) > 1:
        return f"{int_part},{parts[1]}"
    return int_part if not clean.endswith(",") else f"{int_part},"

def normalize_square_meter_text(value: str) -> str:
    text = str(value or "")
    replacements = {
        r"(?i)\bmm\s*2\b": "mm²",
        r"(?i)\bcm\s*2\b": "cm²",
        r"(?i)\bdm\s*2\b": "dm²",
        r"(?i)\bm\s*2\b": "m²",
        r"(?i)\bhm\s*2\b": "hm²",
        r"(?i)\bkm\s*2\b": "km²",
    }
    for pattern, repl in replacements.items():
        text = re.sub(pattern, repl, text)
    return text


def looks_like_long_text(description: str, field_type: str = "") -> bool:
    desc_lower = (description or "").lower()
    field_lower = (field_type or "").lower().strip()
    return field_lower == "textarea" or any(keyword in desc_lower for keyword in LONG_TEXT_HINTS)


def split_dropdown_options(field_type: str) -> list[str]:
    text = str(field_type or "").strip()
    if not text.lower().startswith("dropdown:"):
        return []
    payload = re.split(r"^dropdown\s*:", text, flags=re.IGNORECASE)[1].strip()
    if "/" in payload and "," not in payload:
        return [item.strip() for item in payload.split("/") if item.strip()]
    return [item.strip() for item in payload.split(",") if item.strip()]


def extract_spell_target(field_type: str) -> str:
    text = str(field_type or "").strip()
    match = re.match(r"^spell(?:_area)?\s*:\s*(.+)$", text, flags=re.IGNORECASE)
    if not match:
        return ""
    return var_name(match.group(1).strip())


def normalize_type(field_type: str, spell: str = "") -> str:
    text = str(field_type or "").strip()
    if not text and spell:
        return str(spell).strip()
    return text


def percent_or_blank(value: float) -> str:
    if value is None or not math.isfinite(value):
        return ""
    return f"{value * 100:.1f}%"


def ratio(numerator: float, denominator: float) -> float:
    if denominator in (0, 0.0):
        return 0.0
    return numerator / denominator


def filled_ratio(values: Iterable[str]) -> float:
    values = list(values)
    if not values:
        return 0.0
    filled = sum(1 for value in values if str(value or "").strip())
    return filled / len(values)


def safe_float(value: str | float | int) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return parse_decimal_vi(str(value))
    except Exception:
        return 0.0

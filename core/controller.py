from __future__ import annotations

import json
import os
import re
from datetime import date, datetime, timedelta
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from config import (
    DEFAULT_ASSESSMENT,
    DEFAULT_COMMENTS,
    DEFAULT_DOCUMENT_STATUS,
    DEFAULT_METADATA,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_PHUONGAN,
    DOSSIER_DIR,
    DRAFT_FILE,
    LOAN_TYPE_OPTIONS,
    MANDATORY_DOCUMENTS,
    NOICAP_OPTIONS,
    REQUIRED_PLACEHOLDER_KEYS,
    SAMPLE_CONTEXT,
)
from core.document import FieldDefinition, MappingBundle, generate_word_document, load_mapping_bundle
from core.scanner import decode_qr_offline
from core.utils import (
    coerce_date_text,
    filter_number_text,
    filter_percent_text,
    format_area_input_text,
    format_decimal_vi,
    format_money_input_text,
    format_money_vi,
    is_addr_desc,
    is_cccd_desc,
    is_dob_desc,
    is_gender_desc,
    is_issue_desc,
    is_loan_type_desc,
    is_name_desc,
    is_noicap_desc,
    looks_like_long_text,
    normalize_person_key,
    normalize_square_meter_text,
    parse_cccd_payload,
    parse_decimal_vi,
    parse_number_vi,
    parse_percent_vi,
    parse_rate_vi,
    percent_or_blank,
    ratio,
    safe_filename,
    validate_cccd_12_digits,
    var_name,
    vi_title_name,
)


@dataclass(slots=True)
class ValidationResult:
    warnings: list[str]
    errors: list[str]


class LoanApplicationController:
    """State, calculations and export logic for the PySide6 workbench."""

    def __init__(self, xlsx_path: str, docx_path: str, output_dir: str | None = None, phuongan_path: str | None = None):
        self.xlsx_path = xlsx_path
        self.docx_path = docx_path
        self.output_dir = output_dir or DEFAULT_OUTPUT_DIR
        self.phuongan_path = phuongan_path or DEFAULT_PHUONGAN

        self.bundle: MappingBundle | None = None
        self.fields: list[FieldDefinition] = []
        self.fields_by_placeholder: dict[str, FieldDefinition] = {}
        self.fields_by_key: dict[str, FieldDefinition] = {}
        self.grouped_fields: dict[str, list[FieldDefinition]] = {}

        self.values: dict[str, str] = {}
        self.metadata: dict[str, str] = dict(DEFAULT_METADATA)
        self.assessment: dict[str, str] = dict(DEFAULT_ASSESSMENT)
        self.documents: dict[str, str] = dict(DEFAULT_DOCUMENT_STATUS)
        self.comments: dict[str, str] = dict(DEFAULT_COMMENTS)

        self.person_groups: dict[str, dict[str, str]] = {}
        self.rate_keys: dict[str, dict[int, str]] = {"cp": {}, "tn": {}}
        self.amount_keys: dict[str, dict[int, str]] = {"cp": {}, "tn": {}}
        self.total_keys: dict[str, str | None] = {"cp": None, "tn": None, "profit": None}
        self.loan_keys: dict[str, str | None] = {"amount": None, "term": None, "rate_year": None, "rate_month": None, "own_capital": None}
        self.interest_key: str | None = None
        self.valuation_keys: dict[str, str | None] = {
            "area_residential": None,
            "price_residential": None,
            "total_residential": None,
            "area_03": None,
            "price_03": None,
            "total_03": None,
            "total_valuation": None,
            "total_others": None,
        }
        self.phuongan_map: dict[str, str] = {}
        self.phuongan_data: pd.DataFrame | None = None
        self.audit_log: list[dict[str, str]] = []

    def load_schema(self) -> None:
        self.bundle = load_mapping_bundle(self.xlsx_path)
        self.fields = list(self.bundle.fields)
        self.fields_by_placeholder = {field.placeholder: field for field in self.fields}
        self.fields_by_key = {field.key: field for field in self.fields}
        self.grouped_fields = {}
        self.values = {field.placeholder: "" for field in self.fields}
        self.person_groups.clear()
        self.rate_keys = {"cp": {}, "tn": {}}
        self.amount_keys = {"cp": {}, "tn": {}}
        self.total_keys = {"cp": None, "tn": None, "profit": None}
        self.loan_keys = {"amount": None, "term": None, "rate_year": None, "rate_month": None, "own_capital": None}
        self.interest_key = None
        self.valuation_keys = {key: None for key in self.valuation_keys}

        for field in self.fields:
            self.grouped_fields.setdefault(field.tab, []).append(field)
            self._register_person_field(field)
            self._register_financial_field(field)

        self._load_phuongan_if_available()

    def apply_sample_data(self) -> None:
        for key, value in SAMPLE_CONTEXT.items():
            field = self.fields_by_key.get(key)
            if field:
                self.values[field.placeholder] = value
        self.metadata.update(DEFAULT_METADATA)
        self.assessment.update(DEFAULT_ASSESSMENT)
        self.documents.update(DEFAULT_DOCUMENT_STATUS)
        self.comments.update(DEFAULT_COMMENTS)
        self.values.update(self.calculate_all_updates())


    def _record_event(self, action: str, detail: str, source: str = "system") -> None:
        entry = {
            "time": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "action": action,
            "detail": detail,
            "source": source,
        }
        self.audit_log.append(entry)
        self.audit_log = self.audit_log[-300:]

    def recent_audit_log(self, limit: int = 12) -> list[dict[str, str]]:
        return list(reversed(self.audit_log[-limit:]))

    def current_payload(self) -> dict:
        return {
            "xlsx_path": self.xlsx_path,
            "docx_path": self.docx_path,
            "output_dir": self.output_dir,
            "phuongan_path": self.phuongan_path,
            "values": self.values,
            "metadata": self.metadata,
            "assessment": self.assessment,
            "documents": self.documents,
            "comments": self.comments,
            "audit_log": self.audit_log,
        }

    def _record_id(self) -> str:
        current = safe_filename(self.metadata.get("record_id", "") or self.metadata.get("application_id", "") or self.collect_context().get("tenthanhvien", "HoSo"))
        self.metadata["record_id"] = current
        return current

    def save_record(self, directory: str = DOSSIER_DIR) -> str:
        Path(directory).mkdir(parents=True, exist_ok=True)
        payload = self.current_payload()
        payload["saved_at"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        path = Path(directory) / f"{self._record_id()}.json"
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self._record_event("save_record", f"Lưu hồ sơ {path.name}")
        return str(path)

    def list_records(self, directory: str = DOSSIER_DIR) -> list[dict[str, str]]:
        folder = Path(directory)
        if not folder.exists():
            return []
        items: list[dict[str, str]] = []
        for path in sorted(folder.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True):
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                continue
            values = payload.get("values", {})
            metadata = payload.get("metadata", {})
            items.append({
                "path": str(path),
                "application_id": metadata.get("application_id", path.stem),
                "applicant_name": values.get(self.fields_by_key.get("tenthanhvien").placeholder, "") if self.fields_by_key.get("tenthanhvien") else values.get("{{tenthanhvien}}", ""),
                "status": metadata.get("status", "Nháp"),
                "saved_at": payload.get("saved_at", datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S")),
            })
        return items

    def load_record(self, path: str) -> bool:
        if not path or not os.path.exists(path):
            return False
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
        self.metadata.update(payload.get("metadata", {}))
        self.assessment.update(payload.get("assessment", {}))
        self.documents.update(payload.get("documents", {}))
        self.comments.update(payload.get("comments", {}))
        self.audit_log = list(payload.get("audit_log", []))
        raw_values = payload.get("values", {})
        for placeholder in self.values:
            self.values[placeholder] = str(raw_values.get(placeholder, self.values[placeholder] or ""))
        self.values.update(self.calculate_all_updates())
        self._record_event("open_record", f"Mở hồ sơ {Path(path).name}")
        return True

    def _extract_birth_year(self, text: str) -> int | None:
        digits = re.sub(r"\D", "", str(text or ""))
        if len(digits) >= 4:
            year = int(digits[-4:])
            if 1900 <= year <= date.today().year:
                return year
        return None









    def save_draft(self, path: str = DRAFT_FILE) -> None:
        payload = self.current_payload()
        with open(path, "w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)

    def restore_draft(self, path: str = DRAFT_FILE) -> bool:
        if not os.path.exists(path):
            return False
        with open(path, "r", encoding="utf-8") as file:
            payload = json.load(file)
        self.metadata.update(payload.get("metadata", {}))
        self.assessment.update(payload.get("assessment", {}))
        self.documents.update(payload.get("documents", {}))
        self.comments.update(payload.get("comments", {}))
        self.audit_log = list(payload.get("audit_log", []))
        raw_values = payload.get("values", {})
        for placeholder in self.values:
            self.values[placeholder] = str(raw_values.get(placeholder, self.values[placeholder] or ""))
        self.values.update(self.calculate_all_updates())
        return True

    def export_word(self) -> str:
        context = self.collect_context()
        member = safe_filename(context.get("tenthanhvien", ""))
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        output_path = os.path.join(self.output_dir, f"{member or 'LoanApplication'}.docx")
        generate_word_document(self.docx_path, output_path, context)
        self._record_event("export_word", f"Xuất Word {Path(output_path).name}")
        return output_path

    def set_field_value(self, placeholder: str, value: str, source: str = "user") -> dict[str, str]:
        old_value = str(self.values.get(placeholder, ""))
        self.values[placeholder] = value
        if old_value != str(value):
            field = self.fields_by_placeholder.get(placeholder)
            label = field.description if field else var_name(placeholder)
            self._record_event("field_change", f"{label}: '{old_value}' → '{value}'", source)
        updates = self.calculate_all_updates(trigger_placeholder=placeholder)
        self.values.update(updates)
        return updates

    def set_metadata_value(self, key: str, value: str) -> None:
        self.metadata[key] = value

    def set_assessment_value(self, key: str, value: str) -> None:
        self.assessment[key] = value

    def set_document_status(self, name: str, status: str) -> None:
        self.documents[name] = status

    def set_comment_value(self, key: str, value: str) -> None:
        self.comments[key] = value

    def formatted_live_value(self, field: FieldDefinition, raw_value: str) -> str:
        field_type = field.field_type.lower().strip()
        if field_type.startswith("number"):
            return filter_number_text(raw_value)
        if field_type in {"money"} or field_type.startswith("spell:"):
            return format_money_input_text(raw_value)
        if field_type in {"area"} or field_type.startswith("spell_area:"):
            return format_area_input_text(raw_value)
        if field_type == "percent":
            return filter_percent_text(raw_value)
        if field_type == "date":
            return coerce_date_text(raw_value)
        if is_cccd_desc(field.description):
            return "".join(char for char in str(raw_value or "") if char.isdigit())[:12]
        if is_name_desc(field.description):
            return normalize_square_meter_text(vi_title_name(raw_value))
        return normalize_square_meter_text(raw_value)

    def apply_qr_to_group(self, desc: str, image_path: str) -> dict[str, str]:
        payload = decode_qr_offline(image_path)
        info = parse_cccd_payload(payload)
        group_key = normalize_person_key(desc)
        group = self.person_groups.get(group_key, {})
        updates: dict[str, str] = {}
        if group.get("CCCD"):
            updates[group["CCCD"]] = info.get("CCCD", "")
        if group.get("Ho_va_ten"):
            updates[group["Ho_va_ten"]] = vi_title_name(info.get("Ho_va_ten", ""))
        if group.get("Ngay_thang_nam_sinh"):
            updates[group["Ngay_thang_nam_sinh"]] = info.get("Ngay_thang_nam_sinh", "")
        if group.get("Dia_chi"):
            updates[group["Dia_chi"]] = info.get("Dia_chi", "")
        if group.get("Ngay_cap_CCCD"):
            updates[group["Ngay_cap_CCCD"]] = info.get("Ngay_cap_CCCD", "")
        if group.get("Gioi_tinh"):
            updates[group["Gioi_tinh"]] = info.get("Gioi_tinh", "")
        self.values.update(updates)
        recalculated = self.calculate_all_updates()
        self.values.update(recalculated)
        return updates | recalculated

    def apply_phuongan_selection(self, selected_name: str) -> dict[str, str]:
        updates: dict[str, str] = {}
        selected_name = str(selected_name or "").strip()

        for field in self.fields:
            desc_lower = field.description.lower()
            key_lower = field.key.lower()
            if "tên phương án" in desc_lower or "tên pa" in desc_lower or key_lower in {"tenphuongan", "ten_pa", "ten_phuong_an", "tenpa"}:
                updates[field.placeholder] = selected_name

        if not selected_name or self.phuongan_data is None:
            self.values.update(updates)
            recalculated = self.calculate_all_updates()
            self.values.update(recalculated)
            return updates | recalculated

        code = self.phuongan_map.get(selected_name)
        lowered_columns = {str(col).strip().lower(): str(col) for col in self.phuongan_data.columns}
        name_col = None
        code_col = None
        type_col = None
        item_col = None
        detail_col = None
        ratio_col = None

        for alias in ["tên pa", "tên phương án", "ten pa", "ten phuong an"]:
            if alias in lowered_columns:
                name_col = lowered_columns[alias]
                break
        for alias in ["mã pa", "mã phương án", "ma pa", "ma phuong an"]:
            if alias in lowered_columns:
                code_col = lowered_columns[alias]
                break
        for alias in ["loại", "loai"]:
            if alias in lowered_columns:
                type_col = lowered_columns[alias]
                break
        for alias in ["hạng mục", "hang muc"]:
            if alias in lowered_columns:
                item_col = lowered_columns[alias]
                break
        for alias in ["nội dung chi tiết", "noi dung chi tiet"]:
            if alias in lowered_columns:
                detail_col = lowered_columns[alias]
                break
        for alias in ["tỉ lệ", "tỷ lệ", "ty le"]:
            if alias in lowered_columns:
                ratio_col = lowered_columns[alias]
                break

        filtered = pd.DataFrame()
        if code and code_col:
            normalized_code = str(code).strip().lower()
            filtered = self.phuongan_data[
                self.phuongan_data[code_col].astype(str).str.strip().str.lower() == normalized_code
            ]
        if filtered.empty and name_col:
            normalized_name = selected_name.strip().lower()
            filtered = self.phuongan_data[
                self.phuongan_data[name_col].astype(str).str.strip().str.lower() == normalized_name
            ]

        if filtered.empty or type_col is None:
            self.values.update(updates)
            recalculated = self.calculate_all_updates()
            self.values.update(recalculated)
            return updates | recalculated

        type_series = filtered[type_col].astype(str).str.strip().str.lower()
        expense_rows = filtered[type_series.str.contains("chi|phí|cp", na=False)]
        income_rows = filtered[type_series.str.contains("lợi|nhuận|thu|nhập|ln|tn|loinhuan|thunhap|loi|nhuan", na=False)]

        for index in range(1, 20):
            for key_map in (self.rate_keys["cp"], self.amount_keys["cp"], self.rate_keys["tn"], self.amount_keys["tn"]):
                placeholder = key_map.get(index)
                if placeholder:
                    updates[placeholder] = ""
            for field in self.fields:
                key_lower = field.key.lower()
                if re.match(fr"^(?:hm_cp|hmcp|nd_cp|nd_chiphi|ndcp|chiphi|noidung_cp|hm_tn|hmtn|nd_tn|nd_thunhap|ndtn|thunhap|noidung_tn){index}$", key_lower):
                    updates[field.placeholder] = ""

        for index, (_, row) in enumerate(expense_rows.iterrows(), start=1):
            rate_value = ""
            if ratio_col and pd.notna(row.get(ratio_col)):
                try:
                    rate_value = str(round(float(row[ratio_col]) * 100, 2)).rstrip("0").rstrip(".")
                except Exception:
                    rate_value = str(row[ratio_col])
            hang_muc = str(row.get(item_col, "")) if item_col else ""
            noi_dung = str(row.get(detail_col, "")) if detail_col else ""
            if self.rate_keys["cp"].get(index):
                updates[self.rate_keys["cp"][index]] = rate_value
            self._assign_by_regex(updates, index, hang_muc, noi_dung, kind="cp")

        for index, (_, row) in enumerate(income_rows.iterrows(), start=1):
            rate_value = ""
            if ratio_col and pd.notna(row.get(ratio_col)):
                try:
                    rate_value = str(round(float(row[ratio_col]) * 100, 2)).rstrip("0").rstrip(".")
                except Exception:
                    rate_value = str(row[ratio_col])
            hang_muc = str(row.get(item_col, "")) if item_col else ""
            noi_dung = str(row.get(detail_col, "")) if detail_col else ""
            if self.rate_keys["tn"].get(index):
                updates[self.rate_keys["tn"][index]] = rate_value
            self._assign_by_regex(updates, index, hang_muc, noi_dung, kind="tn")

        self.values.update(updates)
        recalculated = self.calculate_all_updates()
        self.values.update(recalculated)
        self._record_event("apply_phuongan", f"Áp dụng phương án: {selected_name}", "phuongan")
        return updates | recalculated

    def calculate_all_updates(self, trigger_placeholder: str | None = None) -> dict[str, str]:
        updates: dict[str, str] = {}
        updates.update(self._calculate_spell_fields())
        updates.update(self._calculate_financial_fields(trigger_placeholder))
        if trigger_placeholder:
            updates[trigger_placeholder] = self.values.get(trigger_placeholder, "")
        return updates

    def validate(self) -> ValidationResult:
        warnings: list[str] = []
        errors: list[str] = []

        context = self.collect_context()
        applicant = context.get("tenthanhvien", "")
        national_id = context.get("cccd_tv", "")
        requested_amount = parse_number_vi(context.get("sovonxinvay", ""))
        collateral_value = self.collateral_value()
        ltv = self.ltv_ratio()

        if not applicant:
            errors.append("Thiếu tên thành viên / người vay.")
        if not national_id:
            errors.append("Thiếu số CCCD thành viên.")
        elif not validate_cccd_12_digits(national_id):
            errors.append("Số CCCD thành viên phải đủ 12 chữ số.")

        if self._loan_type_placeholder():
            loan_type_value = self.values.get(self._loan_type_placeholder() or "", "")
            if loan_type_value and loan_type_value not in LOAN_TYPE_OPTIONS:
                errors.append("Loại cho vay chưa đúng danh mục cấu hình.")

        for field in self.fields:
            value = self.values.get(field.placeholder, "")
            if var_name(field.placeholder) in REQUIRED_PLACEHOLDER_KEYS and not str(value).strip():
                warnings.append(f"Thiếu dữ liệu bắt buộc: {field.description}.")
            if is_noicap_desc(field.description) and value and value not in NOICAP_OPTIONS:
                errors.append(f"{field.description} chưa đúng danh mục nơi cấp.")
            if is_cccd_desc(field.description) and value and not validate_cccd_12_digits(value):
                errors.append(f"{field.description} phải đủ 12 chữ số.")

        if requested_amount <= 0:
            errors.append("Số vốn xin vay phải lớn hơn 0.")
        if requested_amount > 500_000_000:
            warnings.append("Số tiền vay vượt ngưỡng cảnh báo nội bộ 500 triệu.")
        if requested_amount > 0 and collateral_value <= 0:
            warnings.append("Khoản vay chưa có giá trị tài sản bảo đảm hợp lệ.")
        if ltv > 0.7:
            warnings.append("LTV ratio vượt 70%.")
        if not context.get("GCNQSDĐ") and requested_amount > 0:
            warnings.append("Thiếu thông tin giấy chứng nhận tài sản bảo đảm.")

        return ValidationResult(warnings=sorted(set(warnings)), errors=sorted(set(errors)))

    def summary(self) -> dict[str, str]:
        context = self.collect_context()
        completion = self.completion_progress()
        return {
            "applicant_name": context.get("tenthanhvien", "") or "-",
            "product": context.get("loaichovay", "") or self.metadata.get("product_type", "-"),
            "requested_amount": context.get("sovonxinvay", "") or "-",
            "tenor": context.get("thoigianvay", "") or "-",
            "total_income": "-",
            "monthly_obligations": "-",
            "dti_ratio": "-",
            "collateral_value": format_money_vi(self.collateral_value()) or "-",
            "ltv_ratio": percent_or_blank(self.ltv_ratio()) or "-",
            "completion_text": f"{completion * 100:.0f}%",
            "completion_ratio": f"{completion:.4f}",
            "total_investment": context.get("tvdt", "") or "-",
            "projected_revenue": context.get("tongthu", "") or "-",
            "projected_profit": context.get("loinhuancuapa", "") or "-",
        }

    def collect_context(self) -> dict[str, str]:
        context: dict[str, str] = {}
        for field in self.fields:
            value = self.values.get(field.placeholder, "")
            if is_cccd_desc(field.description):
                value = "".join(char for char in str(value) if char.isdigit())
            if is_name_desc(field.description):
                value = vi_title_name(str(value))
            context[field.key] = value
        return context

    def completion_progress(self) -> float:
        relevant_fields = [field for field in self.fields if not field.readonly and not looks_like_long_text(field.description, field.field_type)]
        if not relevant_fields:
            return 0.0
        filled = sum(1 for field in relevant_fields if str(self.values.get(field.placeholder, "")).strip())
        total_slots = len(relevant_fields)
        return min(1.0, filled / total_slots) if total_slots else 0.0

    def total_monthly_income(self) -> float:
        salary = parse_number_vi(self.assessment.get("monthly_salary", ""))
        other = parse_number_vi(self.assessment.get("other_income", ""))
        return salary + other

    def monthly_obligations(self) -> float:
        return parse_number_vi(self.assessment.get("monthly_debt_payment", ""))

    def living_expenses(self) -> float:
        return parse_number_vi(self.assessment.get("living_expenses", ""))

    def net_disposable_income(self) -> float:
        return self.total_monthly_income() - self.monthly_obligations() - self.living_expenses()

    def dti_ratio(self) -> float:
        return ratio(self.monthly_obligations(), self.total_monthly_income())

    def collateral_value(self) -> float:
        total_placeholder = self.valuation_keys.get("total_valuation")
        if total_placeholder:
            value = parse_number_vi(self.values.get(total_placeholder, ""))
            if value > 0:
                return value
        if self.fields_by_key.get("tgtts"):
            return parse_number_vi(self.values.get(self.fields_by_key["tgtts"].placeholder, ""))
        return 0.0

    def requested_amount(self) -> float:
        if self.loan_keys.get("amount"):
            return parse_number_vi(self.values.get(self.loan_keys["amount"] or "", ""))
        return 0.0

    def ltv_ratio(self) -> float:
        return ratio(self.requested_amount(), self.collateral_value())

    def assessment_snapshot(self) -> dict[str, str]:
        return {
            "total_monthly_income": format_money_vi(self.total_monthly_income()),
            "net_disposable_income": format_money_vi(self.net_disposable_income()),
            "dti_ratio": percent_or_blank(self.dti_ratio()) or "0.0%",
            "ltv_ratio": percent_or_blank(self.ltv_ratio()) or "0.0%",
        }

    def documents_missing(self) -> bool:
        return any(status in {"Pending", "Missing"} for status in self.documents.values())

    def _register_person_field(self, field: FieldDefinition) -> None:
        group_key = normalize_person_key(field.description)
        self.person_groups.setdefault(group_key, {})
        if is_cccd_desc(field.description):
            self.person_groups[group_key]["CCCD"] = field.placeholder
        elif is_name_desc(field.description):
            self.person_groups[group_key]["Ho_va_ten"] = field.placeholder
        elif is_dob_desc(field.description):
            self.person_groups[group_key]["Ngay_thang_nam_sinh"] = field.placeholder
        elif is_addr_desc(field.description):
            self.person_groups[group_key]["Dia_chi"] = field.placeholder
        elif is_issue_desc(field.description):
            self.person_groups[group_key]["Ngay_cap_CCCD"] = field.placeholder
        elif is_noicap_desc(field.description):
            self.person_groups[group_key]["Noi_cap"] = field.placeholder
        elif is_gender_desc(field.description):
            self.person_groups[group_key]["Gioi_tinh"] = field.placeholder

    def _register_financial_field(self, field: FieldDefinition) -> None:
        key = field.key.lower().strip()
        desc_lower = field.description.lower()
        match = re.match(r"^(?:t\.cp|tcp|tlcp)(\d+)$", key)
        if match:
            self.rate_keys["cp"][int(match.group(1))] = field.placeholder
            return
        match = re.match(r"^(?:t\.tn|ttn|tltn)(\d+)$", key)
        if match:
            self.rate_keys["tn"][int(match.group(1))] = field.placeholder
            return
        match = re.match(r"^(?:cp|st_cp)(\d+)$", key)
        if match:
            self.amount_keys["cp"][int(match.group(1))] = field.placeholder
            return
        match = re.match(r"^(?:tn|st_tn)(\d+)$", key)
        if match:
            self.amount_keys["tn"][int(match.group(1))] = field.placeholder
            return

        if key in {"tongthu", "tong_thu", "tongthunhap", "tong_thu_nhap"}:
            self.total_keys["tn"] = field.placeholder
            return
        if key in {"laiphaitra", "lai_phai_tra", "tien_lai"}:
            self.interest_key = field.placeholder
            return
        if key in {"loinhuancuapa", "loinhuan", "loi_nhuan"}:
            self.total_keys["profit"] = field.placeholder
            return
        if key in {"tienvay", "so_tien_vay", "sotienvay", "tongtienvay", "tong_vay", "sovonxinvay"}:
            self.loan_keys["amount"] = field.placeholder
            return
        if key in {"vontuco", "von_tu_co", "tu_co"}:
            self.loan_keys["own_capital"] = field.placeholder
            return
        if key in {"thoihanvay", "thoi_han_vay", "kyhanvay", "thoihan", "kyhan", "thoigianvay"}:
            self.loan_keys["term"] = field.placeholder
            return
        if key in {"laisuatnam", "lai_suat_nam", "laisuat_nam", "lsnam", "laisuatvay"}:
            self.loan_keys["rate_year"] = field.placeholder
            return
        if key in {"laisuatthang", "lai_suat_thang", "laisuat_thang", "lsthang", "laisuatchovaythang", "ls_thang"} or ("lãi suất" in desc_lower and "tháng" in desc_lower):
            self.loan_keys["rate_month"] = field.placeholder
            return
        if key in {"tvdt", "tongvondautu", "tong_von_dau_tu", "tongvon", "tongchi", "tongchiphi", "tong_chi", "tong_chiphi"}:
            self.total_keys["cp"] = field.placeholder
            return

        if self.total_keys["cp"] is None and ("tổng vốn đầu tư" in desc_lower or ("tổng chi" in desc_lower and "chi phí" in desc_lower)):
            self.total_keys["cp"] = field.placeholder
        if "diện tích" in desc_lower and ("đất ở" in desc_lower or "thổ cư" in desc_lower) and "bằng chữ" not in desc_lower:
            self.valuation_keys["area_residential"] = field.placeholder
        if "1m2" in desc_lower and ("đất ở" in desc_lower or "thổ cư" in desc_lower):
            self.valuation_keys["price_residential"] = field.placeholder
        if "tổng" in desc_lower and "định giá" in desc_lower and ("đất ở" in desc_lower or "thổ cư" in desc_lower) and "bằng chữ" not in desc_lower:
            self.valuation_keys["total_residential"] = field.placeholder
        if "diện tích" in desc_lower and "03" in desc_lower and "bằng chữ" not in desc_lower:
            self.valuation_keys["area_03"] = field.placeholder
        if "1m2" in desc_lower and "03" in desc_lower:
            self.valuation_keys["price_03"] = field.placeholder
        if "tổng" in desc_lower and "03" in desc_lower and "bằng chữ" not in desc_lower:
            self.valuation_keys["total_03"] = field.placeholder
        if "tổng giá trị tài sản thế chấp" in desc_lower and "bằng chữ" not in desc_lower:
            self.valuation_keys["total_valuation"] = field.placeholder
        if ("tổng số tiền các tài sản khác" in desc_lower or "tài sản khác" in desc_lower) and "bằng chữ" not in desc_lower:
            self.valuation_keys["total_others"] = field.placeholder

    def _load_phuongan_if_available(self) -> None:
        path = Path(self.phuongan_path)
        self.phuongan_map = {}
        self.phuongan_data = None
        if not path.exists():
            return

        def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
            copy = df.copy()
            copy.columns = [str(col).strip() for col in copy.columns]
            return copy

        def pick_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
            lowered = {str(col).strip().lower(): str(col) for col in df.columns}
            for alias in aliases:
                key = alias.strip().lower()
                if key in lowered:
                    return lowered[key]
            return None

        def looks_like_data_sheet(df: pd.DataFrame) -> bool:
            cols = {str(col).strip().lower() for col in df.columns}
            keywords = {"mã phương án", "loại", "hạng mục", "nội dung chi tiết", "tỉ lệ", "tỷ lệ"}
            return len(cols.intersection(keywords)) >= 3

        def looks_like_map_sheet(df: pd.DataFrame) -> bool:
            cols = {str(col).strip().lower() for col in df.columns}
            name_aliases = {"tên pa", "tên phương án", "ten pa", "ten phuong an"}
            code_aliases = {"mã pa", "mã phương án", "ma pa", "ma phuong an"}
            return bool(cols.intersection(name_aliases)) and bool(cols.intersection(code_aliases))

        try:
            sheets = pd.read_excel(path, sheet_name=None)
            normalized_sheets = {str(name).strip(): normalize_columns(frame) for name, frame in sheets.items()}

            map_df = None
            data_df = None

            for sheet_name, df in normalized_sheets.items():
                key = sheet_name.strip().lower().replace(" ", "")
                if key in {"phuongan", "phươngán", "phuong_an", "danhmucphuongan"} and map_df is None:
                    map_df = df
                if key in {"data", "dữliệu", "dulieu", "du_lieu", "chitietphuongan"} and data_df is None:
                    data_df = df

            for df in normalized_sheets.values():
                if map_df is None and looks_like_map_sheet(df):
                    map_df = df
                if data_df is None and looks_like_data_sheet(df):
                    data_df = df

            if map_df is None and normalized_sheets:
                map_df = next(iter(normalized_sheets.values()))
            if data_df is None:
                for df in normalized_sheets.values():
                    if df is not map_df:
                        data_df = df
                        break
            if data_df is None:
                data_df = map_df

            if map_df is not None:
                name_col = pick_column(map_df, ["Tên PA", "Tên phương án", "Ten PA", "Ten phuong an"])
                code_col = pick_column(map_df, ["Mã PA", "Mã phương án", "Ma PA", "Ma phuong an"])
                if name_col and code_col:
                    self.phuongan_map = {
                        str(name).strip(): str(code).strip()
                        for name, code in zip(map_df[name_col], map_df[code_col])
                        if pd.notna(name) and str(name).strip() and pd.notna(code) and str(code).strip()
                    }

            if data_df is not None:
                self.phuongan_data = data_df.copy()
                self.phuongan_data.columns = [str(col).strip() for col in self.phuongan_data.columns]

                if not self.phuongan_map:
                    name_col = pick_column(self.phuongan_data, ["Tên PA", "Tên phương án", "Ten PA", "Ten phuong an"])
                    code_col = pick_column(self.phuongan_data, ["Mã PA", "Mã phương án", "Ma PA", "Ma phuong an"])
                    if name_col and code_col:
                        rows = (
                            self.phuongan_data[[name_col, code_col]]
                            .dropna()
                            .drop_duplicates()
                            .values
                            .tolist()
                        )
                        self.phuongan_map = {
                            str(name).strip(): str(code).strip()
                            for name, code in rows
                            if str(name).strip() and str(code).strip()
                        }

        except Exception:
            self.phuongan_map = {}
            self.phuongan_data = None

    def available_phuongan_names(self) -> list[str]:
        names: list[str] = []
        if self.phuongan_map:
            names.extend([name for name in self.phuongan_map.keys() if str(name).strip()])
        if self.phuongan_data is not None:
            lowered = {str(col).strip().lower(): str(col) for col in self.phuongan_data.columns}
            name_col = None
            for alias in ["tên pa", "tên phương án", "ten pa", "ten phuong an"]:
                if alias in lowered:
                    name_col = lowered[alias]
                    break
            if name_col:
                names.extend(
                    [
                        str(value).strip()
                        for value in self.phuongan_data[name_col].dropna().tolist()
                        if str(value).strip()
                    ]
                )
        seen = set()
        ordered = []
        for name in names:
            key = name.lower()
            if key not in seen:
                seen.add(key)
                ordered.append(name)
        return ordered

    def _assign_by_regex(self, updates: dict[str, str], index: int, hang_muc: str, noi_dung: str, kind: str) -> None:
        for field in self.fields:
            key = field.key.lower()
            if kind == "cp":
                if re.match(fr"^(?:hm_cp|hmcp){index}$", key):
                    updates[field.placeholder] = hang_muc
                elif re.match(fr"^(?:nd_cp|nd_chiphi|ndcp){index}$", key):
                    updates[field.placeholder] = noi_dung
                elif re.match(fr"^(?:chiphi|noidung_cp){index}$", key):
                    updates[field.placeholder] = f"{hang_muc}: {noi_dung}".strip(": ")
            else:
                if re.match(fr"^(?:hm_tn|hmtn){index}$", key):
                    updates[field.placeholder] = hang_muc
                elif re.match(fr"^(?:nd_tn|nd_thunhap|ndtn){index}$", key):
                    updates[field.placeholder] = noi_dung
                elif re.match(fr"^(?:thunhap|noidung_tn){index}$", key):
                    updates[field.placeholder] = f"{hang_muc}: {noi_dung}".strip(": ")

    def _calculate_spell_fields(self) -> dict[str, str]:
        updates: dict[str, str] = {}
        for field in self.fields:
            value = self.values.get(field.placeholder, "")
            lower_type = field.field_type.lower().strip()
            if lower_type.startswith("spell:"):
                target_name = var_name(re.split(r"^spell\s*:", field.field_type, flags=re.IGNORECASE)[1].strip())
                target_field = self.fields_by_placeholder.get(target_name) or self.fields_by_key.get(target_name)
                if target_field:
                    amount = parse_number_vi(value)
                    updates[target_field.placeholder] = self.number_to_vn_money(amount) if amount else ""
            elif lower_type.startswith("spell_area:"):
                target_name = var_name(re.split(r"^spell_area\s*:", field.field_type, flags=re.IGNORECASE)[1].strip())
                target_field = self.fields_by_placeholder.get(target_name) or self.fields_by_key.get(target_name)
                if target_field:
                    updates[target_field.placeholder] = self.number_to_vn_decimal(str(value)) if str(value).strip() else ""
        return updates

    def _calculate_financial_fields(self, trigger_placeholder: str | None = None) -> dict[str, str]:
        updates: dict[str, str] = {}
        loan_amount = parse_number_vi(self.values.get(self.loan_keys.get("amount") or "", ""))
        own_capital = parse_number_vi(self.values.get(self.loan_keys.get("own_capital") or "", ""))
        total_cost = parse_number_vi(self.values.get(self.total_keys.get("cp") or "", ""))

        if self.total_keys.get("cp") and (loan_amount > 0 or own_capital > 0):
            total_cost = loan_amount + own_capital
            updates[self.total_keys["cp"]] = format_money_vi(total_cost)

        expense_indexes = sorted(self.rate_keys["cp"].keys())
        if expense_indexes:
            interest_index = 8 if 8 in expense_indexes else expense_indexes[-1]
            residual_index = 7 if 7 in expense_indexes else (expense_indexes[-2] if len(expense_indexes) >= 2 else None)
            user_indexes = [index for index in expense_indexes if index not in {interest_index, residual_index}]
            annual_rate = parse_percent_vi(self.values.get(self.loan_keys.get("rate_year") or "", ""))
            if self.loan_keys.get("rate_month"):
                updates[self.loan_keys["rate_month"]] = format_decimal_vi(annual_rate / 12.0, 4) if annual_rate else ""
            interest_amount = loan_amount * annual_rate / 100.0 if loan_amount > 0 and annual_rate > 0 else 0.0
            interest_rate = ratio(interest_amount, total_cost) * 100.0 if total_cost > 0 else 0.0
            if self.interest_key:
                updates[self.interest_key] = format_money_vi(interest_amount)
            if self.amount_keys["cp"].get(interest_index):
                updates[self.amount_keys["cp"][interest_index]] = format_money_vi(interest_amount)
            if self.rate_keys["cp"].get(interest_index):
                updates[self.rate_keys["cp"][interest_index]] = format_decimal_vi(interest_rate, 3)

            used_user_rate = 0.0
            user_amount_total = 0.0
            for index in user_indexes:
                rate_placeholder = self.rate_keys["cp"][index]
                rate = parse_rate_vi(self.values.get(rate_placeholder, ""))
                used_user_rate += rate
                amount = total_cost * rate / 100.0
                user_amount_total += amount
                if self.amount_keys["cp"].get(index):
                    updates[self.amount_keys["cp"][index]] = format_money_vi(amount)

            if residual_index and self.rate_keys["cp"].get(residual_index):
                residual_rate = max(0.0, 100.0 - used_user_rate - interest_rate)
                updates[self.rate_keys["cp"][residual_index]] = format_decimal_vi(residual_rate, 3)
                residual_amount = max(0.0, total_cost - interest_amount - user_amount_total)
                if self.amount_keys["cp"].get(residual_index):
                    updates[self.amount_keys["cp"][residual_index]] = format_money_vi(residual_amount)

        income_indexes = sorted(self.rate_keys["tn"].keys())
        total_income_placeholder = self.total_keys.get("tn")
        profit_placeholder = self.total_keys.get("profit")
        total_income = parse_number_vi(self.values.get(total_income_placeholder or "", ""))

        if trigger_placeholder and profit_placeholder and trigger_placeholder == profit_placeholder:
            entered_profit = parse_number_vi(self.values.get(profit_placeholder or "", ""))
            total_income = max(0.0, total_cost + entered_profit)
            if total_income_placeholder:
                updates[total_income_placeholder] = format_money_vi(total_income)

        if income_indexes:
            last_index = max(income_indexes)
            used_income_rate = sum(parse_rate_vi(self.values.get(self.rate_keys["tn"][idx], "")) for idx in income_indexes if idx != last_index)
            if self.rate_keys["tn"].get(last_index):
                updates[self.rate_keys["tn"][last_index]] = format_decimal_vi(max(0.0, 100.0 - used_income_rate), 2)
            total_income = parse_number_vi(updates.get(total_income_placeholder or "", self.values.get(total_income_placeholder or "", "")))
            if total_income > 0:
                for index in income_indexes:
                    rate = parse_rate_vi(updates.get(self.rate_keys["tn"][index], self.values.get(self.rate_keys["tn"][index], "")))
                    if self.amount_keys["tn"].get(index):
                        updates[self.amount_keys["tn"][index]] = format_money_vi(total_income * rate / 100.0)

        if profit_placeholder:
            total_income = parse_number_vi(updates.get(total_income_placeholder or "", self.values.get(total_income_placeholder or "", "")))
            updates[profit_placeholder] = format_money_vi(total_income - total_cost)

        area_res = parse_decimal_vi(self.values.get(self.valuation_keys.get("area_residential") or "", ""))
        price_res = parse_number_vi(self.values.get(self.valuation_keys.get("price_residential") or "", ""))
        if self.valuation_keys.get("total_residential") and area_res > 0 and price_res > 0:
            updates[self.valuation_keys["total_residential"]] = format_money_vi(area_res * price_res)

        area_03 = parse_decimal_vi(self.values.get(self.valuation_keys.get("area_03") or "", ""))
        price_03 = parse_number_vi(self.values.get(self.valuation_keys.get("price_03") or "", ""))
        if self.valuation_keys.get("total_03") and area_03 > 0 and price_03 > 0:
            updates[self.valuation_keys["total_03"]] = format_money_vi(area_03 * price_03)

        total_collateral = parse_number_vi(self.values.get(self.valuation_keys.get("total_valuation") or "", ""))
        residential_total = parse_number_vi(updates.get(self.valuation_keys.get("total_residential") or "", self.values.get(self.valuation_keys.get("total_residential") or "", "")))
        zero3_total = parse_number_vi(updates.get(self.valuation_keys.get("total_03") or "", self.values.get(self.valuation_keys.get("total_03") or "", "")))
        if self.valuation_keys.get("total_others") and total_collateral > 0:
            updates[self.valuation_keys["total_others"]] = format_money_vi(max(0.0, total_collateral - residential_total - zero3_total))

        return updates

    def _loan_type_placeholder(self) -> str | None:
        for field in self.fields:
            if is_loan_type_desc(field.description):
                return field.placeholder
        return None

    @staticmethod
    def number_to_vn_money(number: float) -> str:
        try:
            from num2words import num2words
        except Exception:
            return ""
        text = num2words(int(number), lang="vi").replace("-", " ")
        return text[:1].upper() + text[1:] + " đồng chẵn."

    @staticmethod
    def number_to_vn_decimal(value: str) -> str:
        try:
            from num2words import num2words
        except Exception:
            return ""
        clean = format_area_input_text(value)
        if not clean:
            return ""
        if "," not in clean:
            text = num2words(int(clean.replace(".", "") or 0), lang="vi").replace("-", " ")
            return text[:1].upper() + text[1:] + " mét vuông."
        integer_part, decimal_part = clean.split(",", 1)
        integer_text = num2words(int(integer_part.replace(".", "") or 0), lang="vi").replace("-", " ")
        if not decimal_part:
            return integer_text[:1].upper() + integer_text[1:] + " mét vuông."
        leading_zeros = ""
        for char in decimal_part:
            if char == "0":
                leading_zeros += "không "
            else:
                break
        decimal_value = int(decimal_part or "0")
        decimal_text = leading_zeros.strip() if decimal_value == 0 else (leading_zeros + num2words(decimal_value, lang="vi").replace("-", " ")).strip()
        result = f"{integer_text} phẩy {decimal_text}".replace("  ", " ").strip()
        return result[:1].upper() + result[1:] + " mét vuông."

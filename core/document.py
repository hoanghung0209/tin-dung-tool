from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import pandas as pd

from core.utils import extract_spell_target, looks_like_long_text, normalize_type, split_dropdown_options, var_name

try:
    from docxtpl import DocxTemplate
except Exception:  # pragma: no cover
    DocxTemplate = None


@dataclass(slots=True)
class FieldDefinition:
    placeholder: str
    key: str
    description: str
    field_type: str
    spell: str
    tab: str
    options: list[str]
    long_text: bool
    readonly: bool = False
    section: str = ""
    section_row: int | None = None
    section_col: int | None = None
    section_span: int = 1
    field_row: int | None = None
    field_col: int | None = None
    field_span: int = 1
    label_width: int | None = None
    input_width: int | None = None
    visible: bool = True
    order: int = 0


@dataclass(slots=True)
class MappingBundle:
    mapping: dict[str, str]
    field_types: dict[str, str]
    tabs: dict[str, str]
    fields: list[FieldDefinition]
    tab_sequence: list[str]
    uses_layout_metadata: bool


DEFAULT_TAB_NAME = "Thông tin chung"


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(column).strip() for column in df.columns]
    return df


def _pick_fields_sheet(xlsx_path: str) -> pd.DataFrame:
    sheets = pd.read_excel(xlsx_path, sheet_name=None)
    if "Fields" in sheets:
        return _normalize_columns(sheets["Fields"]).dropna(how="all")
    first_name = next(iter(sheets))
    return _normalize_columns(sheets[first_name]).dropna(how="all")


def _find_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    alias_set = {alias.strip().lower() for alias in aliases}
    for column in df.columns:
        if str(column).strip().lower() in alias_set:
            return str(column)
    return None


def _row_text(row: dict, col: str | None, default: str = "") -> str:
    if not col:
        return default
    value = row.get(col, default)
    if pd.isna(value):
        return default
    return str(value).strip()


def _row_int(row: dict, col: str | None, default: int | None = None) -> int | None:
    if not col:
        return default
    value = row.get(col, None)
    if value is None or pd.isna(value) or str(value).strip() == "":
        return default
    try:
        return int(float(str(value).strip()))
    except Exception:
        return default


def _row_bool(row: dict, col: str | None, default: bool = True) -> bool:
    if not col:
        return default
    value = row.get(col, None)
    if value is None or pd.isna(value):
        return default
    text = str(value).strip().lower()
    if text in {"0", "false", "no", "n", "ẩn", "hide"}:
        return False
    if text in {"1", "true", "yes", "y", "hiện", "show"}:
        return True
    return default


def read_mapping(xlsx_path: str) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    bundle = load_mapping_bundle(xlsx_path)
    return bundle.mapping, bundle.field_types, bundle.tabs


def load_mapping_bundle(xlsx_path: str) -> MappingBundle:
    df = _pick_fields_sheet(xlsx_path)

    tab_col = _find_col(df, ["tab", "nhóm"])
    spell_col = _find_col(df, ["spell"])
    section_col_name = _find_col(df, ["section", "group", "nhóm hiển thị", "khối"])
    section_row_col = _find_col(df, ["section_row", "section row"])
    section_col_col = _find_col(df, ["section_col", "section col"])
    section_span_col = _find_col(df, ["section_span", "section span"])
    field_row_col = _find_col(df, ["field_row", "field row"])
    field_col_col = _find_col(df, ["field_col", "field col"])
    field_span_col = _find_col(df, ["field_span", "field span"])
    label_width_col = _find_col(df, ["label_width", "label width"])
    input_width_col = _find_col(df, ["input_width", "input width"])
    visible_col = _find_col(df, ["visible", "show", "hiển thị"])
    order_col = _find_col(df, ["order", "thứ tự"])

    mapping: dict[str, str] = {}
    field_types: dict[str, str] = {}
    tabs: dict[str, str] = {}
    raw_fields: list[FieldDefinition] = []
    spell_targets: set[str] = set()
    tab_sequence: list[str] = []
    uses_layout_metadata = any(
        column is not None
        for column in [
            section_col_name,
            section_row_col,
            section_col_col,
            field_row_col,
            field_col_col,
            label_width_col,
            input_width_col,
        ]
    )

    for row_index, row in enumerate(df.to_dict(orient="records")):
        placeholder = _row_text(row, "placeholder")
        if not placeholder:
            continue
        description = _row_text(row, "description")
        raw_type = _row_text(row, "type")
        raw_spell = _row_text(row, spell_col)
        field_type = normalize_type(raw_type, raw_spell)
        tab_name = _row_text(row, tab_col, DEFAULT_TAB_NAME) or DEFAULT_TAB_NAME
        section_name = _row_text(row, section_col_name, tab_name) or tab_name

        mapping[placeholder] = description
        field_types[placeholder] = field_type
        tabs[placeholder] = tab_name
        if tab_name not in tab_sequence:
            tab_sequence.append(tab_name)

        target = extract_spell_target(field_type)
        if target:
            spell_targets.add(target)

        raw_fields.append(
            FieldDefinition(
                placeholder=placeholder,
                key=var_name(placeholder),
                description=description,
                field_type=field_type,
                spell=raw_spell,
                tab=tab_name,
                options=split_dropdown_options(field_type),
                long_text=looks_like_long_text(description, field_type),
                readonly=False,
                section=section_name,
                section_row=_row_int(row, section_row_col),
                section_col=_row_int(row, section_col_col),
                section_span=max(1, _row_int(row, section_span_col, 1) or 1),
                field_row=_row_int(row, field_row_col),
                field_col=_row_int(row, field_col_col),
                field_span=max(1, _row_int(row, field_span_col, 1) or 1),
                label_width=_row_int(row, label_width_col),
                input_width=_row_int(row, input_width_col),
                visible=_row_bool(row, visible_col, True),
                order=_row_int(row, order_col, row_index) or row_index,
            )
        )

    fields: list[FieldDefinition] = []
    for field in raw_fields:
        readonly = field.key in spell_targets or field.description.lower().endswith("bằng chữ")
        fields.append(
            FieldDefinition(
                placeholder=field.placeholder,
                key=field.key,
                description=field.description,
                field_type=field.field_type,
                spell=field.spell,
                tab=field.tab,
                options=field.options,
                long_text=field.long_text,
                readonly=readonly,
                section=field.section,
                section_row=field.section_row,
                section_col=field.section_col,
                section_span=field.section_span,
                field_row=field.field_row,
                field_col=field.field_col,
                field_span=field.field_span,
                label_width=field.label_width,
                input_width=field.input_width,
                visible=field.visible,
                order=field.order,
            )
        )

    return MappingBundle(
        mapping=mapping,
        field_types=field_types,
        tabs=tabs,
        fields=fields,
        tab_sequence=tab_sequence,
        uses_layout_metadata=uses_layout_metadata,
    )


def generate_word_document(template_path: str, output_path: str, context: dict) -> None:
    if DocxTemplate is None:
        raise RuntimeError("Thiếu thư viện docxtpl. Hãy cài đặt trước khi xuất Word.")
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)


def group_fields_by_tab(fields: Iterable[FieldDefinition]) -> dict[str, list[FieldDefinition]]:
    grouped: dict[str, list[FieldDefinition]] = {}
    for field in fields:
        grouped.setdefault(field.tab, []).append(field)
    return grouped

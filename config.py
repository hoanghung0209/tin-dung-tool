from __future__ import annotations

import json
from pathlib import Path

APP_TITLE = "Màn hình nhập hồ sơ vay vốn"
APP_SUBTITLE = "Giao diện tác nghiệp nội bộ / hồ sơ vay vốn"

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_XLSX = str(BASE_DIR / "Data(17).xlsx")
DEFAULT_DOCX = str(BASE_DIR / "HSCV.docx")
DEFAULT_PHUONGAN = str(BASE_DIR / "phuongan.xlsx")
DEFAULT_OUTPUT_DIR = str(Path.home() / "Desktop" / "HoSoTinDung")
DRAFT_FILE = str(BASE_DIR / "draft.json")
DOSSIER_DIR = str(BASE_DIR / "records")
SETTINGS_FILE = str(BASE_DIR / "config.json")

NOICAP_OPTIONS = [
    "Bộ Công An",
    "CCSQLHCVTTXH",
]

LOAN_TYPE_OPTIONS = [
    "Ngắn hạn thỏa thuận",
    "Trung hạn thỏa thuận",
]

LONG_TEXT_HINTS = ["ghi chú"]

SECTION_TITLES = {
    "Thông tin người vay": "Thông tin người vay",
    "Thông tin khoản vay": "Thông tin khoản vay",
    "Thông tin về phương án vay vốn": "Thông tin về phương án vay vốn",
    "Thông tin về tài sản thế chấp": "Thông tin về tài sản thế chấp",
    "Thông tin về người thụ hưởng": "Thông tin về người thụ hưởng",
    "Thông tin khác": "Thông tin khác",
}

DISPLAY_SECTION_MAP = {
    "Thông tin người vay": "Thông tin người vay",
    "Thông tin khoản vay": "Thông tin khoản vay",
    "Thông tin phương án vay vốn": "Thông tin về phương án vay vốn",
    "Thông tin về tài sản thế chấp": "Thông tin về tài sản thế chấp",
    "Thông tin về người thụ hưởng": "Thông tin về người thụ hưởng",
    "Thông tin khác": "Thông tin khác",
}

SIDEBAR_ORDER = [
    "Thông tin người vay",
    "Thông tin khoản vay",
    "Thông tin về phương án vay vốn",
    "Thông tin về tài sản thế chấp",
    "Thông tin về người thụ hưởng",
    "Thông tin khác",
]

FIELD_ORDER_BY_SECTION = {
    "Thông tin người vay": [
        "tenthanhvien", "thethanhvien", "ntnstv", "cccd_tv", "cccd_tv_nc", "cccd_tv_noi", "nn_tv", "diachithanhvien",
        "ten_dvv1", "qhvtv_1", "ntns_dvv1", "cccd_dvv1", "cccd_dvv1_nc", "cccd_dvv1_noi", "nn_dvv1", "diachi_dvv1",
    ],
    "Thông tin khoản vay": [
        "loaichovay", "thoigianvay", "laisuatvay", "ls_thang", "sovonxinvay", "svxv_bc", "vontuco", "tvdt",
    ],
    "Thông tin về phương án vay vốn": [
        "tenphuongan", "nguontrano",
        "hm_cp1", "nd_chiphi1", "tcp1", "cp1",
        "hm_cp2", "nd_chiphi2", "tcp2", "cp2",
        "hm_cp3", "nd_chiphi3", "tcp3", "cp3",
        "hm_cp4", "nd_chiphi4", "tcp4", "cp4",
        "hm_cp5", "nd_chiphi5", "tcp5", "cp5",
        "hm_cp6", "nd_chiphi6", "tcp6", "cp6",
        "tcp7", "cp7", "tcp8", "cp8",
        "hm_tn1", "nd_thunhap1", "ttn1", "tn1",
        "hm_tn2", "nd_thunhap2", "ttn2", "tn2",
        "hm_tn3", "nd_thunhap3", "ttn3", "tn3",
        "tongthu", "laiphaitra", "loinhuancuapa",
    ],
    "Thông tin về tài sản thế chấp": [
        "GCNQSDĐ", "ngc_GCNQSDĐ", "nc_GCNQSDĐ", "sovaoso_GCN",
        "dientichdato", "dt_do_bc", "stbd_do", "tbd_do", "vitri_do",
        "dt_nhao", "sgst", "kc_nhao", "dt_ctp", "kc_ctp",
        "sl_so03", "ssruong03", "nc_sr03", "dt_sr03", "tbd_03", "stbd_03", "vitri_03", "bbck",
        "tstc_1", "tstc_2", "tstc_3", "tstc_4",
        "tgtts", "tgtts_bc",
        "gt_do", "ts_gt_do", "gt_d03", "ts_gt_d03", "ts_khac",
        "dong_do", "cd_dong", "tay_do", "cd_tay", "nam_do", "cd_nam", "bac_do", "cd_bac",
    ],
    "Thông tin về người thụ hưởng": [
        "nth", "cccd_nth", "ngc_cccd_nth", "nc_cccd_nth", "dc_nth",
    ],
    "Thông tin khác": [
        "datcoc", "datcoc_bc", "st_conlai", "st_conlai_bc",
    ],
}

REQUIRED_PLACEHOLDER_KEYS = {
    "tenthanhvien",
    "cccd_tv",
    "loaichovay",
    "thoigianvay",
    "sovonxinvay",
    "vontuco",
    "tvdt",
    "tenphuongan",
    "nguontrano",
    "tgtts",
}

MANDATORY_DOCUMENTS = [
    "Đơn đề nghị vay vốn",
    "CCCD thành viên",
    "Hồ sơ phương án vay vốn",
    "Hồ sơ tài sản bảo đảm",
    "Biên bản thẩm định nội bộ",
]

DEFAULT_METADATA = {
    "application_id": "APP-2026-00125",
    "status": "Nháp",
    "branch": "PGD Phùng Hưng",
    "officer": "Nguyễn Minh Hải",
    "product_type": "Hồ sơ vay vốn",
    "application_date": "21/04/2026",
    "priority": "Bình thường",
    "channel": "Nhập tại quầy",
    "created_by": "credit.officer01",
}

DEFAULT_ASSESSMENT = {
    "monthly_salary": "18.000.000",
    "other_income": "7.500.000",
    "monthly_debt_payment": "9.000.000",
    "living_expenses": "8.500.000",
}

DEFAULT_DOCUMENT_STATUS = {name: "Đủ" for name in MANDATORY_DOCUMENTS}

DEFAULT_COMMENTS = {
    "internal_note": "",
    "risk_flags": "",
    "special_conditions": "",
}

SAMPLE_CONTEXT = {
    "tenthanhvien": "Nguyễn Văn An",
    "ntnstv": "01/02/1987",
    "cccd_tv": "001234567890",
    "cccd_tv_nc": "15/08/2021",
    "cccd_tv_noi": "Bộ Công An",
    "thethanhvien": "120045",
    "diachithanhvien": "Xóm Trung Tâm, xã Phùng Hưng, huyện Khoái Châu, Hưng Yên",
    "nn_tv": "Làm ruộng",
    "ten_dvv1": "Nguyễn Thị Hạnh",
    "qhvtv_1": "Vợ",
    "ntns_dvv1": "04/10/1989",
    "cccd_dvv1": "001234567891",
    "cccd_dvv1_nc": "20/08/2021",
    "cccd_dvv1_noi": "Bộ Công An",
    "diachi_dvv1": "Xóm Trung Tâm, xã Phùng Hưng, huyện Khoái Châu, Hưng Yên",
    "nn_dvv1": "Lao động tự do",
    "loaichovay": "Ngắn hạn thỏa thuận",
    "thoigianvay": "12 tháng",
    "laisuatvay": "10,8",
    "sovonxinvay": "350.000.000",
    "vontuco": "120.000.000",
    "tvdt": "470.000.000",
    "tenphuongan": "Phương án đầu tư trồng cây ăn quả",
    "nguontrano": "Nguồn thu từ sản xuất nông nghiệp và thu nhập hộ gia đình.",
    "tongthu": "620.000.000",
    "laiphaitra": "37.800.000",
    "loinhuancuapa": "150.000.000",
    "GCNQSDĐ": "CH 005612",
    "ngc_GCNQSDĐ": "05/05/2018",
    "nc_GCNQSDĐ": "UBND huyện Khoái Châu",
    "sovaoso_GCN": "CS 00125",
    "dientichdato": "120,5",
    "dt_do_bc": "Một trăm hai mươi phẩy năm mét vuông.",
    "stbd_do": "152",
    "tbd_do": "18",
    "vitri_do": "Đất ở nông thôn, mặt đường liên xã",
    "sl_so03": "1",
    "ssruong03": "03-0098",
    "nc_sr03": "02/03/2019",
    "dt_sr03": "860",
    "vitri_03": "Khu đồng Mả Tre",
    "tstc_1": "Quyền sử dụng đất và tài sản gắn liền với đất",
    "tgtts": "850.000.000",
    "tgtts_bc": "Tám trăm năm mươi triệu đồng chẵn.",
    "gt_do": "5.500.000",
    "ts_gt_do": "662.750.000",
    "gt_d03": "180.000",
    "ts_gt_d03": "154.800.000",
    "ts_khac": "32.450.000",
    "nth": "Nguyễn Văn An",
    "cccd_nth": "001234567890",
    "nc_cccd_nth": "Bộ Công An",
    "ngc_cccd_nth": "15/08/2021",
    "dc_nth": "Xóm Trung Tâm, xã Phùng Hưng, huyện Khoái Châu, Hưng Yên",
    "datcoc": "50.000.000",
    "st_conlai": "300.000.000",
}

APP_STYLESHEET = """
QMainWindow, QWidget {
    background: #f4f6f8;
    color: #111827;
    font-family: 'Segoe UI';
    font-size: 12px;
}
QFrame#HeaderFrame, QFrame#FooterFrame, QFrame#SidebarFrame, QFrame#SummaryFrame, QFrame#CardFrame, QTabWidget::pane {
    background: #ffffff;
    border: 1px solid #d8dee6;
    border-radius: 0px;
}
QFrame#HeaderFrame {
    border-top-left-radius: 0px;
    border-top-right-radius: 0px;
    border-bottom-left-radius: 10px;
    border-bottom-right-radius: 10px;
}
QFrame#FooterFrame {
    border-bottom-left-radius: 0px;
    border-bottom-right-radius: 0px;
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
}
QFrame#SidebarFrame, QFrame#SummaryFrame {
    background: #f8fafc;
}
QFrame#SummaryFrame QFrame#CardFrame {
    background: #ffffff;
    border: 1px solid #dde3ea;
    border-radius: 0px;
}
QPushButton {
    background: #f3f4f6;
    color: #1f2937;
    border: 1px solid #c9d0da;
    border-radius: 0px;
    padding: 4px 10px;
    min-height: 28px;
    font-weight: 600;
}
QPushButton:hover {
    background: #eef3f8;
    border-color: #c4cfdd;
}
QPushButton[role="primary"] {
    background: #1f5fbf;
    color: white;
    border: 1px solid #1f5fbf;
}
QPushButton[role="primary"]:hover { background: #184e9c; }
QPushButton[role="warning"] {
    background: #fff7ea;
    color: #7c4a00;
    border: 1px solid #ead6ab;
}
QPushButton[role="icon"] {
    min-width: 24px; max-width: 24px;
    min-height: 24px; max-height: 24px;
    padding: 0px; font-size: 11px; border-radius: 0px;
}
QLineEdit, QTextEdit, QComboBox, QDateEdit {
    background: #ffffff;
    border: 1px solid #c5ccd6;
    border-radius: 0px;
    padding: 3px 8px;
    min-height: 26px;
    selection-background-color: #dbeafe;
}
QTextEdit { padding-top: 6px; }
QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDateEdit:focus {
    border: 1px solid #6c86b0;
    background: #f7f9fc;
}
QLineEdit:read-only, QTextEdit:read-only {
    background: #f6f8fa;
    color: #475569;
}
QLabel[role="page_title"] {
    font-size: 18px;
    font-weight: 700;
    color: #111827;
}
QLabel[role="page_subtitle"], QLabel[role="subtle"] { color: #667085; }
QLabel[role="section_title"] {
    font-size: 13px;
    font-weight: 700;
    color: #10233b;
}
QLabel[role="group_title"] {
    font-size: 15px;
    font-weight: 800;
    color: #111827;
}
QLabel[role="hero_title"] {
    font-size: 17px;
    font-weight: 800;
    color: #111827;
    background: #ffffff;
    border: 1px solid #7bb661;
    border-radius: 0px;
    padding: 6px 18px;
}
QLabel[role="value"] {
    font-size: 12px;
    font-weight: 600;
    color: #111827;
}
QLabel[role="metric"] {
    font-size: 15px;
    font-weight: 700;
    color: #111827;
}
QLabel[role="footer_status"] {
    color: #667085;
    font-size: 11px;
    padding-right: 4px;
}
QLabel[role="borrower_label"] {
    color: #111827;
    font-size: 12px;
    font-weight: 500;
}
QLineEdit[role="borrower_input"], QComboBox[role="borrower_input"] {
    background: #bccbe4;
    border: 1px solid #aabbd7;
    border-radius: 0px;
    min-height: 24px;
    max-height: 24px;
    padding: 2px 8px;
}
QLineEdit[role="borrower_input"]:focus, QComboBox[role="borrower_input"]:focus {
    background: #c8d5ea;
    border: 1px solid #7f97bf;
}
QLabel[role="table_header"] {
    font-size: 12px;
    font-weight: 700;
    color: #111827;
    padding-left: 2px;
}
QLabel[role="row_index"] {
    border: 1px solid #87bf67;
    border-radius: 6px;
    background: #ffffff;
    padding: 2px 0px;
    min-width: 24px;
    max-width: 24px;
    min-height: 20px;
    font-weight: 600;
    color: #4b5563;
}
QListWidget { background: transparent; border: none; outline: none; }
QListWidget::item {
    padding: 8px 10px;
    border-radius: 8px;
    margin-bottom: 1px;
}
QListWidget::item:selected {
    background: #e7f1ff;
    color: #16406a;
    border: 1px solid #c7daf9;
}
QProgressBar {
    background: #eef2f6;
    border: 1px solid #dbe2ea;
    border-radius: 8px;
    text-align: center;
    min-height: 16px;
}
QProgressBar::chunk { background: #1f5fbf; border-radius: 7px; }
QTabWidget::pane { top: -1px; padding: 4px; }
QTabBar::tab {
    background: #eef1f5;
    border: 1px solid #d6dde6;
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
    padding: 7px 14px;
    margin-right: 2px;
    min-width: 112px;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #163d67;
    border-bottom-color: #ffffff;
}
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 3px; }
QScrollBar::handle:vertical { background: #cad5e3; min-height: 24px; border-radius: 5px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
"""


def load_config() -> dict:
    try:
        if Path(SETTINGS_FILE).exists():
            return json.loads(Path(SETTINGS_FILE).read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_config(data: dict) -> None:
    try:
        Path(SETTINGS_FILE).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

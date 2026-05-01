# Tin Dung Tool - bản chỉnh sửa ổn định

Ứng dụng Streamlit để nhập hồ sơ tín dụng, quét QR CCCD, áp dụng phương án vay vốn từ Excel và xuất hồ sơ Word.

## 1. Cấu trúc thư mục

```text
tin-dung-tool/
├── config.py
├── packages.txt
├── requirements.txt
├── web_app.py
└── core/
    ├── __init__.py
    ├── document.py
    ├── scanner.py
    └── utils.py
```

## 2. Các file cần thay thế

Chép đè các file sau vào repo hiện tại:

```text
web_app.py
core/utils.py
core/scanner.py
requirements.txt
```

`core/document.py`, `config.py`, `packages.txt` có thể giữ nguyên.

## 3. Cài đặt local

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows
pip install --upgrade pip
pip install -r requirements.txt
```

## 4. Chạy app

```bash
streamlit run web_app.py
```

Sau đó mở trình duyệt theo URL Streamlit trả về.

## 5. Thứ tự kiểm tra sau khi chạy

1. Upload `Data.xlsx`.
2. Kiểm tra app render đủ các tab/trường.
3. Upload `Mẫu Word.docx`.
4. Nhập thử tên, vốn xin vay, vốn tự có, lãi suất năm.
5. Kiểm tra tổng vốn, lãi suất tháng, tiền bằng chữ.
6. Upload ảnh QR CCCD và bấm quét từng đối tượng.
7. Upload `Phương án.xlsx`, chọn phương án và bấm áp dụng.
8. Bấm xuất Word và mở file tải về.

## 6. Định dạng file Phương án.xlsx

Cần có 2 sheet:

### Sheet `phuongan`

| Tên PA | Mã PA |
|---|---|
| Phương án A | PA01 |

### Sheet `data`

| Mã phương án | Loại | Tỉ lệ | Hạng mục | Nội dung chi tiết |
|---|---|---:|---|---|
| PA01 | Chi phí | 0.1 | Giống cây | Mua giống cây |
| PA01 | Thu nhập | 0.2 | Bán hàng | Doanh thu bán sản phẩm |

Lưu ý: cột `Tỉ lệ` nên nhập dạng số thập phân, ví dụ `0.1` tương ứng 10%.

## 7. Những lỗi đã sửa

- Sửa lỗi callback gọi `execute_financial_engine` nhưng app chỉ định nghĩa `calculate_globals`.
- Sửa lỗi upload Word không được lưu nhưng xuất lại dùng `temp_template.docx` cố định.
- Sửa lỗi thiếu import `is_name_desc` khi đặt tên file xuất.
- Đổi file tạm cố định sang `tempfile` để tránh ghi đè giữa nhiều người dùng.
- Thêm cache cho đọc Excel mapping và phương án.
- Pin version dependencies trong `requirements.txt`.
- Gom logic format/parse tiền, số, phần trăm, diện tích, bằng chữ về `core/utils.py`.
- Tối ưu scanner bằng cách giới hạn ảnh lớn trước khi thử nhiều biến thể QR.
- Dùng `MappingBundle`/`FieldDefinition` từ `core/document.py` để render form tốt hơn.

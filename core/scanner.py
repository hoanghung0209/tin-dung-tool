from __future__ import annotations

try:
    import cv2
    import numpy as np
    import zxingcpp
except Exception:  # pragma: no cover - runtime dependency
    cv2 = None
    np = None
    zxingcpp = None


class ScannerUnavailableError(RuntimeError):
    """Raised when QR scanning dependencies are not installed."""


def _ensure_runtime() -> None:
    if cv2 is None or np is None or zxingcpp is None:
        raise ScannerUnavailableError(
            "Thiếu thư viện quét QR (cv2, numpy, zxingcpp). Hãy cài đủ dependencies để dùng chức năng này."
        )


def _variants(img_bgr):
    outs = []
    green_channel = img_bgr[:, :, 1]
    outs.append(green_channel)
    outs.append(cv2.equalizeHist(green_channel))
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    outs.append(clahe.apply(green_channel))
    blur = cv2.GaussianBlur(green_channel, (0, 0), 1.2)
    sharp = cv2.addWeighted(green_channel, 1.6, blur, -0.6, 0)
    outs.append(sharp)
    _, th_otsu = cv2.threshold(green_channel, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    outs.append(th_otsu)
    th_adapt = cv2.adaptiveThreshold(
        green_channel, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 5
    )
    outs.append(th_adapt)
    outs.append(cv2.resize(green_channel, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC))
    outs.append(cv2.resize(green_channel, None, fx=3.0, fy=3.0, interpolation=cv2.INTER_CUBIC))
    return outs


def decode_qr_offline(image_path: str) -> str:
    _ensure_runtime()
    try:
        img_array = np.fromfile(image_path, np.uint8)
        img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    except Exception as exc:
        raise FileNotFoundError(f"Lỗi hệ thống khi đọc file: {exc}") from exc

    if img is None:
        raise FileNotFoundError(f"Không thể giải mã ảnh tại: {image_path}.")

    for variant in _variants(img):
        result = zxingcpp.read_barcode(variant)
        if result and result.text:
            return result.text

    for angle in (-15, -10, -7, -5, 5, 7, 10, 15):
        height, width = img.shape[:2]
        matrix = cv2.getRotationMatrix2D((width / 2, height / 2), angle, 1.0)
        rotated = cv2.warpAffine(
            img,
            matrix,
            (width, height),
            flags=cv2.INTER_LINEAR,
            borderMode=cv2.BORDER_REPLICATE,
        )
        for variant in _variants(rotated):
            result = zxingcpp.read_barcode(variant)
            if result and result.text:
                return result.text

    raise ValueError(
        "Không decode được QR. Ảnh bị mờ, lóa sáng hoặc QR quá nhỏ. Hãy crop sát QR hoặc chụp lại ảnh rõ hơn."
    )

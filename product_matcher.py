"""
product_matcher.py
-------------------
Đối chiếu Tên hàng trên hóa đơn PDF <-> Mã hàng trong danh mục MISA đã có sẵn.

Vì tên gọi lệch nhau (VD: "Trà sữa Hồng Bao (M)" trên PDF
   <-> "Trà sữa Hồng Bao-Size M" trong MISA),
dùng fuzzy matching thay vì so khớp chuỗi tuyệt đối.

Yêu cầu: pip install rapidfuzz pandas
Danh mục hàng hóa: xuất từ MISA ra Excel/CSV với ít nhất 2 cột: MaHang, TenHang
"""

import os
import re
from dataclasses import dataclass
from typing import Optional
import pandas as pd
from rapidfuzz import process, fuzz

# Ngưỡng tin cậy: dưới mức này -> gắn cờ "cần xác nhận tay", không tự nhập.
CONFIDENCE_AUTO_ACCEPT = 90   # >= mức này: tự động chọn, không cần người duyệt
CONFIDENCE_MIN_SUGGEST = 60   # dưới mức này: không gợi ý, để trống bắt buộc chọn tay


@dataclass
class MatchResult:
    ten_hang_pdf: str
    ma_hang: Optional[str]
    ten_hang_misa: Optional[str]
    score: float
    can_auto_accept: bool


def _normalize(s: str) -> str:
    """Chuẩn hoá chuỗi để so khớp công bằng hơn: bỏ hoa/thường, ký tự phụ, khoảng trắng thừa."""
    s = s.lower()
    s = s.replace("(", " ").replace(")", " ").replace("-", " ")
    s = re.sub(r"\bsize\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


class ProductMatcher:
    def __init__(self, catalog_csv_path: str):
        """
        catalog_csv_path: file CSV/Excel xuất từ danh mục hàng hóa MISA.
        Cột bắt buộc: 'MaHang', 'TenHang'
        """
        df = pd.read_csv(catalog_csv_path) if catalog_csv_path.endswith(".csv") else pd.read_excel(catalog_csv_path)
        self.catalog = df
        self._norm_names = {row.TenHang: _normalize(row.TenHang) for row in df.itertuples()}
        # map: tên đã chuẩn hoá -> (MaHang, TenHang gốc)
        self._lookup = {
            _normalize(row.TenHang): (row.MaHang, row.TenHang) for row in df.itertuples()
        }

    def match(self, ten_hang_pdf: str) -> MatchResult:
        norm_query = _normalize(ten_hang_pdf)
        choices = list(self._lookup.keys())

        best = process.extractOne(norm_query, choices, scorer=fuzz.WRatio)
        if best is None:
            return MatchResult(ten_hang_pdf, None, None, 0.0, False)

        matched_norm, score, _ = best
        ma_hang, ten_hang_misa = self._lookup[matched_norm]

        return MatchResult(
            ten_hang_pdf=ten_hang_pdf,
            ma_hang=ma_hang if score >= CONFIDENCE_MIN_SUGGEST else None,
            ten_hang_misa=ten_hang_misa if score >= CONFIDENCE_MIN_SUGGEST else None,
            score=score,
            can_auto_accept=score >= CONFIDENCE_AUTO_ACCEPT,
        )


# --- Ví dụ danh mục mẫu (để test nhanh khi chưa có file catalog thật) ---
SAMPLE_CATALOG = [
    {"MaHang": "TP0021", "TenHang": "Trà sữa Hồng Bao-Size M"},
    {"MaHang": "TP0022", "TenHang": "Trà sữa Hồng Bao-Size XL"},
    {"MaHang": "TP0031", "TenHang": "Trà bí đao sương sáo-Size XL"},
    {"MaHang": "TP0045", "TenHang": "Trân châu 3Q"},
    {"MaHang": "TP0046", "TenHang": "Trân châu đen"},
    {"MaHang": "TP0052", "TenHang": "Kem trứng"},
]


def build_sample_matcher() -> ProductMatcher:
    """Dùng cho test cục bộ, không cần file catalog thật."""
    import tempfile

    df = pd.DataFrame(SAMPLE_CATALOG)
    tmp_path = os.path.join(tempfile.gettempdir(), "_sample_catalog.csv")
    df.to_csv(tmp_path, index=False)
    return ProductMatcher(tmp_path)
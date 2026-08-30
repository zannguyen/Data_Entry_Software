"""
pdf_extractor.py
----------------
Đọc hóa đơn điện tử PDF (dạng có lớp text, ví dụ hệ thống MobiFone Invoice)
và trích xuất dữ liệu theo NHÃN TRƯỜNG (label-based), không theo tọa độ cố định.

Yêu cầu: pip install pdfplumber pydantic
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional
import pdfplumber


@dataclass
class InvoiceLine:
    ten_hang: str
    dvt: str
    so_luong: float
    don_gia: float
    thanh_tien_truoc_thue: float
    thue_suat_pct: float
    tien_thue: float
    thanh_tien: float


@dataclass
class Invoice:
    so_hoa_don: str
    ky_hieu: str
    ngay: str                      # dd/mm/yyyy
    nguoi_ban: str
    mst_nguoi_ban: str
    nguoi_mua: str
    lines: List[InvoiceLine] = field(default_factory=list)
    tong_tien_truoc_thue: float = 0.0
    tong_tien_thue: float = 0.0
    tong_thanh_toan: float = 0.0
    so_tien_bang_chu: str = ""
    file_path: str = ""

    # cờ cảnh báo — dùng để quyết định có cần người kiểm tra tay không
    warnings: List[str] = field(default_factory=list)


def _to_number(s: str) -> float:
    """'40.741' hoặc '40,741' -> 40741.0 (định dạng số kiểu Việt Nam)."""
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _find(pattern: str, text: str, group: int = 1, default: str = "") -> str:
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(group).strip() if m else default


def extract_invoices(pdf_path: str) -> List[Invoice]:
    """
    Trích xuất TẤT CẢ hóa đơn trong 1 file PDF.

    QUAN TRỌNG: 1 file PDF không phải lúc nào cũng chỉ chứa 1 hóa đơn — file gộp
    nhiều trang (mỗi trang 1 hóa đơn) khá phổ biến khi tải hàng loạt từ hệ thống
    hóa đơn điện tử. Hàm này xử lý THEO TỪNG TRANG để không bị gộp nhầm dữ liệu
    của nhiều hóa đơn vào làm một (lỗi thực tế đã gặp khi test).
    """
    invoices: List[Invoice] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            if "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" not in text.upper() and "VAT INVOICE" not in text.upper():
                continue  # trang không phải hóa đơn (trang trắng, phụ lục...) -> bỏ qua
            tables = page.extract_tables() or []
            inv = _parse_single_invoice(text, tables, pdf_path, page_no)
            invoices.append(inv)
    return invoices


def extract_invoice(pdf_path: str) -> Invoice:
    """Tiện ích cho file chắc chắn chỉ có 1 hóa đơn/1 trang. Nếu nhiều trang, chỉ lấy trang đầu."""
    result = extract_invoices(pdf_path)
    if not result:
        raise ValueError(f"Không tìm thấy hóa đơn hợp lệ trong {pdf_path}")
    if len(result) > 1:
        raise ValueError(
            f"{pdf_path} chứa {len(result)} hóa đơn (nhiều trang) — dùng extract_invoices() thay vì extract_invoice()."
        )
    return result[0]


def _parse_single_invoice(text: str, tables: list, pdf_path: str, page_no: int) -> Invoice:
    """
    Hỗ trợ 2 định dạng hóa đơn điện tử đã gặp thực tế:
      1. Định dạng gốc dùng để test ban đầu (hệ thống MobiFone Invoice, file mẫu
         Hồng Bao) — nhãn trường đầy đủ kiểu "Số (No.): 911", "Tên người bán:...".
      2. Định dạng THẬT của công ty (MISA meInvoice, xuất từ chính MISA SME.NET —
         hóa đơn do CÔNG TY TNHH GUABAO ONE tự phát hành) — nhãn ngắn gọn hơn, và
         pdfplumber trích chữ SÁT NHAU KHÔNG CÓ KHOẢNG TRẮNG giữa nhãn và số/chữ kế
         tiếp (VD "Ngày01tháng07năm2026", "Ký hiệu:1C26MTT", "Số: 00000631") do PDF
         gốc không có khoảng trắng ở đó (phải dùng \\s* thay vì \\s+ trong regex).
      Định dạng 2 là dữ liệu THẬT đang dùng — ưu tiên đúng cho định dạng này.
    """
    # "Số:" (định dạng thật) hoặc "Số (No.):" (định dạng mẫu cũ)
    so_hoa_don = _find(r"S[ốo]\s*(?:\(No\.?\))?\s*[:\.]\s*(\d+)", text)
    # "Ký hiệu:1C26MTT" (thật) hoặc "Ký hiệu ... (Serial No): XXX" (mẫu cũ)
    ky_hieu = _find(r"K[ýy]\s*hi[ệe]u\s*(?:\(Serial No\))?\s*[:\.]\s*([A-Z0-9]+)", text)
    ngay_m = re.search(r"Ng[àa]y\s*(\d{1,2})\s*th[áa]ng\s*(\d{1,2})\s*n[ăa]m\s*(\d{4})", text, re.IGNORECASE)
    ngay_fmt = f"{ngay_m.group(1).zfill(2)}/{ngay_m.group(2).zfill(2)}/{ngay_m.group(3)}" if ngay_m else ""

    # Tên người bán: định dạng thật KHÔNG có nhãn "Tên người bán" — tên công ty nằm
    # ngay dòng dưới "Mã CQT:...", trước dòng "Mã số thuế:...". Fallback về nhãn cũ
    # nếu có (định dạng mẫu).
    nguoi_ban = _find(r"T[êe]n\s+ng[ưu][ờo]i\s+b[áa]n.*?:\s*(.+)", text)
    if not nguoi_ban:
        m = re.search(r"M[ãa]\s*CQT\s*:\s*\S+\s*\n(.+?)\s*\n\s*M[ãa]\s+s[ốo]\s+thu[ếe]", text, re.IGNORECASE)
        nguoi_ban = m.group(1).strip() if m else ""

    mst = _find(r"M[ãa]\s+s[ốo]\s+thu[ếe]\s*[:\.]\s*([\d]+)", text)

    # Người mua: định dạng thật ghi "Tên đơn vị:Khách lẻ" (không phải "Họ tên người mua").
    nguoi_mua = _find(r"H[ọo]\s+t[êe]n\s+ng[ưu][ờo]i\s+mua.*?:\s*(.+)", text)
    if not nguoi_mua:
        nguoi_mua = _find(r"T[êe]n\s+đ[ơo]n\s+v[ịi]\s*:\s*(.+)", text)

    # Dòng tổng: định dạng thật không có nhãn "Tổng tiền trước thuế" / "Tổng tiền
    # thanh toán" riêng — cả 3 số nằm chung 1 dòng "Tổng cộng: X Y Z".
    tong_truoc_thue = _to_number(_find(r"T[ổo]ng\s+ti[ềe]n\s+tr[ưu][ớo]c\s+thu[ếe].*?:\s*([\d\.,]+)", text))
    tong_thue = _to_number(_find(r"T[ổo]ng\s+ti[ềe]n\s+thu[ếe]\s+GTGT.*?:\s*([\d\.,]+)", text))
    tong_thanh_toan = _to_number(_find(r"T[ổo]ng\s+ti[ềe]n\s+thanh\s+to[áa]n.*?:\s*([\d\.,]+)", text))
    if not tong_truoc_thue and not tong_thanh_toan:
        m = re.search(
            r"T[ổo]ng\s+c[ộo]ng\s*:\s*([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)", text, re.IGNORECASE
        )
        if m:
            tong_truoc_thue = _to_number(m.group(1))
            tong_thue = _to_number(m.group(2))
            tong_thanh_toan = _to_number(m.group(3))
    so_tien_chu = _find(r"S[ốo]\s+ti[ềe]n\s+vi[ếe]t\s+b[ằa]ng\s+ch[ữu].*?:\s*(.+)", text)

    inv = Invoice(
        so_hoa_don=so_hoa_don,
        ky_hieu=ky_hieu,
        ngay=ngay_fmt,
        nguoi_ban=nguoi_ban,
        mst_nguoi_ban=mst,
        nguoi_mua=nguoi_mua,
        tong_tien_truoc_thue=tong_truoc_thue,
        tong_tien_thue=tong_thue,
        tong_thanh_toan=tong_thanh_toan,
        so_tien_bang_chu=so_tien_chu,
        file_path=f"{pdf_path}#trang{page_no}",
    )

    # --- Trích bảng hàng hóa ---
    # Định dạng thật (MISA meInvoice) có 11 cột, xen kẽ cột None do merge cell:
    #   [0]=STT [1]=Tên hàng [2]=None [3]=ĐVT [4]=SL [5]=Đơn giá [6]=None
    #   [7]=Thành tiền [8]=Thuế suất% [9]=None [10]=Tiền thuế GTGT
    # Định dạng mẫu MobiFone Invoice (Hồng Bao) có 9 cột liền, KHÔNG có cột None xen kẽ,
    # NHƯNG VẪN CÓ đủ cột thuế suất/tiền thuế/thành tiền cuối riêng theo từng dòng:
    #   [0]=STT [1]=Tên hàng [2]=ĐVT [3]=SL [4]=Đơn giá [5]=Thành tiền trước thuế
    #   [6]=Thuế suất% [7]=Tiền thuế GTGT [8]=Thành tiền (Total)
    # XÁC NHẬN THẬT (30/08/2026, đọc lại bảng thô qua extract_tables()): TRƯỚC ĐÂY code
    # coi định dạng này là "6 cột, không có cột thuế riêng" (SAI — bảng thật có đủ 9
    # cột) -> Tiền thuế GTGT mỗi dòng bị đọc thành 0.0 một cách ÂM THẦM, không cảnh báo
    # nào bắt được (vì tổng Tiền thuế GTGT toàn hóa đơn vẫn đọc đúng từ dòng nhãn riêng
    # "Tổng tiền thuế GTGT", nên _cross_validate() không phát hiện ra sai số ở đây). Nếu
    # định dạng này từng được nạp thật vào MISA, MỌI dòng đều bị ghi đè Tiền thuế GTGT =
    # 0 một cách sai lệch — chỉ giữ 1 nhánh dự phòng tối thiểu (< 9 cột) cho trường hợp
    # bảng thật sự không có cột thuế riêng.
    seen_rows = set()
    for table in tables:
        for row in table:
            if not row or len(row) < 6:
                continue
            # Bỏ dòng tiêu đề / dòng tổng
            joined = " ".join(c or "" for c in row).lower()
            if "stt" in joined or "tổng" in joined or "description" in joined:
                continue

            stt = (row[0] or "").strip()
            # Bỏ dòng rác: STT không phải số nguyên nhỏ (dòng tiêu đề công thức kiểu
            # "1 2 3 4 5 6=4x5..." bị đọc lẫn vào bảng, hoặc dòng "Tổng hợp"/"Thuế suất X%:")
            if not stt.isdigit():
                continue

            try:
                if len(row) >= 11:
                    # Định dạng thật (MISA meInvoice): có cột None xen kẽ do merge cell.
                    ten = (row[1] or "").strip()
                    dvt = (row[3] or "").strip()
                    sl = row[4]
                    dg = row[5]
                    tt_truoc = row[7]
                    thue_suat_str = (row[8] or "").strip()
                    tien_thue_str = row[10]
                    tt_final_str = None
                elif len(row) >= 9:
                    # Định dạng mẫu MobiFone Invoice (Hồng Bao): 9 cột liền, ĐỦ cột thuế
                    # suất/tiền thuế/thành tiền cuối riêng theo từng dòng.
                    ten = (row[1] or "").strip()
                    dvt = (row[2] or "").strip()
                    sl = row[3]
                    dg = row[4]
                    tt_truoc = row[5]
                    thue_suat_str = (row[6] or "").strip()
                    tien_thue_str = row[7]
                    tt_final_str = row[8]
                else:
                    # Dự phòng tối thiểu: bảng thật sự không có cột thuế riêng theo dòng.
                    _, ten, dvt, sl, dg, tt_truoc = row[:6]
                    ten = (ten or "").strip()
                    dvt = (dvt or "").strip()
                    thue_suat_str = ""
                    tien_thue_str = None
                    tt_final_str = None

                if ten.isdigit() or len(ten) < 2:
                    continue
                row_key = (ten, sl, dg)
                if row_key in seen_rows:
                    continue
                seen_rows.add(row_key)

                # "8%" -> 8.0 ; nếu không đọc được, mặc định 8.0 theo hóa đơn mẫu ban đầu
                thue_suat_pct = _to_number(thue_suat_str.replace("%", "")) if thue_suat_str else 8.0

                line = InvoiceLine(
                    ten_hang=ten,
                    dvt=dvt,
                    so_luong=_to_number(sl),
                    don_gia=_to_number(dg),
                    thanh_tien_truoc_thue=_to_number(tt_truoc),
                    thue_suat_pct=thue_suat_pct,
                    tien_thue=_to_number(tien_thue_str) if tien_thue_str else 0.0,
                    thanh_tien=_to_number(tt_final_str) if tt_final_str else 0.0,
                )
                inv.lines.append(line)
            except ValueError:
                continue

    _cross_validate(inv)
    return inv


def _cross_validate(inv: Invoice) -> None:
    """
    Đối chiếu chéo số liệu để tự phát hiện sai số đọc PDF.
    Gắn cờ vào inv.warnings — hóa đơn có cảnh báo PHẢI được người kiểm tra tay
    trước khi cho script tự động nhập vào MISA.
    """
    if inv.lines:
        sum_lines = sum(l.thanh_tien_truoc_thue for l in inv.lines)
        if abs(sum_lines - inv.tong_tien_truoc_thue) > 1:  # sai số làm tròn 1đ
            inv.warnings.append(
                f"Tổng các dòng hàng ({sum_lines:,.0f}) không khớp Tổng tiền trước thuế "
                f"({inv.tong_tien_truoc_thue:,.0f})"
            )

        # BỔ SUNG (30/08/2026): trước đây KHÔNG đối chiếu tổng Tiền thuế GTGT theo từng
        # dòng với tổng khai báo trên hóa đơn -> lỗi thật đã lọt qua (Tiền thuế/dòng bị
        # đọc thành 0.0 cho định dạng MobiFone Invoice 9 cột, do code hiểu nhầm bảng chỉ
        # có 6 cột) mà KHÔNG có cảnh báo nào bắt được, vì tổng thuế toàn hóa đơn vẫn đọc
        # đúng từ dòng nhãn riêng. Thêm đối chiếu này để bắt được lớp lỗi tương tự trong
        # tương lai (VD gặp thêm 1 định dạng PDF mới mà code chưa hỗ trợ đúng).
        sum_tien_thue = sum(l.tien_thue for l in inv.lines)
        if abs(sum_tien_thue - inv.tong_tien_thue) > 1:
            inv.warnings.append(
                f"Tổng Tiền thuế GTGT các dòng hàng ({sum_tien_thue:,.0f}) không khớp "
                f"Tổng tiền thuế GTGT khai báo ({inv.tong_tien_thue:,.0f})"
            )

    calc_total = inv.tong_tien_truoc_thue + inv.tong_tien_thue
    if abs(calc_total - inv.tong_thanh_toan) > 1:
        inv.warnings.append(
            f"Tiền hàng + Tiền thuế ({calc_total:,.0f}) không khớp Tổng thanh toán "
            f"({inv.tong_thanh_toan:,.0f})"
        )

    if not inv.so_hoa_don:
        inv.warnings.append("Không đọc được Số hóa đơn — kiểm tra định dạng PDF")
    if not inv.lines:
        inv.warnings.append("Không trích được dòng hàng hóa nào — kiểm tra extract_tables()")

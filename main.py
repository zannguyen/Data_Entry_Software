"""
main.py
--------
Script điều phối chính: quét thư mục hóa đơn PDF -> trích xuất -> đối chiếu mã hàng
-> tự động nhập vào MISA SME.NET 2019 theo đúng thứ tự Số hóa đơn tăng dần.

Cách chạy:
    python main.py --pdf-dir ./hoa_don --catalog ./danh_muc_hang_hoa.csv --dry-run
    python main.py --pdf-dir ./hoa_don --catalog ./danh_muc_hang_hoa.csv        # chạy thật

QUY TẮC NGHIỆP VỤ ĐÃ ÁP DỤNG TRONG SCRIPT NÀY:
  1. Khách hàng luôn = "Khách lẻ", bất kể PDF ghi tên gì.
  2. Trạng thái luôn = "Chưa thu tiền".
  3. Thuế suất trong MISA chọn cố định 10%, nhưng số Tiền thuế được GHI ĐÈ
     bằng đúng số tiền thuế thật trên hóa đơn (vốn tính theo 8%).
  4. Mã hàng đối chiếu gần đúng (fuzzy match) với danh mục MISA đã có sẵn;
     hóa đơn có dòng hàng độ tin cậy thấp sẽ DỪNG và đẩy vào hàng đợi duyệt tay,
     không tự đoán bừa.
  5. Xử lý tuần tự theo Số hóa đơn tăng dần.
"""

import argparse
import glob
import logging
import os
import sys
from pathlib import Path

from pdf_extractor import extract_invoices, Invoice
from product_matcher import ProductMatcher, build_sample_matcher
from misa_automation import (
    MisaAutomation, InvoiceToEnter, LineToEnter, LowConfidenceMatchError,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("main")


def load_and_sort_invoices(pdf_dir: str) -> list:
    """
    Đọc toàn bộ PDF trong thư mục, sắp xếp theo Số hóa đơn tăng dần.
    Mỗi file PDF có thể chứa NHIỀU hóa đơn (nhiều trang) — xử lý theo từng trang.
    """
    paths = sorted(glob.glob(os.path.join(pdf_dir, "*.pdf")))
    invoices = []
    for p in paths:
        try:
            for inv in extract_invoices(p):
                invoices.append(inv)
        except Exception as e:  # noqa: BLE001 — log đủ chi tiết để soát lỗi từng file
            log.error("Lỗi đọc file %s: %s", p, e)

    invoices.sort(key=lambda i: int(i.so_hoa_don) if i.so_hoa_don.isdigit() else 0)
    return invoices


def to_entry_payload(inv: Invoice, matcher: ProductMatcher):
    """Chuyển Invoice (dữ liệu đọc từ PDF) -> InvoiceToEnter (dữ liệu sẽ điền vào MISA)."""
    lines = [
        LineToEnter(
            ma_hang=l.ten_hang, so_luong=l.so_luong, don_gia=l.don_gia,
            tien_thue_dong=l.tien_thue,
        )
        for l in inv.lines
    ]
    matches = [matcher.match(l.ten_hang) for l in inv.lines]

    payload = InvoiceToEnter(
        ngay=inv.ngay,
        lines=lines,
        tien_thue_that=inv.tong_tien_thue,
        tong_tien_hang=inv.tong_tien_truoc_thue,
        tong_thanh_toan=inv.tong_thanh_toan,
        so_hoa_don_goc=inv.so_hoa_don,
    )
    return payload, matches


def run(pdf_dir: str, catalog_path: str, dry_run: bool):
    log.info("=== BẮT ĐẦU QUY TRÌNH NHẬP HÓA ĐƠN (dry_run=%s) ===", dry_run)

    invoices = load_and_sort_invoices(pdf_dir)
    log.info("Tìm thấy %d hóa đơn, đã sắp xếp theo số tăng dần.", len(invoices))

    matcher = build_sample_matcher() if catalog_path is None else ProductMatcher(catalog_path)

    automation = MisaAutomation(dry_run=dry_run)
    if not dry_run:
        automation.connect()

    results = {"da_cat": [], "can_duyet_tay": [], "loi": []}

    for inv in invoices:
        log.info("---- Xử lý hóa đơn số %s (%s) ----", inv.so_hoa_don, Path(inv.file_path).name)

        if inv.warnings:
            log.warning("Hóa đơn %s có cảnh báo đối chiếu số liệu: %s", inv.so_hoa_don, inv.warnings)
            results["can_duyet_tay"].append((inv.so_hoa_don, inv.warnings))
            continue  # KHÔNG tự động nhập hóa đơn có sai lệch số liệu

        payload, matches = to_entry_payload(inv, matcher)

        try:
            popup = automation.open_new_invoice_popup()
            automation.fill_invoice(popup, payload, matches)
            automation.save(popup)
            results["da_cat"].append(inv.so_hoa_don)
            log.info("Đã cất hóa đơn số %s thành công.", inv.so_hoa_don)

        except LowConfidenceMatchError as e:
            log.warning("Hóa đơn %s tạm dừng — %s", inv.so_hoa_don, e)
            results["can_duyet_tay"].append((inv.so_hoa_don, [str(e)]))
            automation.close_popup_if_open(locals().get("popup"))

        except Exception as e:  # noqa: BLE001
            log.error("Lỗi khi nhập hóa đơn %s: %s", inv.so_hoa_don, e)
            results["loi"].append((inv.so_hoa_don, str(e)))
            automation.close_popup_if_open(locals().get("popup"))

    log.info("=== KẾT THÚC ===")
    log.info("Đã cất: %s", results["da_cat"])
    log.info("Cần duyệt tay: %s", [r[0] for r in results["can_duyet_tay"]])
    log.info("Lỗi: %s", [r[0] for r in results["loi"]])
    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="RPA nhập hóa đơn PDF vào MISA SME.NET 2019")
    parser.add_argument("--pdf-dir", required=True, help="Thư mục chứa các file PDF hóa đơn")
    parser.add_argument("--catalog", default=None, help="File CSV/Excel danh mục hàng hóa MISA (MaHang, TenHang)")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ log ra các bước sẽ làm, KHÔNG thao tác thật lên MISA")
    args = parser.parse_args()

    run(args.pdf_dir, args.catalog, dry_run=args.dry_run)
    
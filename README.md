# RPA nhập hóa đơn PDF vào MISA SME.NET 2019

## Cài đặt
```
pip install -r requirements.txt --break-system-packages
```

## Cấu trúc
- `pdf_extractor.py`  — đọc PDF hóa đơn theo nhãn trường, tự đối chiếu số liệu chéo.
  Đã kiểm chứng với file hóa đơn thật (911/912/913): đọc đúng 100% dữ liệu số.
- `product_matcher.py` — đối chiếu Tên hàng PDF ↔ Mã hàng MISA bằng fuzzy matching.
- `misa_automation.py` — điều khiển giao diện MISA bằng pywinauto (CẦN dò lại
  auto_id thật trên máy chạy MISA — xem hướng dẫn "DÒ CONTROL" cuối file).
- `main.py` — script điều phối chính, chạy toàn bộ luồng.

## Chạy thử an toàn (không đụng vào MISA thật)
```
python main.py --pdf-dir ./test_pdfs --dry-run
```

## Chạy thật (sau khi đã dò control thật và test kỹ)
```
python main.py --pdf-dir ./hoa_don --catalog ./danh_muc_hang_hoa.csv
```

## Quy tắc nghiệp vụ đã áp dụng
1. Khách hàng luôn = "Khách lẻ".
2. Trạng thái luôn = "Chưa thu tiền".
3. Thuế suất chọn 10% trong MISA, Tiền thuế ghi đè bằng số thật trên hóa đơn (8%).
4. Mã hàng đối chiếu gần đúng; độ tin cậy thấp -> dừng, đẩy vào hàng đợi duyệt tay.
5. Hóa đơn có sai lệch đối chiếu số liệu (SL×ĐG ≠ Thành tiền, Tổng không khớp...)
   -> KHÔNG tự động nhập, đẩy vào hàng đợi duyệt tay.
6. Xử lý tuần tự theo Số hóa đơn tăng dần.
7. 1 file PDF có thể chứa NHIỀU hóa đơn (nhiều trang) -> xử lý theo từng trang,
   không gộp nhầm dữ liệu (lỗi thực tế đã phát hiện và sửa khi test).

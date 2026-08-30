# Dự án: RPA nhập hóa đơn PDF vào MISA SME.NET 2019

Tài liệu này tổng hợp TOÀN BỘ ngữ cảnh đã trao đổi trong phiên làm việc trên claude.ai
trước khi chuyển sang Claude Code. Đọc kỹ trước khi tiếp tục bất kỳ công việc nào.

---

## 1. Bối cảnh & mục tiêu tổng thể

Công ty đang dùng **MISA SME.NET 2019 R7 Enterprise** (phần mềm kế toán desktop Windows
dành cho SME) để hạch toán bán hàng. Quy trình hiện tại là NHÂN VIÊN TỰ TAY đọc từng hóa
đơn bán lẻ (PDF hóa đơn điện tử) rồi nhập thủ công vào MISA qua màn hình
**Bán hàng ▸ Chứng từ bán hàng ▸ Thêm**.

**Mục tiêu:** xây dựng công cụ (RPA) đọc hàng loạt PDF hóa đơn, tự động điền vào MISA theo
đúng thứ tự và đúng quy tắc nghiệp vụ đặc thù của công ty, giảm việc nhập tay.

Thao tác trên MISA hiện được thực hiện **qua UltraViewer** (remote desktop) — người dùng
ngồi máy A, điều khiển máy B (máy cài MISA thật) qua UltraViewer 6.6 Free.

---

## 2. Giao diện MISA SME.NET 2019 — đã phân tích từ ảnh chụp màn hình thật

### Màn hình danh sách chứng từ bán hàng
Menu **Bán hàng ▸ tab Bán hàng**, gồm 2 phần:
- **Header (trên):** mỗi dòng = 1 hóa đơn đã nhập — Ngày hạch toán, Ngày chứng từ,
  Số chứng từ (dạng `BH00xxx` — MISA tự sinh tuần tự), Số hóa đơn (dạng `00000xxx` —
  số hóa đơn thật do cơ quan thuế cấp), Khách hàng, Diễn giải, Tổng tiền hàng,
  Tiền chiết khấu, Tiền thuế GTGT.
- **Chi tiết (dưới):** khi chọn 1 dòng ở trên, hiện chi tiết mặt hàng: Mã hàng, Tên hàng,
  TK công nợ/chi phí (= 131), TK doanh thu (= 5111), ĐVT, Số lượng, Đơn giá, Thành tiền,
  Tỷ lệ CK (%).

### Popup nhập hóa đơn mới
Bấm **Thêm** trên thanh công cụ → mở popup **"Chứng từ bán hàng"**, loại
**"1. Bán hàng hóa, dịch vụ trong nước"**, với các trường:
- Radio button **Chưa thu tiền / Thu tiền ngay**
- Ô **Khách hàng** (combobox có nút thêm mới +)
- **Mã số thuế**, **Người liên hệ**, **Địa chỉ**
- **NV bán hàng**, **Tham chiếu**
- **Điều khoản TT**, **Số ngày được nợ**, **Hạn thanh toán**
- **Ngày hạch toán**, **Ngày chứng từ**, **Số chứng từ** (tự sinh)
- Các tab: **1. Hàng tiền** (bảng chi tiết hàng hóa), **2. Thuế**, **3. Giá vốn**,
  **4. Thống kê**, **5. Khác**
- Bảng "Hàng tiền": Mã hàng, Tên hàng, TK công nợ/chi phí, TK doanh thu, ĐVT, Số lượng,
  Đơn giá, Thành tiền, Tỷ lệ CK (%)
- Cuối form: **Phân bổ chiết khấu**, **Xem công nợ**, và 3 tổng: Tổng tiền hàng,
  Tiền chiết khấu, Tiền thuế GTGT, Tổng tiền thanh toán
- Nút **Cất** trên toolbar để lưu chứng từ

Công ty đang thao tác trên đơn vị "GUABAO ONE" (nhà hàng) trong ảnh chụp mẫu ban đầu.

### Mẫu hóa đơn PDF nguồn (đã đọc thật 3 hóa đơn mẫu)
Hóa đơn điện tử **HÓA ĐƠN GIÁ TRỊ GIA TĂNG** khởi tạo từ **hệ thống MobiFone Invoice**,
công ty bán: **CÔNG TY TRÁCH NHIỆM HỮU HẠN HỒNG BAO** (trà sữa), MST 5901251350.
File PDF mẫu (`hoa_don.pdf`) do người dùng upload **thực chất chứa 3 hóa đơn trong 3
trang** — số 911, 912, 913, ngày 28/08/2026, thuế suất 8% trên từng hóa đơn.

Cấu trúc dữ liệu mỗi hóa đơn: Ký hiệu (Serial No), Số (No), Mã cơ quan thuế, Tên/địa
chỉ/MST người bán, Họ tên người mua, bảng hàng hóa (Stt, Tên hàng hóa dịch vụ, Đvt, Số
lượng, Đơn giá, Thành tiền trước thuế, Thuế suất %VAT, Tiền thuế, Thành tiền), dòng tổng
(Tổng tiền trước thuế / Tổng tiền thuế GTGT / Tổng tiền thanh toán), số tiền bằng chữ,
chữ ký số + mã nhận hóa đơn + link tra cứu.

**Dữ liệu 3 hóa đơn mẫu đã dùng để test (đọc thật, khớp 100%):**

| Số HĐ | Người mua | Mặt hàng | SL | Đơn giá | Tổng trước thuế | Thuế (8%) | Tổng TT |
|---|---|---|---|---|---|---|---|
| 911 | Bán cho người tiêu dùng | Trân châu đen (Phần) | 1 | 3.889 | 19.630 | 1.570 | 21.200 |
| | | Trà sữa Hồng Bao (XL) (Ly) | 1 | 15.741 | | | |
| 912 | TRAN BAO QUAN | Trà sữa Hồng Bao (M) (Ly) | 1 | 8.333 | 16.666 | 1.334 | 18.000 |
| | | KEM TRỨNG (Phần) | 1 | 8.333 | | | |
| 913 | Bán cho người tiêu dùng | TRÀ BÍ ĐAO SƯƠNG SÁO (XL) (Ly) | 2 | 17.593 | 40.741 | 3.259 | 44.000 |
| | | TRÂN CHÂU 3Q (Phần) | 1 | 5.556 | | | |

---

## 3. QUY TẮC NGHIỆP VỤ ĐÃ CHỐT (bắt buộc tuân thủ tuyệt đối — đây là phần quan trọng nhất)

1. **Khách hàng: LUÔN LUÔN = "Khách lẻ"** — bất kể PDF ghi "Bán cho người tiêu dùng" hay
   ghi tên cụ thể (VD "TRAN BAO QUAN"). Không tạo/tìm khách hàng theo tên trên hóa đơn.
2. **Trạng thái thanh toán: LUÔN LUÔN = "Chưa thu tiền"** — dù hóa đơn ghi hình thức
   thanh toán là "Tiền mặt/Chuyển khoản". Không bao giờ chọn "Thu tiền ngay".
3. **Thuế suất VAT — quy tắc đặc biệt, dễ nhầm nhất:**
   - Hóa đơn PDF ghi thuế suất **8%**.
   - Trong MISA, thuế suất được **chọn cố định 10%** (khác với hóa đơn gốc).
   - Nhưng **số tiền thuế GTGT phải được ghi đè bằng đúng số tiền thuế thật trên hóa đơn**
     (tính theo 8%), KHÔNG lấy số MISA tự động tính theo 10%.
   - Nói cách khác: %VAT và số tiền thuế là 2 trường độc lập khi nhập — %VAT theo cấu
     hình mã thuế công ty (10%), số tiền thuế theo thực tế hóa đơn (8%).
4. **Mã hàng: đối chiếu gần đúng (fuzzy match), không khớp tuyệt đối chuỗi ký tự.**
   Danh mục hàng hóa đã có sẵn trong MISA, nhưng tên gọi lệch với tên trên PDF. Ví dụ:
   - PDF: "Trà sữa Hồng Bao (M)" ↔ MISA: "Trà sữa Hồng Bao-Size M"
   - PDF: "Trà sữa Hồng Bao (XL)" ↔ MISA: "Trà sữa Hồng Bao-Size XL"
   Cần thuật toán so khớp gần đúng (đã dùng `rapidfuzz`), có ngưỡng tin cậy: dưới ngưỡng
   → KHÔNG tự chọn, đẩy hóa đơn vào hàng đợi duyệt tay.
5. **Thứ tự xử lý: tuần tự theo Số hóa đơn tăng dần** (911 → 912 → 913 → ...).
6. **Đối chiếu số liệu bắt buộc trước khi lưu** — 3 chỉ tiêu phải khớp tuyệt đối với PDF:
   Tổng tiền hàng, Tiền thuế GTGT (số đã ghi đè), Tổng tiền thanh toán. Nếu không khớp
   (sai số đọc PDF, lỗi tính toán...) → KHÔNG tự động Cất, đẩy vào hàng đợi duyệt tay.
7. **1 file PDF có thể chứa NHIỀU hóa đơn** (nhiều trang) — đã phát hiện thực tế khi test
   với file mẫu. Phải xử lý theo từng trang, không gộp nhầm dữ liệu nhiều hóa đơn.

---

## 4. Bản mockup giao diện đã dựng (tham khảo UX, không phải code chạy thật)

File: `invoice_entry_mockup.html` (đã gửi cho người dùng, KHÔNG có trong repo code RPA).
Mô phỏng 1 giao diện web độc lập (không phải MISA thật) minh họa luồng:
- Bên trái: bản phỏng hóa đơn PDF gốc (nền giấy vàng nhạt, viền đỏ).
- Bên phải: form nhập liệu — Khách hàng khóa cứng "Khách lẻ", trạng thái khóa cứng
  "Chưa thu tiền", bảng hàng hóa hiển thị mã hàng đối chiếu (chip xanh = khớp chắc chắn,
  chip vàng "cần xác nhận" = độ tin cậy thấp), khối thuế thể hiện số tự tính 10% (gạch
  ngang) cạnh số ghi đè theo hóa đơn thật, 3 dòng tổng có nhãn "✓ khớp PDF".
- Thanh hàng đợi phía trên theo dõi tiến độ nhập tuần tự (chờ/đang nhập/đã cất).

Lưu ý: ban đầu định thiết kế trong **Figma** (link file được cung cấp,
fileKey=`wZAsypbbTl3CJI7Xqutl5s`, node `8603:1415` là 1 page trống) nhưng **bị chặn do
giới hạn Figma MCP tool call trên gói Starter** — đã chuyển sang dựng mockup HTML/CSS/JS
thuần thay thế. Nếu sau này nâng cấp gói Figma, có thể quay lại dựng trong file Figma đó.

---

## 5. Kiến trúc kỹ thuật đã chọn cho RPA

**Lý do:** MISA SME.NET 2019 là ứng dụng desktop Windows (.NET WinForms), **KHÔNG có API
công khai** để ghi dữ liệu — không thể gọi API như phần mềm hiện đại. Đã cân nhắc ghi
thẳng vào database SQL Server của MISA nhưng **loại bỏ vì quá rủi ro** (sai cấu trúc bảng,
vi phạm ràng buộc nghiệp vụ nội bộ, mất toàn vẹn dữ liệu kế toán).

**→ Hướng đã chọn: RPA / UI Automation** — mô phỏng thao tác chuột/bàn phím như người
dùng thật.

| Lớp xử lý | Công nghệ | Trạng thái |
|---|---|---|
| Đọc PDF hóa đơn | Python + `pdfplumber` | ✅ Đã viết & TEST THẬT, hoạt động đúng |
| Đối chiếu số liệu chéo | `pandas` (thủ công trong code) | ✅ Đã viết & test |
| Đối chiếu mã hàng (fuzzy) | `rapidfuzz` | ✅ Đã viết & test |
| Điều khiển giao diện MISA | `pywinauto` (backend `uia`) | ⚠️ Đã viết KHUNG code, CHƯA chạy thật — cần dò `auto_id` thật |
| Điều phối tổng thể | `main.py` (argparse CLI) | ✅ Đã viết |

---

## 6. Code đã có sẵn trong thư mục dự án

```
misa_rpa/
├── pdf_extractor.py      # Đọc PDF theo nhãn trường (regex), đối chiếu chéo số liệu
├── product_matcher.py    # Fuzzy match Tên hàng PDF <-> Mã hàng MISA (rapidfuzz)
├── misa_automation.py    # pywinauto — điều khiển popup MISA (khung code, cần hoàn thiện)
├── main.py                # CLI điều phối: đọc PDF -> đối chiếu -> nhập MISA, --dry-run
├── requirements.txt
├── README.md
└── test_pdfs/
    └── hoa_don_test.pdf   # File PDF mẫu thật (3 hóa đơn: 911, 912, 913)
```

### `pdf_extractor.py`
- `extract_invoices(pdf_path) -> List[Invoice]`: trích TẤT CẢ hóa đơn trong file (xử lý
  theo từng trang, lọc trang không phải hóa đơn).
- `extract_invoice(pdf_path) -> Invoice`: tiện ích cho file chỉ có 1 hóa đơn (raise lỗi
  nếu file có nhiều hóa đơn).
- Trích theo NHÃN TRƯỜNG bằng regex (không theo tọa độ cố định) — ví dụ tìm
  `"Số *(No\.?)*[:\.]?\s*(\d+)"` để lấy số hóa đơn.
- **`_cross_validate()`**: tự đối chiếu SL×Đơn giá=Thành tiền, Tổng dòng=Tổng khai báo,
  Tiền hàng+Thuế=Tổng thanh toán → gắn cảnh báo vào `Invoice.warnings` nếu lệch.
- Đã sửa 2 lỗi phát hiện khi test với file thật:
  1. 1 PDF nhiều trang = nhiều hóa đơn (không gộp chung).
  2. Lọc dòng rác trong bảng (dòng tiêu đề công thức "1 2 3 4 5 6=4x5..." bị đọc lẫn vào
     dữ liệu — loại bằng điều kiện `ten.isdigit()` hoặc `len(ten) < 2`).

### `product_matcher.py`
- `ProductMatcher(catalog_csv_path)`: nạp danh mục hàng hóa MISA (CSV/Excel, cột bắt
  buộc `MaHang`, `TenHang`).
- `.match(ten_hang_pdf) -> MatchResult`: chuẩn hoá chuỗi (bỏ hoa/thường, dấu ngoặc, chữ
  "size"...) rồi fuzzy match bằng `rapidfuzz.process.extractOne` + `fuzz.WRatio`.
- Ngưỡng: `CONFIDENCE_AUTO_ACCEPT = 90` (tự động chọn), `CONFIDENCE_MIN_SUGGEST = 60`
  (dưới mức này không gợi ý gì, để trống bắt buộc chọn tay).
- `build_sample_matcher()`: danh mục mẫu dựng sẵn để test khi CHƯA có file catalog thật
  từ MISA (6 mã: TP0021, TP0022, TP0031, TP0045, TP0046, TP0052).

### `misa_automation.py`
- Class `MisaAutomation`: `connect()`, `open_new_invoice_popup()`, `fill_invoice()`,
  `save()`, `close_popup_if_open()`.
- Class `Controls`: **TOÀN BỘ giá trị `auto_id` trong này là VÍ DỤ MINH HỌA, CHƯA XÁC
  THỰC** — bắt buộc phải dò lại trên máy chạy MISA thật.
- `LowConfidenceMatchError`: raise khi mã hàng có độ tin cậy thấp — dừng lại, không tự
  đoán, không tự Cất.
- Hỗ trợ `dry_run=True` để log ra các bước SẼ làm mà không thao tác thật lên MISA.
- **Hướng dẫn dò control (ở cuối file, giữ nguyên khi đọc code):**
  1. `pip install pywinauto`
  2. Mở MISA, mở đúng popup "Chứng từ bán hàng".
  3. Chạy:
     ```python
     from pywinauto import Application
     app = Application(backend="uia").connect(title_re="Bán hàng hóa, dịch vụ.*")
     win = app.window(title_re="Bán hàng hóa, dịch vụ.*")
     win.print_control_identifiers(depth=6)
     ```
  4. Nếu bảng "Hàng tiền" là lưới custom (DevExpress/Infragistics — phổ biến ở app .NET
     WinForms cũ) và không lộ automation ID từng ô, dùng thêm **Accessibility Insights
     for Windows** (Microsoft, miễn phí) để hover từng ô xem AutomationId/ControlType.
  5. Nếu backend `"uia"` không thấy control, thử `Application(backend="win32")`.
  6. LUÔN test trên dữ liệu demo trước, không chạy thẳng dữ liệu công ty thật.

### `main.py`
- CLI: `python main.py --pdf-dir ./hoa_don --catalog ./danh_muc.csv [--dry-run]`
- `load_and_sort_invoices()`: đọc toàn bộ PDF trong thư mục (dùng `extract_invoices`,
  hỗ trợ multi-invoice-per-file), sort theo số hóa đơn tăng dần.
- `to_entry_payload()`: chuyển `Invoice` → `InvoiceToEnter` + chạy fuzzy match từng dòng.
- `run()`: vòng lặp chính — hóa đơn có `warnings` (lệch đối chiếu số liệu) → đẩy vào
  `can_duyet_tay`, KHÔNG tự động nhập; hóa đơn có `LowConfidenceMatchError` → tương tự;
  còn lại → mở popup, điền, Cất, log kết quả vào `results["da_cat"]`.

---

## 7. Kết quả TEST THẬT đã chạy (bằng chứng độ chính xác của lớp đọc PDF)

Đã cài đặt: `pip install pdfplumber rapidfuzz pandas pydantic --break-system-packages`

Chạy `extract_invoices()` trên file PDF thật (`hoa_don_test.pdf`, chứa 3 hóa đơn 3 trang):

```
Tìm thấy 3 hóa đơn trong file

--- Hóa đơn số 913 ---
Ngày: 28/08/2026 | Người mua: Bán cho người tiêu dùng
   TRÀ BÍ ĐAO SƯƠNG SÁO (XL) | Ly | SL=2.0 | ĐG=17,593 | TT=35,185
   TRÂN CHÂU 3Q | Phần | SL=1.0 | ĐG=5,556 | TT=5,556
Tổng trước thuế: 40,741 | Tổng thuế: 3,259 | Tổng TT: 44,000
Cảnh báo: Không có — khớp số liệu ✓

--- Hóa đơn số 912 ---
Ngày: 28/08/2026 | Người mua: TRAN BAO QUAN
   Trà sữa Hồng Bao (M) | Ly | SL=1.0 | ĐG=8,333 | TT=8,333
   KEM TRỨNG | Phần | SL=1.0 | ĐG=8,333 | TT=8,333
Tổng trước thuế: 16,666 | Tổng thuế: 1,334 | Tổng TT: 18,000
Cảnh báo: Không có — khớp số liệu ✓

--- Hóa đơn số 911 ---
Ngày: 28/08/2026 | Người mua: Bán cho người tiêu dùng
   TRÂN CHÂU ĐEN | Phần | SL=1.0 | ĐG=3,889 | TT=3,889
   Trà sữa Hồng Bao (XL) | Ly | SL=1.0 | ĐG=15,741 | TT=15,741
Tổng trước thuế: 19,630 | Tổng thuế: 1,570 | Tổng TT: 21,200
Cảnh báo: Không có — khớp số liệu ✓
```

Chạy `ProductMatcher` (danh mục mẫu) trên 6 dòng hàng của 3 hóa đơn trên: **cả 6 dòng
khớp `score=100`, `AUTO ✓`** — ví dụ `"TRÀ BÍ ĐAO SƯƠNG SÁO (XL)" -> TP0031 (Trà bí đao
sương sáo-Size XL)`.

**Trước khi sửa** (phiên bản đầu, lỗi): gộp cả 3 trang thành 1 "hóa đơn", đọc lẫn dòng
tiêu đề công thức bảng ("2","3","4","5" là rác từ dòng `1 2 3 4 5 6=4x5...`) làm dòng
hàng giả, và **lớp đối chiếu chéo đã tự phát hiện đúng lỗi này** (cảnh báo "Tổng các
dòng hàng (77,037) không khớp Tổng tiền trước thuế (40,741)") — minh chứng cơ chế
validate hoạt động đúng như thiết kế, dù dữ liệu đầu vào khi đó có lỗi.

---

## 8. VIỆC CÒN DANG DỞ — cần Claude Code tiếp tục

1. **Dò `auto_id` thật của từng control trong popup MISA** (ưu tiên cao nhất) — cần truy
   cập máy đang chạy MISA thật (qua UltraViewer hoặc trực tiếp), chạy
   `print_control_identifiers()` hoặc dùng Accessibility Insights for Windows, rồi cập
   nhật class `Controls` trong `misa_automation.py`.
2. **Xử lý bảng "Hàng tiền" (grid)** — hàm `_grid_set_cell()` hiện là khung mẫu
   (double-click + gõ + Tab), cần kiểm chứng thực tế có hoạt động với loại lưới MISA
   dùng hay không (nghi ngờ là lưới custom DevExpress/Infragistics, có thể cần cách khác
   như điều hướng bằng phím Tab/mũi tên tuần tự thay vì auto_id từng ô).
3. **Danh mục hàng hóa thật** — hiện đang dùng `build_sample_matcher()` (6 mã mẫu), cần
   người dùng xuất danh mục thật từ MISA ra CSV/Excel (cột `MaHang`, `TenHang`) để thay
   thế.
4. **Test toàn bộ luồng end-to-end với `--dry-run`** trước, sau đó thử trên 1 hóa đơn
   thật/dữ liệu demo trước khi chạy hàng loạt vào dữ liệu công ty thật.
5. Cân nhắc bổ sung: cơ chế tránh nhập trùng hóa đơn (lưu log số hóa đơn đã xử lý), xử lý
   ngoại lệ khi MISA hiện popup cảnh báo/lỗi ngoài dự kiến giữa chừng.

---

## 9. Ghi chú khác từ cuộc trò chuyện

- Người dùng thao tác MISA qua **UltraViewer 6.6 Free** — lưu ý bản Free giới hạn thời
  gian phiên kết nối liên tục, cần tính đến nếu chạy RPA kéo dài qua remote desktop.
- Claude Code và claude.ai (nơi diễn ra cuộc trò chuyện gốc) **không tự chia sẻ ngữ cảnh
  với nhau** — đây là lý do file `CLAUDE.md` này được tạo ra, để nạp thủ công toàn bộ bối
  cảnh vào Claude Code.
- Claude Code có khả năng đọc ảnh (JPEG/PNG/GIF/WebP) qua dán/kéo-thả/đường dẫn file —
  hữu ích khi cần đưa ảnh chụp kết quả dò control hoặc ảnh lỗi giao diện MISA vào để nhờ
  debug tiếp.

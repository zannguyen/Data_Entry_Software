"""
misa_automation.py
--------------------
Điều khiển giao diện MISA SME.NET 2019 (ứng dụng Windows Desktop, không có API)
bằng pywinauto — mô phỏng thao tác con người: mở popup, điền ô, bấm Cất.

Yêu cầu: pip install pywinauto

*** QUAN TRỌNG — ĐỌC TRƯỚC KHI DÙNG ***
Các chuỗi trong `CONTROLS` bên dưới (title, auto_id, class_name...) là VÍ DỤ MINH HỌA
dựa theo nhãn nhìn thấy trên ảnh giao diện bạn cung cấp. pywinauto cần định danh CHÍNH XÁC
của từng control (auto_id hoặc control_type nội bộ), mà chỉ có thể lấy được bằng cách
DÒ TRỰC TIẾP trên máy đang chạy MISA thật — xem hướng dẫn ở cuối file (phần "DÒ CONTROL").
Chạy thử với --dry-run trước, và luôn test trên 1 hóa đơn / dữ liệu demo trước khi chạy hàng loạt.
"""

import time
import logging
import unicodedata
from dataclasses import dataclass
from typing import Optional

from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import TimeoutError as PywinautoTimeout
from pywinauto.keyboard import send_keys
from rapidfuzz import fuzz

from pdf_extractor import Invoice
from product_matcher import MatchResult, _normalize

# Ngưỡng tin cậy RIÊNG cho bước tra cứu mã hàng TRỰC TIẾP trong MISA lúc nhập thật
# (khác CONFIDENCE_AUTO_ACCEPT của product_matcher.py, vốn dùng cho file catalog cũ).
# Đặt cao vì đây là bước ghi thật vào sổ sách kế toán — thà đẩy nhiều hóa đơn hơn vào
# hàng đợi duyệt tay còn hơn chọn nhầm hàng. Việc chấm điểm dùng hàm
# MisaAutomation._name_match_score() (ưu tiên khớp TIỀN TỐ tuyệt đối trước, vì danh
# mục MISA của công ty đặt tên đầy đủ hơn PDF — VD PDF "Miến Xào Cua" ↔ MISA "Miến Xào
# Cua Lột-2026", đã xác nhận ĐÚNG sản phẩm qua đối chiếu Đơn giá trong video demo người
# dùng cung cấp 30/08/2026 — chỉ dùng fuzz.ratio làm phương án dự phòng khi không phải
# quan hệ tiền tố, để vẫn bắt được các trường hợp thật sự khác sản phẩm).
LIVE_MATCH_CONFIDENCE_THRESHOLD = 92

# Bí danh THỦ CÔNG cho các trường hợp tên PDF và tên MISA khác nhau HOÀN TOÀN về từ
# ngữ (không phải lệch chính tả/viết tắt mà thuật toán tự suy luận được ở
# _name_match_score) — VD "Nước suối" (PDF, tên gọi chung) <-> "Nước tinh khiết
# Aquafina 500ml" (MISA, đặt theo đúng nhãn hiệu công ty đang bán) — XÁC NHẬN THẬT
# (30/08/2026, người dùng chỉ định trực tiếp). Bổ sung dần key mới (chuỗi con, đã qua
# _normalize) khi gặp thêm trường hợp tương tự trong thực tế. Dùng để THAY THẾ nội
# dung gõ vào ô tìm kiếm Mã hàng ở Bước 1 (không đụng gì đến việc chấm điểm ở Bước 2).
KNOWN_PRODUCT_ALIASES = {
    "nước suối": "aquafina",
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("misa_rpa")


# =========================================================================
# CẤU HÌNH — cần chỉnh lại theo kết quả dò control thật trên máy của bạn
# =========================================================================
class Controls:
    """
    Toàn bộ auto_id dưới đây đã DÒ THẬT trên máy chạy MISA SME.NET 2019 R7 Enterprise
    (đơn vị CÔNG TY TNHH GUABAO ONE) bằng pywinauto (backend uia), chạy dưới quyền
    Administrator (bắt buộc — MISA chạy elevated nên cần quyền tương đương để UI
    Automation đọc/ghi được, do cơ chế UIPI của Windows).

    Lưu ý kiến trúc quan trọng phát hiện khi dò:
    - MISA dùng MDI nội bộ: popup "Bán hàng hóa, dịch vụ..." KHÔNG phải cửa sổ Windows
      riêng (app.windows() không thấy) mà là 1 Window nhúng bên trong frmMain
      (auto_id="frmSAVoucherDetail"). Phải tìm qua main_win.child_window(...), không
      qua Application.connect(title_re=...) riêng.
    - Bảng "Hàng tiền" (grdDetail) là grid custom kiểu Janus GridEX — mỗi dòng là
      DataItem nhưng auto_id="-1" GIỐNG NHAU cho mọi dòng -> phải đánh địa chỉ theo
      CHỈ SỐ (index), không theo auto_id.
    - % thuế GTGT và Tiền thuế GTGT theo dòng nằm NGAY TRONG grdDetail (hiện thêm cột
      khi tab "2. Thuế" active), KHÔNG phải ô rời ngoài bảng.
    """

    MAIN_WINDOW_AUTO_ID = "frmMain"
    POPUP_DETAIL_AUTO_ID = "frmSAVoucherDetail"  # form nhập/sửa chứng từ, nhúng trong frmMain

    # ---- Điều hướng vào màn hình "Bán hàng" (danh sách chứng từ) ----
    # ĐÃ DÒ THẬT: không phải menu trên cùng (Tệp/Danh mục/Nghiệp vụ...) mà là thanh điều
    # hướng bên trái kiểu Outlook navigation pane, auto_id="ebMain". Cấu trúc: Group (module)
    # -> DataItem (mục con). "[Group] N" / "[Item] N" là CHỈ SỐ TƯƠNG ĐỐI, lặp lại ở nhiều
    # nhóm khác nhau -> PHẢI tìm theo window_text(), không dùng auto_id cố định.
    NAV_PANE_AUTO_ID = "ebMain"
    NAV_ITEM_CHUNG_TU_BAN_HANG_TEXT = "Chứng từ bán hàng"
    # Khi đã mở, danh sách "Bán hàng" xuất hiện dưới dạng tab MDI với auto_id cố định này
    # -> dùng để kiểm tra đã mở sẵn chưa, tránh mở trùng nhiều tab.
    TAB_BAN_HANG_LIST_AUTO_ID = "TabItem Key SAVoucher"

    TOOLBAR_BTN_THEM_AUTO_ID = "[Toolbar : ListToolbar Tools] Tool : mnuAdd - Index : 0 "
    TOOLBAR_BTN_XEM_AUTO_ID = "[Toolbar : ListToolbar Tools] Tool : mnuView - Index : 1 "
    TOOLBAR_BTN_XOA_AUTO_ID = "[Toolbar : ListToolbar Tools] Tool : mnuDelete - Index : 2 "
    TOOLBAR_BTN_GHI_SO_AUTO_ID = "[Toolbar : ListToolbar Tools] Tool : mnuPost - Index : 3 "
    # Nút "Cất" trong popup detail — ĐÃ DÒ + XÁC NHẬN CẤT THẬT THÀNH CÔNG (30/08/2026).
    POPUP_BTN_CAT_AUTO_ID = "[Toolbar : MainToolbar Tools] Tool : mnuSave - Index : 6 "
    # Nút "Đóng" trong popup detail — dùng để đóng tab hóa đơn vừa Cất, quay lại danh
    # sách, tránh tích tụ nhiều tab mở khi chạy hàng loạt nhiều hóa đơn liên tiếp.
    POPUP_BTN_DONG_AUTO_ID = "[Toolbar : MainToolbar Tools] Tool : mnuClose - Index : 19 "

    # ---- Trạng thái thanh toán: radio ẩn trong combobox optPaymentMethod ----
    FIELD_PAYMENT_METHOD_COMBO = "optPaymentMethod"
    RADIO_CHUA_THU_TIEN_AUTO_ID = "[Editor] [valuelist] ValueListItem 0"   # "Chưa thu tiền"
    RADIO_THU_TIEN_NGAY_AUTO_ID = "[Editor] [valuelist] ValueListItem 1"  # "Thu tiền ngay"

    # ---- Khách hàng & thông tin chung ----
    FIELD_KHACH_HANG = "cboAccountObjectID"       # combobox chọn khách hàng
    FIELD_TEN_KHACH_HANG = "txtAccountObjectName"
    FIELD_DIA_CHI = "txtAccountObjectAddress"
    FIELD_MA_SO_THUE = "txtAccountObjectTaxCode"
    FIELD_NGUOI_LIEN_HE = "txtPayer"
    FIELD_NV_BAN_HANG = "cboEmployeeID"
    FIELD_DIEU_KHOAN_TT = "cboPaymentTerm"
    FIELD_SO_NGAY_NO = "numDueDay"
    FIELD_HAN_THANH_TOAN = "dteDueDate"
    FIELD_NGAY_HACH_TOAN = "dtePostedDate"
    FIELD_NGAY_CHUNG_TU = "dteRefDate"
    FIELD_THAM_CHIEU = "txtRefNo"
    FIELD_DIEN_GIAI = "txtJournalMemo"
    FIELD_HINH_THUC_TT = "cboPayType"             # Tiền mặt / Chuyển khoản (không dùng theo rule #2)

    # ---- Các tab trong form ----
    TAB_HANG_TIEN_AUTO_ID = "TabItem Index : 0"
    TAB_THUE_AUTO_ID = "TabItem Index : 1"
    TAB_GIA_VON_AUTO_ID = "TabItem Index : 2"
    TAB_THONG_KE_AUTO_ID = "TabItem Index : 3"
    TAB_KHAC_AUTO_ID = "TabItem Key Other"
    # Tab nhóm TRÊN CÙNG (khác nhóm tab Hàng tiền/Thuế ở dưới) — "Chứng từ ghi nợ" /
    # "Phiếu xuất" / "Hóa đơn". ĐÃ DÒ THẬT (30/08/2026, theo yêu cầu người dùng): cần
    # qua tab "Hóa đơn" để kiểm tra/sửa "Số hóa đơn" (txtInvNo) khớp với số hóa đơn
    # thật trên PDF trước khi Cất — nếu không sửa, MISA có thể lưu sai số tham chiếu
    # hóa đơn gốc, ảnh hưởng tính hợp lệ chứng từ.
    TAB_HOA_DON_AUTO_ID = "TabItem Key Invoice"
    FIELD_SO_HOA_DON_THAT = "txtInvNo"  # "Số hóa đơn" — phải khớp số hóa đơn gốc trên PDF

    # ---- Bảng "Hàng tiền" (grid custom, kiểu Janus GridEX) ----
    GRID_HANG_TIEN_AUTO_ID = "grdDetail"
    # QUAN TRỌNG: đây phải là TÊN HIỂN THỊ TIẾNG VIỆT (khớp window_text() của từng ô
    # trong grid, dùng bởi _grid_cell()) — KHÔNG PHẢI tên field nội bộ tiếng Anh (VD
    # "InventoryItemCode") vốn chỉ dùng cho auto_id="[Column Header] <field>" của HEADER
    # cột, không phải của từng Ô dữ liệu. Bug thật đã xảy ra (29/08/2026): dùng nhầm tên
    # field nội bộ khiến _grid_cell() không tìm thấy ô nào, crash ngay dòng đầu tiên
    # trước khi gõ được bất kỳ ký tự nào vào grid -> hóa đơn RỖNG bị lưu nhầm (kết hợp
    # với bug close_popup_if_open cũ, đã sửa riêng ở trên).
    GRID_COL_MA_HANG = "Mã hàng"
    GRID_COL_TEN_HANG = "Tên hàng"
    GRID_COL_TK_NO = "TK công nợ/chi phí"
    GRID_COL_TK_DOANH_THU = "TK doanh thu"
    GRID_COL_DVT = "ĐVT"
    GRID_COL_SO_LUONG = "Số lượng"
    GRID_COL_DON_GIA = "Đơn giá"
    GRID_COL_THANH_TIEN = "Thành tiền"
    GRID_COL_TY_LE_CK = "Tỷ lệ CK (%)"
    GRID_COL_TIEN_CK = "Tiền chiết khấu"
    GRID_COL_TK_CHIET_KHAU = "TK chiết khấu"
    # Cột thuế — chỉ hiện trong cây UIA khi tab "2. Thuế" đang active (lazy-load)
    GRID_COL_THUE_SUAT_PCT = "% thuế GTGT"          # combobox trong ô: 0% / 5% / 10%
    GRID_COL_TIEN_THUE_DONG = "Tiền thuế GTGT"       # tiền thuế của riêng dòng này
    GRID_COL_TK_THUE = "TK thuế GTGT"
    GRID_COL_DIEN_GIAI_THUE = "Diễn giải thuế"
    # Giá trị %VAT khả dụng trong dropdown ô — theo quy tắc #3, LUÔN chọn 10%
    THUE_SUAT_LISTITEM_10PCT_AUTO_ID = "[Editor] [valuelist] ValueListItem 2"  # "10%"

    # ---- Khối tổng cuối form (đây mới là nơi ghi đè Tiền thuế GTGT theo quy tắc #3) ----
    FIELD_TONG_TIEN_HANG = "creTotalAmountOC"
    FIELD_TONG_CHIET_KHAU = "creTotalDiscountAmountOC"
    FIELD_TONG_TIEN_THUE = "creTotalVATAmountOC"     # <-- GHI ĐÈ bằng số thuế thật (8%) ở đây
    FIELD_TONG_THANH_TOAN = "creTotalTotalSaleAmountOC"


def _format_vn_number(value: float) -> str:
    """
    Định dạng số theo chuẩn Việt Nam trước khi gõ vào MISA: dấu CHẤM ngăn cách hàng
    nghìn, dấu PHẨY cho phần thập phân (ngược định dạng Mỹ mặc định của Python).
    LÝ DO: xác nhận lỗi thật (29/08/2026) — gõ str(1.0) = "1.0" vào ô Số lượng, MISA
    hiểu thành "10" (dấu chấm kiểu Mỹ bị bỏ qua, "1" và "0" bị ghép liền). Số nguyên
    (VD số lượng =1) trả về chuỗi thuần không dấu phân cách; số lớn (đơn giá, tiền
    thuế) dùng dấu chấm hàng nghìn khớp đúng định dạng hiển thị trên hóa đơn PDF gốc
    (VD "180.000", "14.400").
    """
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value)):,}".replace(",", ".")
    s = f"{value:,.2f}"
    int_part, dec_part = s.split(".")
    return f"{int_part.replace(',', '.')},{dec_part}"


def _raw_click(x: int, y: int):
    """
    Click chuột THẲNG qua Windows API (SetCursorPos + mouse_event), KHÔNG dùng
    pywinauto's click_input(). XÁC NHẬN THẬT (30/08/2026): dù rectangle() đọc được của
    ô "Mã hàng" hoàn toàn chính xác (đã tự kiểm chứng qua dump control), và dù đã bật
    đúng DPI-awareness cho tiến trình, click_input() của pywinauto (gọi row_el.click_
    input() hoặc cell.click_input()) VẪN nhiều lần bấm trúng nhầm cột "TK chiết khấu"
    (cột cuối cùng bên phải) một cách nhất quán, lặp lại nhiều lần dù đã thử nhiều cách
    khác nhau (click ô, click dòng+Home, focus lại cửa sổ...). Nghi ngờ pywinauto tự
    làm 1 bước quy đổi/tính toán tọa độ nội bộ nào đó (không rõ cơ chế chính xác) gây
    sai lệch — nên bỏ hẳn lớp trung gian này, gọi thẳng API chuột của Windows bằng toạ
    độ tự tính từ rectangle() (đã xác nhận rectangle() luôn đúng), loại trừ hoàn toàn
    khả năng pywinauto tính sai.
    """
    import ctypes

    ctypes.windll.user32.SetCursorPos(int(x), int(y))
    time.sleep(0.1)
    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.1)


def _click_element_center(element):
    """Click vào chính giữa rectangle() thật của 1 control, dùng _raw_click()."""
    rect = element.rectangle()
    cx = (rect.left + rect.right) // 2
    cy = (rect.top + rect.bottom) // 2
    _raw_click(cx, cy)


def _double_click_element_center(element):
    """
    Double-click (2 click liên tiếp nhanh) vào giữa rectangle() thật của 1 control.
    XÁC NHẬN THẬT (30/08/2026): click ĐƠN vào ô grid chỉ CHỌN ô (chưa vào chế độ edit
    text thật) — ở trạng thái đó, gửi phím Home/End bị GRID hiểu thành lệnh điều hướng
    "nhảy tới cột đầu/cuối" thay vì di chuyển con trỏ trong text (đã xác nhận qua quan
    sát trực tiếp: click đúng "Mã hàng" nhưng gửi Home rồi End đều làm "nhảy" sang "TK
    chiết khấu" — cột cuối). Double-click mới thật sự vào chế độ edit text.
    """
    rect = element.rectangle()
    cx = (rect.left + rect.right) // 2
    cy = (rect.top + rect.bottom) // 2
    _raw_click(cx, cy)
    time.sleep(0.08)
    _raw_click(cx, cy)


KHACH_LE_NAME = "Khách lẻ"
FIXED_VAT_PERCENT_LABEL = "10%"


@dataclass
class LineToEnter:
    ma_hang: str
    so_luong: float
    don_gia: float
    # Tiền thuế GTGT THẬT của riêng dòng này (lấy từ PDF, đã tính theo thuế suất
    # thật của dòng — có thể là 8% hoặc 10% tuỳ dòng). LUÔN ghi đè vào MISA bằng
    # đúng số này, bất kể dòng gốc là 8% hay 10% — vì PDF đã tính đúng theo thuế
    # suất thật của dòng đó rồi, ghi đè y nguyên là đủ, không cần phân biệt %.
    tien_thue_dong: float = 0.0


@dataclass
class InvoiceToEnter:
    ngay: str                  # dd/mm/yyyy
    lines: list                # List[LineToEnter]
    tien_thue_that: float      # tổng tiền thuế THẬT lấy từ PDF (= tổng tien_thue_dong các dòng)
    tong_tien_hang: float
    tong_thanh_toan: float
    so_hoa_don_goc: str        # chỉ để log / đối chiếu, không nhập vào MISA


class MisaAutomation:
    def __init__(self, dry_run: bool = True):
        self.dry_run = dry_run
        self.app: Optional[Application] = None
        self.main_win = None

    # --------------------------------------------------------------- #
    @staticmethod
    def _find_main_window_handle() -> int:
        """
        Dò handle cửa sổ chính MISA bằng win32 EnumWindows thô (ctypes) — KHÔNG dùng
        Application.connect(title_re=...) vì trên thực tế thường có NHIỀU cửa sổ cùng
        khớp tiêu đề "MISA SME.NET 2019..." (VD cửa sổ chính + toast thông báo/popup
        phụ nhỏ) khiến pywinauto báo lỗi "There are N elements that match the criteria".
        Cách chắc chắn: liệt kê toàn bộ cửa sổ top-level, lọc theo tiêu đề, rồi chọn
        cửa sổ có DIỆN TÍCH LỚN NHẤT (cửa sổ chính luôn to hơn hẳn toast/popup phụ).
        """
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        results = []

        def foreach_window(hwnd, _lparam):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0 and user32.IsWindowVisible(hwnd):
                buff = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buff, length + 1)
                if "MISA SME.NET 2019" in buff.value:
                    rect = wintypes.RECT()
                    user32.GetWindowRect(hwnd, ctypes.byref(rect))
                    area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
                    results.append((area, hwnd, buff.value))
            return True

        user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(foreach_window), 0)

        if not results:
            raise ElementNotFoundError(
                "Không tìm thấy cửa sổ nào có tiêu đề chứa 'MISA SME.NET 2019' — MISA đã mở chưa?"
            )
        results.sort(key=lambda t: t[0], reverse=True)
        area, hwnd, title = results[0]
        log.info("Tìm thấy %d cửa sổ khớp tiêu đề MISA, chọn cửa sổ lớn nhất: '%s' (area=%d)", len(results), title, area)
        return hwnd

    def connect(self):
        """
        Kết nối vào cửa sổ MISA đang MỞ SẴN (không tự khởi chạy MISA).
        QUAN TRỌNG: MISA chạy quyền Administrator -> script này CŨNG PHẢI chạy dưới
        quyền Administrator, nếu không Windows (UIPI) sẽ chặn đọc/ghi control (chỉ
        thấy được lớp ngoài cùng Dialog/TitleBar, xác nhận thực tế khi dò control).
        Đã xác nhận connect bằng handle (qua EnumWindows thô + Application.connect(handle=...))
        hoạt động ổn định; title_re trực tiếp bị lỗi khi có >1 cửa sổ khớp tiêu đề.
        """
        log.info("Đang kết nối vào cửa sổ MISA SME.NET 2019 ...")
        hwnd = self._find_main_window_handle()
        self.app = Application(backend="uia").connect(handle=hwnd)
        self.main_win = self.app.window(handle=hwnd)
        self.main_win.set_focus()
        log.info("Đã kết nối: %s", self.main_win.window_text())

    # --------------------------------------------------------------- #
    def open_new_invoice_popup(self):
        """
        Bán hàng ▸ Chứng từ bán hàng ▸ Thêm.
        LƯU Ý KIẾN TRÚC: popup không phải cửa sổ Windows riêng — MISA dùng MDI nội bộ,
        form nhập/sửa chứng từ NHÚNG bên trong frmMain với auto_id="frmSAVoucherDetail".
        Giả định người dùng đã điều hướng sẵn tới đúng danh sách "Bán hàng" trước khi
        gọi hàm này (chưa tự động hoá bước vào menu Bán hàng, vì auto_id của mục menu
        con "Chứng từ bán hàng" chưa được dò).
        """
        log.info("Bấm nút 'Thêm' trên danh sách Bán hàng ...")
        if self.dry_run:
            log.info("[DRY-RUN] (bỏ qua click nút Thêm thật)")
            return None

        _click_element_center(self.main_win.child_window(
            auto_id=Controls.TOOLBAR_BTN_THEM_AUTO_ID, control_type="Button"
        ))

        popup = self.main_win.child_window(
            auto_id=Controls.POPUP_DETAIL_AUTO_ID, control_type="Window"
        )
        popup.wait("visible", timeout=10)
        return popup

    # --------------------------------------------------------------- #
    def ensure_ban_hang_list_open(self):
        """
        Điều hướng vào danh sách "Bán hàng" qua thanh nav trái (ebMain) nếu chưa mở sẵn.
        CHƯA TEST THẬT (mới xác nhận cấu trúc qua đọc cây, chưa từng click mục nav này —
        lần dò control trước đó, danh sách "Bán hàng" đã được người dùng tự mở sẵn tay).
        """
        if self.dry_run:
            log.info("[DRY-RUN] Sẽ điều hướng vào 'Bán hàng ▸ Chứng từ bán hàng' nếu chưa mở.")
            return

        # Đã mở sẵn tab "Bán hàng" (frmSAVoucher) thì không cần click lại.
        try:
            tab = self.main_win.child_window(
                auto_id=Controls.TAB_BAN_HANG_LIST_AUTO_ID, control_type="TabItem"
            )
            if tab.exists():
                _click_element_center(tab)
                return
        except Exception:
            pass

        nav = self.main_win.child_window(auto_id=Controls.NAV_PANE_AUTO_ID, control_type="Custom")
        for item in nav.descendants(control_type="DataItem"):
            if item.window_text() == Controls.NAV_ITEM_CHUNG_TU_BAN_HANG_TEXT:
                _click_element_center(item)
                return
        raise LookupError(
            f"Không tìm thấy mục nav '{Controls.NAV_ITEM_CHUNG_TU_BAN_HANG_TEXT}' — "
            "có thể MISA đổi giao diện, cần dò lại."
        )

    # --------------------------------------------------------------- #
    def fill_invoice(self, popup, inv: InvoiceToEnter, matches: list):
        """
        Điền dữ liệu vào popup theo đúng quy tắc nghiệp vụ đã thống nhất, dùng cơ chế
        THẬT đã xác nhận qua test trực tiếp trên MISA ngày 29/08/2026:
          - Mọi ô nhập liệu (kể cả grid) phải dùng pywinauto.keyboard.send_keys() (giả
            lập bàn phím thật) — set_edit_text() (qua UIA ValuePattern) CÓ set được giá
            trị nhưng KHÔNG kích hoạt được các sự kiện nội bộ của MISA (search/filter,
            tính lại Thành tiền...). Đọc lại giá trị PHẢI qua legacy_properties()['Value']
            — window_text() trả về rỗng với các control này (không lộ qua UIA Value).
          - Cột "Mã hàng" trong grid là 1 ô TÌM KIẾM: gõ TÊN HÀNG (không phải mã) rồi
            Tab -> MISA tự tra cứu danh mục và điền ngược Mã hàng/Tên hàng/TK/ĐVT. Vì
            kết quả gợi ý có thể GẦN ĐÚNG chứ không hoàn toàn chính xác (đã gặp thực tế:
            gõ "Miến Xào Cua" ra kết quả "Miến Xào Cua Lột-2026"), sau khi Tab phải ĐỌC
            LẠI Tên hàng MISA trả về và so khớp gần đúng (rapidfuzz) với tên gốc trên
            PDF — dưới ngưỡng tin cậy thì DỪNG lại, không tự Cất (không cần file danh
            mục CSV rời nữa — tra cứu trực tiếp trong MISA).
        `matches` được giữ tham số để tương thích chữ ký hàm cũ nhưng KHÔNG còn dùng để
        quyết định mã hàng — giờ quyết định bằng kết quả tra cứu trực tiếp trong MISA.
        """
        if self.dry_run:
            log.info("[DRY-RUN] Sẽ điền hóa đơn nguồn số %s như sau:", inv.so_hoa_don_goc)
            log.info("  Khách hàng = %s (cố định)", KHACH_LE_NAME)
            log.info("  Trạng thái = Chưa thu tiền (cố định)")
            log.info("  Ngày = %s", inv.ngay)
            for line in inv.lines:
                log.info(
                    "  Dòng hàng: gõ tên '%s' vào Mã hàng (MISA tự tra cứu) | SL=%s DG=%s | "
                    "%%VAT=%s (cố định) | TienThue ghi đè=%s",
                    line.ma_hang, line.so_luong, line.don_gia,
                    FIXED_VAT_PERCENT_LABEL, f"{line.tien_thue_dong:,.0f}",
                )
            log.info(
                "  Tổng tiền thuế = %s | Tổng tiền hàng = %s | Tổng thanh toán = %s",
                f"{inv.tien_thue_that:,.0f}", f"{inv.tong_tien_hang:,.0f}", f"{inv.tong_thanh_toan:,.0f}",
            )
            return True

        # Ép cửa sổ MISA giữ focus trước khi bắt đầu gõ — giảm rủi ro tổ hợp phím
        # (VD Ctrl+A) bị "kẹt"/lệch sang ứng dụng khác giữa chừng (đã gặp thật:
        # Task View của Windows tự bật lên, làm phím gõ bay lạc, đọc ra Tên hàng rỗng).
        self.main_win.set_focus()
        time.sleep(0.2)

        # ---- Trạng thái: Chưa thu tiền ----
        payment_combo = popup.child_window(auto_id=Controls.FIELD_PAYMENT_METHOD_COMBO, control_type="ComboBox")
        _click_element_center(payment_combo.child_window(
            auto_id=Controls.RADIO_CHUA_THU_TIEN_AUTO_ID, control_type="RadioButton"
        ))

        # ---- Khách hàng: luôn Khách lẻ ----
        combo = popup.child_window(auto_id=Controls.FIELD_KHACH_HANG, control_type="ComboBox")
        _click_element_center(combo)
        time.sleep(0.3)
        send_keys(KHACH_LE_NAME, pause=0.03, with_spaces=True)
        time.sleep(0.5)
        send_keys("{TAB}")
        time.sleep(0.5)

        # ---- Ngày hạch toán ----
        # XÁC NHẬN THẬT (người dùng, 29/08/2026): chỉ cần nhập "Ngày hạch toán" rồi Tab
        # -> MISA TỰ ĐỘNG điền "Ngày chứng từ" giống theo, không cần nhập tay ô đó nữa.
        date_ctrl = popup.child_window(auto_id=Controls.FIELD_NGAY_HACH_TOAN, control_type="Edit")
        _click_element_center(date_ctrl)
        time.sleep(0.2)
        send_keys("^a{DELETE}")
        send_keys(inv.ngay, pause=0.02)
        send_keys("{TAB}")
        time.sleep(0.3)

        # ---- Bảng chi tiết hàng hóa (grdDetail) ----
        # QUAN TRỌNG (người dùng nhắc 30/08/2026): cột %VAT + Tiền thuế GTGT theo dòng
        # nằm ở tab "2. Thuế", KHÔNG cùng tab với Mã hàng/SL/Đơn giá (tab "1. Hàng
        # tiền") — cây UIA chỉ lộ ra cột tương ứng khi ĐÚNG tab đang active (lazy-load,
        # xác nhận thật khi dò control). Vì vậy tách 2 vòng lặp theo 2 tab, KHÔNG gộp
        # chung 1 vòng như trước (bug thật đã tồn tại, có thể gây lỗi "không tìm thấy
        # cột" ở bước set thuế vì chưa từng chuyển tab).
        grid = popup.child_window(auto_id=Controls.GRID_HANG_TIEN_AUTO_ID, control_type="Custom")

        # ---- Vòng 1: tab "1. Hàng tiền" — Mã hàng (tra cứu), Số lượng, Đơn giá ----
        _click_element_center(
            popup.child_window(auto_id=Controls.TAB_HANG_TIEN_AUTO_ID, control_type="TabItem")
        )
        time.sleep(0.5)
        ten_hang_misa_list = []
        for row_idx, line in enumerate(inv.lines):
            ten_hang_misa = self._grid_search_ma_hang(grid, row_idx, line.ma_hang)
            ten_hang_misa_list.append(ten_hang_misa)
            self._grid_set_cell(grid, row_idx, Controls.GRID_COL_SO_LUONG, _format_vn_number(line.so_luong))
            self._grid_set_cell(grid, row_idx, Controls.GRID_COL_DON_GIA, _format_vn_number(line.don_gia))
            log.info(
                "Dòng %d: PDF='%s' -> MISA tra cứu ra='%s' | SL=%s DG=%s",
                row_idx, line.ma_hang, ten_hang_misa, line.so_luong, line.don_gia,
            )

        # ---- Vòng 2: tab "2. Thuế" — %VAT (luôn 10%) + Tiền thuế GTGT ghi đè theo dòng ----
        _click_element_center(
            popup.child_window(auto_id=Controls.TAB_THUE_AUTO_ID, control_type="TabItem")
        )
        time.sleep(0.5)
        for row_idx, line in enumerate(inv.lines):
            self._grid_set_cell_dropdown(grid, row_idx, Controls.GRID_COL_THUE_SUAT_PCT, "10%")
            self._grid_set_cell(
                grid, row_idx, Controls.GRID_COL_TIEN_THUE_DONG, _format_vn_number(line.tien_thue_dong)
            )
            log.info("Dòng %d: TienThue ghi đè=%s", row_idx, line.tien_thue_dong)

        # ---- Ô tổng cuối form (creTotalVATAmountOC) ----
        # CHƯA XÁC MINH: nếu MISA tự cộng dồn Tiền thuế GTGT từ các dòng ở trên thì ô
        # tổng này có thể đã tự đúng — set tay ở đây để an toàn, có thể là thao tác thừa.
        tien_thue_ctrl = popup.child_window(auto_id=Controls.FIELD_TONG_TIEN_THUE, control_type="Edit")
        _click_element_center(tien_thue_ctrl)
        time.sleep(0.2)
        send_keys("^a{DELETE}")
        send_keys(_format_vn_number(inv.tien_thue_that), pause=0.02)
        send_keys("{TAB}")

        # ---- Tab "Hóa đơn": kiểm tra/sửa "Số hóa đơn" khớp số hóa đơn thật trên PDF ----
        # THEO YÊU CẦU NGƯỜI DÙNG (30/08/2026): sau khi điền xong hàng+thuế, qua tab
        # "Hóa đơn" (nhóm tab trên cùng, khác nhóm "1. Hàng tiền/2. Thuế" ở dưới), đọc
        # "Số hóa đơn" (txtInvNo) — nếu đã khớp inv.so_hoa_don_goc thì giữ nguyên, nếu
        # không thì xoá và gõ lại đúng số thật.
        _click_element_center(
            popup.child_window(auto_id=Controls.TAB_HOA_DON_AUTO_ID, control_type="TabItem")
        )
        time.sleep(0.5)
        so_hd_ctrl = popup.child_window(auto_id=Controls.FIELD_SO_HOA_DON_THAT, control_type="Edit")
        so_hd_edit_area = self._cell_edit_area(so_hd_ctrl)
        current_value = (
            so_hd_edit_area.legacy_properties().get("Value", "") if so_hd_edit_area is not None else ""
        ).strip()
        target_value = (inv.so_hoa_don_goc or "").strip()
        if target_value and current_value != target_value:
            log.info(
                "Số hóa đơn hiện tại '%s' KHÔNG khớp số thật '%s' -> sửa lại.",
                current_value, target_value,
            )
            _click_element_center(so_hd_ctrl)
            time.sleep(0.2)
            send_keys("^a{DELETE}")
            send_keys(target_value, pause=0.02)
            send_keys("{TAB}")
            time.sleep(0.3)
        else:
            log.info("Số hóa đơn '%s' đã khớp số thật trên PDF, giữ nguyên.", current_value)

        return True

    # --------------------------------------------------------------- #
    def _grid_row(self, grid, row: int):
        """
        Lấy phần tử DataItem thứ `row` trong grdDetail. Mỗi DataItem có auto_id="-1"
        GIỐNG NHAU ở mọi dòng -> phải đánh theo index. LUÔN gọi lại hàm này để lấy tham
        chiếu MỚI sau mỗi lần ghi (không cache) vì grid có thể sinh thêm dòng mới sau
        khi commit dữ liệu vào dòng cuối.

        LỖI THẬT ĐÃ GẶP (30/08/2026): thứ tự `grid.descendants(control_type="DataItem")`
        trả về KHÔNG ổn định đúng theo thứ tự hiển thị trên màn hình giữa các lần gọi
        (nghi do grid ảo hoá/tái sử dụng control theo kiểu Janus GridEX) — dẫn tới gõ
        nhầm dữ liệu của dòng N vào đúng vị trí màn hình của dòng khác (dòng 3 "Cánh Gà
        Cay Tứ Xuyên" bị gõ đè bằng text của dòng 2 "Tôm Hoàng Kim"). Sửa: sắp xếp lại
        theo toạ độ Y THẬT trên màn hình (từ trên xuống) mỗi lần truy vấn — vị trí hiển
        thị mới là căn cứ đáng tin cậy, không phải thứ tự UIA trả về.
        """
        rows = grid.descendants(control_type="DataItem")
        # Bỏ các dòng "ảo" không có rect thật (VD dòng mẫu trống cuối bảng, rect
        # (0,0,0,0)) — nếu để lẫn vào sẽ bị sắp xếp lên đầu (top=0), làm lệch toàn bộ
        # chỉ số các dòng thật phía sau.
        rows = [r for r in rows if r.rectangle().height() > 0 and r.rectangle().width() > 0]
        if row >= len(rows):
            raise IndexError(f"Grid chỉ có {len(rows)} dòng thật, không có dòng thứ {row}")
        rows_sorted = sorted(rows, key=lambda r: r.rectangle().top)
        return rows_sorted[row]

    def _grid_cell(self, grid, row: int, col_title: str):
        """
        Tìm ô (ComboBox/Edit) theo title cột trong dòng `row`. row đã là phần tử cụ thể
        (không phải WindowSpecification) nên KHÔNG dùng .child_window() được — phải lọc
        thủ công qua .children() và so window_text(). Các cột chỉ hiện trong .children()
        khi tab tương ứng đang active (VD cột %VAT chỉ thấy khi tab "2. Thuế" đang mở).

        XÁC NHẬN THẬT (30/08/2026): hóa đơn > 9 dòng hàng — dòng nào vượt quá khung nhìn
        đang hiển thị của grid bị "ảo hoá" (DataItem có rect thật, nhưng CHƯA có ô con
        nào render, nên không tìm thấy cột dù dòng có tồn tại) cho tới khi được cuộn tới.
        Sửa: nếu lần tìm đầu không thấy cột, thử CUỘN grid xuống (click dòng cuối đang
        thấy + gửi phím mũi tên xuống — cách cuộn dùng phím, tôn trọng focus/visibility
        của grid, KHÁC với click toạ độ tuyệt đối vốn không tự cuộn) rồi tìm lại, tối đa
        vài lần, trước khi thật sự báo lỗi.
        """
        row_el = self._grid_row(grid, row)
        for c in row_el.children():
            if c.window_text() == col_title:
                return c
        for _ in range(4):
            self._scroll_grid_toward_row(grid, row)
            row_el = self._grid_row(grid, row)
            for c in row_el.children():
                if c.window_text() == col_title:
                    return c
        raise LookupError(
            f"Không thấy cột '{col_title}' ở dòng {row} — kiểm tra đã mở đúng tab chứa cột này chưa."
        )

    def _scroll_grid_toward_row(self, grid, row: int):
        """Cuộn grid Hàng tiền để dòng `row` lọt vào khung nhìn — xem _grid_cell()."""
        rows = grid.descendants(control_type="DataItem")
        rows = [r for r in rows if r.rectangle().height() > 0 and r.rectangle().width() > 0]
        if not rows:
            return
        rows_sorted = sorted(rows, key=lambda r: r.rectangle().top)
        last_idx = len(rows_sorted) - 1
        last_visible = rows_sorted[-1]
        self.main_win.set_focus()
        _click_element_center(last_visible)
        time.sleep(0.15)
        for _ in range(max(1, row - last_idx)):
            send_keys("{DOWN}")
            time.sleep(0.15)
        time.sleep(0.2)

    @staticmethod
    def _cell_edit_area(cell):
        for c in cell.children():
            if c.element_info.automation_id == "[Editor] Edit Area":
                return c
        return None

    @staticmethod
    def _read_cell_value(cell) -> str:
        """
        Đọc giá trị thật của 1 ô. XÁC NHẬN THẬT: window_text() trả về rỗng cho các ô
        trong grid này — phải đọc qua legacy_properties()['Value'] (bridge MSAA/Legacy
        IAccessible), khác hẳn các control UIA chuẩn khác trong ứng dụng.
        """
        edit_area = MisaAutomation._cell_edit_area(cell)
        if edit_area is None:
            return cell.window_text()
        try:
            return edit_area.legacy_properties().get("Value", "") or ""
        except Exception:
            return ""

    def _grid_set_cell(self, grid, row: int, col_title: str, value: str):
        """
        XÁC NHẬN THẬT (30/08/2026): PHẢI double-click để vào chế độ edit text thật —
        click đơn chỉ "chọn" ô, khiến phím Home/End sau đó bị GRID hiểu thành lệnh điều
        hướng "nhảy cột đầu/cuối" thay vì di chuyển con trỏ trong text (đã xác nhận qua
        quan sát trực tiếp). Double-click -> xoá sạch (End, Backspace nhiều lần) ->
        send_keys() gõ giá trị (giả lập bàn phím thật, KHÔNG dùng set_edit_text vì không
        kích hoạt được sự kiện nội bộ của MISA như tính lại Thành tiền) -> Tab để commit.
        """
        self.main_win.set_focus()
        cell = self._grid_cell(grid, row, col_title)
        edit_area = self._cell_edit_area(cell)
        if edit_area is None:
            raise LookupError(f"Không tìm thấy Edit Area trong ô '{col_title}' dòng {row}.")
        _click_element_center(edit_area)
        time.sleep(0.2)
        send_keys("^a{DELETE}")
        time.sleep(0.1)
        send_keys(str(value), pause=0.02, with_spaces=True)
        time.sleep(0.2)
        send_keys("{TAB}")
        time.sleep(0.3)

    def _grid_set_cell_dropdown(self, grid, row: int, col_title: str, option_text: str):
        """
        Chọn 1 lựa chọn trong ô dạng combobox có sẵn danh sách cố định (VD '% thuế GTGT':
        0%/5%/10%/KCT) — click nút dropdown, tìm ListItem con có window_text() khớp
        `option_text` (không dùng auto_id cố định vì thứ tự có thể khác giữa các dòng),
        rồi click chọn.
        """
        self.main_win.set_focus()
        cell = self._grid_cell(grid, row, col_title)
        for c in cell.children():
            if c.element_info.automation_id == "[Editor] dropdown button":
                _click_element_center(c)
                break
        time.sleep(0.3)
        cell_after = self._grid_cell(grid, row, col_title)
        for c in cell_after.children():
            if c.element_info.control_type == "ListItem" and c.window_text() == option_text:
                _click_element_center(c)
                time.sleep(0.2)
                return
        raise LookupError(f"Không tìm thấy lựa chọn '{option_text}' trong dropdown ô '{col_title}' dòng {row}.")

    @staticmethod
    def _name_match_score(ten_hang_pdf: str, ten_hang_misa: str) -> float:
        """
        Chấm điểm so khớp Tên hàng PDF với Tên hàng MISA trả về.
        XÁC NHẬN THẬT (30/08/2026, xem video demo người dùng quay): danh mục MISA của
        công ty đặt tên ĐẦY ĐỦ HƠN hóa đơn in ra — hóa đơn PDF ghi tên RÚT GỌN, còn MISA
        lưu tên đầy đủ + hậu tố "-<năm>" (VD PDF "Miến Xào Cua" ↔ MISA "Miến Xào Cua
        Lột-2026"; PDF "Sườn kinh đô" ↔ MISA "Sườn Kinh Đô-2026" — ĐÃ XÁC NHẬN ĐÚNG SẢN
        PHẨM qua đối chiếu Đơn giá khớp chính xác trong video demo, KHÔNG PHẢI hàng sai
        như nghi ngờ ban đầu). fuzz.ratio thông thường phạt nặng vì chênh lệch độ dài
        chuỗi, cho điểm quá thấp với các cặp ĐÚNG kiểu này.
        Sửa: nếu tên PDF (đã chuẩn hoá) là TIỀN TỐ đúng nghĩa của tên MISA (hoặc ngược
        lại) — tức MISA chỉ thêm mô tả/hậu tố vào SAU tên gốc, không đổi phần đầu — coi
        là khớp tuyệt đối (100 điểm). Nếu không phải quan hệ tiền tố, dùng lại
        fuzz.ratio làm phương án dự phòng (vẫn bắt được các trường hợp thật sự sai như
        2 tên hoàn toàn khác nhau).
        """
        a = _normalize(ten_hang_pdf)
        b = _normalize(ten_hang_misa)

        # XÁC NHẬN THẬT (30/08/2026, đọc danh mục thật 1 công ty khác — gần như TOÀN BỘ
        # sản phẩm chỉ khác nhau đúng hậu tố size, VD "...-Size M" / "...-Size L" /
        # "...-Size XL"): 2 tên chỉ lệch đúng chữ size cuối cùng vẫn giống nhau ~90-95%
        # theo fuzz.ratio (vì chỉ khác 1 ký tự trong chuỗi dài) — ĐỦ để vượt ngưỡng tin
        # cậy 92 một cách SAI, chọn nhầm size (khác giá tiền hẳn). Chặn cứng: nếu CẢ HAI
        # tên (sau chuẩn hoá) đều kết thúc bằng 1 từ size (S/M/L/XL/XXL) và 2 từ đó KHÁC
        # NHAU -> coi là sai chắc chắn (0 điểm), bất kể phần còn lại giống nhau bao
        # nhiêu. Không cần biết trước danh sách size cụ thể của từng công ty — tự suy ra
        # từ đúng vị trí (từ cuối cùng của tên).
        size_a = MisaAutomation._trailing_size_token(a)
        size_b = MisaAutomation._trailing_size_token(b)
        if size_a and size_b and size_a != size_b:
            return 0.0

        if a and b and (b.startswith(a) or a.startswith(b)):
            return 100.0
        a_acronym = MisaAutomation._acronym_expand(a, b)
        if a_acronym != a and (b.startswith(a_acronym) or a_acronym.startswith(b)):
            return 100.0

        # XÁC NHẬN THẬT (30/08/2026): hóa đơn PDF ghi "sốt" trong khi MISA lưu "xốt" (2
        # cách viết cùng phát âm /s/ giọng Bắc, hay bị lẫn khi nhập liệu/in hóa đơn) —
        # gộp "x" thành "s" CHỈ để tính điểm so khớp (không đụng gì tới dữ liệu gõ thật
        # vào MISA), giúp không bị phạt điểm oan vì khác 1 chữ cái do lẫn chính tả vùng
        # miền, vẫn giữ nguyên toàn bộ phần còn lại của phép so khớp.
        def _phonetic_fold(s: str) -> str:
            return s.replace("x", "s")

        a_fold, b_fold = _phonetic_fold(a), _phonetic_fold(b)
        if a_fold and b_fold and (b_fold.startswith(a_fold) or a_fold.startswith(b_fold)):
            return 100.0

        scores = [fuzz.ratio(a, b), fuzz.ratio(a_fold, b_fold)]
        if a_acronym != a:
            scores.append(fuzz.ratio(a_acronym, b))
        return float(max(scores))

    @staticmethod
    def _trailing_size_token(normalized: str) -> Optional[str]:
        """Trả về từ size (s/m/l/xl/xxl) nếu đó là TỪ CUỐI CÙNG của tên đã chuẩn hoá,
        ngược lại trả None. Xem _name_match_score() để biết lý do cần hàm này."""
        parts = normalized.split()
        if parts and parts[-1] in {"s", "m", "l", "xl", "xxl"}:
            return parts[-1]
        return None

    @staticmethod
    def _strip_diacritics(s: str) -> str:
        return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

    @staticmethod
    def _acronym_expand(pdf_norm: str, misa_norm: str) -> str:
        """
        XÁC NHẬN THẬT (30/08/2026): hóa đơn PDF đôi khi viết tắt 1 cụm từ theo kiểu CHỮ
        CÁI ĐẦU (VD "Đậu phụ cay TX" ↔ MISA "Đậu Phụ Cay Tứ Xuyên-2026" — "TX" = chữ
        cái đầu của "Tứ" + "Xuyên"). fuzz.ratio phạt nặng kiểu lệch này vì chênh lệch độ
        dài chuỗi lớn. Hàm này KHÔNG cần một danh sách viết tắt lập sẵn — tự dò: với mỗi
        từ ngắn (2-4 ký tự, toàn chữ cái) trong tên PDF, thử xem nó có khớp chữ cái đầu
        của một dãy từ liên tiếp trong tên MISA hay không; nếu có, thay từ viết tắt đó
        bằng chính cụm từ đầy đủ lấy từ tên MISA rồi mới so tiếp bằng fuzz.ratio.
        """
        misa_words = misa_norm.split()
        if not misa_words:
            return pdf_norm
        misa_initials = [
            MisaAutomation._strip_diacritics(w)[0] if w else "" for w in misa_words
        ]
        out_words = []
        for w in pdf_norm.split():
            replaced = False
            if w.isalpha() and 2 <= len(w) <= 4:
                for start in range(len(misa_words)):
                    span = len(w)
                    if (
                        start + span <= len(misa_initials)
                        and "".join(misa_initials[start:start + span]) == w
                    ):
                        out_words.append(" ".join(misa_words[start:start + span]))
                        replaced = True
                        break
            if not replaced:
                out_words.append(w)
        return " ".join(out_words)

    def _activate_row_ma_hang(self, grid, row: int):
        """
        Click trực tiếp vào giữa ô "Mã hàng" bằng _click_element_center() (gọi thẳng
        Windows API, không qua pywinauto's click_input()).
        LỊCH SỬ SỬA LỖI (30/08/2026, nhiều vòng):
          1. Nghi rectangle() bị cache sai -> thử "click dòng rồi Home" -> không hết lỗi.
          2. Nghi pywinauto's click_input() tính sai toạ độ -> đổi sang
             _click_element_center() (SetCursorPos+mouse_event thẳng) -> người dùng xác
             nhận: click ĐÃ ĐÚNG vào "Mã hàng" (nhìn thấy con trỏ chuột nằm đúng ô), NHƯNG
             ngay sau đó "nhảy" sang "TK chiết khấu" — hoá ra là do bước gửi phím
             "{HOME}" NGAY SAU KHI CLICK: trong grid này, Home ở trạng thái ô mới chỉ
             "được chọn" (chưa vào edit mode) bị hiểu thành lệnh điều hướng grid "nhảy
             tới cột CUỐI" thay vì "về đầu dòng" như kỳ vọng — không phải hành vi chuẩn
             nhưng đã xác nhận qua quan sát trực tiếp của người dùng. Vì click đã tự
             đúng vị trí Mã hàng sẵn (đây luôn là ô đầu tiên của 1 dòng trống mới), phím
             Home là THỪA và có hại -> đã bỏ hẳn.
        """
        row_el = self._grid_row(grid, row)
        for c in row_el.children():
            if c.window_text() == Controls.GRID_COL_MA_HANG:
                _click_element_center(c)
                time.sleep(0.3)
                return
        # Dòng > 9 có thể còn "ảo hoá" (chưa cuộn tới) -> thử cuộn rồi tìm lại (xem
        # _grid_cell() để biết chi tiết hiện tượng này), trước khi rơi vào dự phòng.
        for _ in range(4):
            self._scroll_grid_toward_row(grid, row)
            row_el = self._grid_row(grid, row)
            for c in row_el.children():
                if c.window_text() == Controls.GRID_COL_MA_HANG:
                    _click_element_center(c)
                    time.sleep(0.3)
                    return
        # Dự phòng: không thấy ô Mã hàng riêng (VD dòng còn quá mới) -> click vào dòng.
        _click_element_center(row_el)
        time.sleep(0.3)

    def _grid_search_ma_hang(self, grid, row: int, ten_hang_pdf: str) -> str:
        """
        XÁC NHẬN THẬT (30/08/2026, đã sửa nhiều lần):
          1. Gõ tên hàng rồi Tab ngay KHÔNG đủ tin cậy — MISA lọc ra NHIỀU ứng viên gần
             giống, và Tab chỉ commit ứng viên MISA tự highlight sẵn, không nhất thiết
             là ứng viên đúng nhất trong số đó.
          2. CLICK vào ListItem ứng viên trong dropdown KHÔNG có tác dụng chọn thật (xác
             nhận qua test — Tên hàng vẫn rỗng dù đã click + Tab). CHỈ gõ chữ (tên hoặc
             mã) + Tab mới được MISA nhận.
        Cách đúng: bước 1 gõ TÊN HÀNG để MISA lọc ra danh sách MÃ ứng viên (chỉ đọc
        window_text(), không click). Bước 2: với TỪNG mã ứng viên, xoá ô — gõ LẠI CHÍNH
        MÃ ĐÓ (không phải tên) — Tab để MISA nhận diện chính xác 1 mã duy nhất — đọc lại
        Tên hàng MISA điền ra — chấm điểm so khớp với tên gốc PDF. Chọn lại mã có điểm
        CAO NHẤT bằng đúng cách gõ-mã-rồi-Tab đó làm lựa chọn CUỐI CÙNG. Dưới ngưỡng
        LIVE_MATCH_CONFIDENCE_THRESHOLD thì xoá sạch ô, raise LowConfidenceMatchError.
        """

        def _type_and_commit(text: str):
            """
            Xoá ô Mã hàng, gõ `text`, Tab để MISA commit. Sau đó CLICK THÊM vào chính ô
            "Tên hàng" của dòng — XÁC NHẬN THẬT (30/08/2026, người dùng chỉ ra): giá trị
            Tên hàng chỉ thực sự "hiện ra" (render vào cây control, đọc được qua
            legacy_properties) sau khi ô đó được click/focus tới, Tab từ Mã hàng sang
            KHÔNG tự động làm việc này.
            """
            self.main_win.set_focus()
            self._activate_row_ma_hang(grid, row)
            cell = self._grid_cell(grid, row, Controls.GRID_COL_MA_HANG)
            send_keys("^a{DELETE}")
            time.sleep(0.1)
            if text:
                send_keys(text, pause=0.03, with_spaces=True)
                time.sleep(2.2)  # tăng từ 1.3s -> đủ thời gian MISA lọc xong với tên hàng dài
            send_keys("{TAB}")
            time.sleep(1.0)

            # Click thêm vào ô "Tên hàng" cùng dòng để ép MISA render giá trị thật.
            try:
                row_now = self._grid_row(grid, row)
                for c in row_now.children():
                    if c.window_text() == "Tên hàng":
                        _click_element_center(c)
                        time.sleep(0.5)
                        break
            except Exception:
                pass
            return cell

        def _read_row_fields():
            row_now = self._grid_row(grid, row)
            ma_hang_c, ten_hang_c = None, None
            for c in row_now.children():
                if c.window_text() == "Mã hàng":
                    ma_hang_c = c
                if c.window_text() == "Tên hàng":
                    ten_hang_c = c
            return (
                self._read_cell_value(ma_hang_c) if ma_hang_c is not None else "",
                self._read_cell_value(ten_hang_c) if ten_hang_c is not None else "",
            )

        # ---- Bước 1: gõ tên hàng để lấy danh sách MÃ ứng viên (chỉ đọc, không chọn) ----
        # XÁC NHẬN THẬT (30/08/2026): hóa đơn PDF đôi khi VIẾT TẮT 1 từ cuối tên (VD "Đậu
        # phụ cay TX" thay vì "...Tứ Xuyên") -> gõ NGUYÊN VĂN ra 0 ứng viên vì MISA lọc
        # kiểu "chứa chuỗi con", mà "tx" không phải chuỗi con của "tứ xuyên". Sửa: nếu gõ
        # nguyên văn ra 0 ứng viên, thử lại với phần NGẮN HƠN (bớt dần TỪ CUỐI CÙNG) — vế
        # đầu của tên luôn là phần chắc chắn có thật trong tên MISA, đủ để lọc ra ứng
        # viên; ứng viên nào đúng nhất vẫn do Bước 2 chấm điểm quyết định (dùng
        # _name_match_score đã có khả năng nhận diện lại viết tắt kiểu chữ cái đầu).
        # Ngoài ra, nếu tên PDF khớp 1 BÍ DANH đã biết trước (KNOWN_PRODUCT_ALIASES —
        # trường hợp tên PDF và tên MISA khác nhau hoàn toàn về từ ngữ, VD "Nước suối"
        # <-> "Aquafina", không phải lệch chính tả/viết tắt để tự suy luận được), thử
        # bí danh đó TRƯỚC TIÊN.
        pdf_norm_for_alias = _normalize(ten_hang_pdf)
        alias_texts = [
            alias_query
            for alias_key, alias_query in KNOWN_PRODUCT_ALIASES.items()
            if alias_key in pdf_norm_for_alias
        ]
        words = ten_hang_pdf.split()
        query_candidates = alias_texts + [" ".join(words[:n]) for n in range(len(words), 0, -1)]

        candidate_codes = []
        for query_text in query_candidates:
            self.main_win.set_focus()
            self._activate_row_ma_hang(grid, row)
            cell = self._grid_cell(grid, row, Controls.GRID_COL_MA_HANG)
            send_keys("^a{DELETE}")
            time.sleep(0.1)
            send_keys(query_text, pause=0.03, with_spaces=True)
            time.sleep(2.2)  # tăng từ 1.3s -> đủ thời gian MISA lọc xong với tên hàng dài
            cell = self._grid_cell(grid, row, Controls.GRID_COL_MA_HANG)
            candidate_codes = [
                c.window_text() for c in cell.children() if c.element_info.control_type == "ListItem"
            ]
            candidate_codes = [c for c in candidate_codes if c]
            if candidate_codes:
                if query_text != ten_hang_pdf:
                    log.info(
                        "  Gõ nguyên văn '%s' không ra ứng viên -> thử '%s' -> %d ứng viên.",
                        ten_hang_pdf, query_text, len(candidate_codes),
                    )
                break

        if not candidate_codes:
            _type_and_commit("")
            raise LowConfidenceMatchError(ten_hang_pdf, 0.0)

        # ---- Bước 2: với TỪNG mã ứng viên, gõ lại chính mã đó + Tab, đọc Tên hàng thật ----
        best_score, best_code, best_name = -1.0, None, None
        for code in candidate_codes:
            _type_and_commit(code)
            _, name_now = _read_row_fields()
            score = self._name_match_score(ten_hang_pdf, name_now)
            # Nếu tên PDF khớp 1 bí danh đã biết (VD "Nước suối") VÀ chính bí danh đó
            # (VD "aquafina") xuất hiện thật trong Tên hàng MISA trả về -> xác nhận
            # chắc chắn là đúng hàng theo định nghĩa nghiệp vụ, không phụ thuộc điểm
            # fuzz.ratio (vốn thấp vì 2 tên gần như không còn chữ nào giống nhau).
            name_now_norm = _normalize(name_now)
            for alias_text in alias_texts:
                if alias_text and alias_text in name_now_norm:
                    score = 100.0
                    break
            log.info("  Ứng viên mã '%s' -> Tên MISA='%s' (score=%.0f)", code, name_now, score)
            if score > best_score:
                best_score, best_code, best_name = score, code, name_now

        if best_code is None or best_score < LIVE_MATCH_CONFIDENCE_THRESHOLD:
            _type_and_commit("")
            raise LowConfidenceMatchError(ten_hang_pdf, max(best_score, 0.0))

        # ---- Chọn lại đúng mã điểm cao nhất làm lựa chọn CUỐI CÙNG ----
        _type_and_commit(best_code)
        ma_hang_value, ten_hang_value = _read_row_fields()

        log.info(
            "Tra cứu MISA: '%s' -> chọn Mã=%s Tên MISA='%s' (score=%.0f, ngưỡng=%.0f, %d ứng viên)",
            ten_hang_pdf, ma_hang_value, ten_hang_value, best_score,
            LIVE_MATCH_CONFIDENCE_THRESHOLD, len(candidate_codes),
        )
        if not ma_hang_value:
            raise LowConfidenceMatchError(ten_hang_pdf, best_score)

        return ten_hang_value

    # --------------------------------------------------------------- #
    def save(self, popup):
        """
        Bấm Cất — hoàn tất 1 hóa đơn.
        auto_id nút "Cất" ĐÃ DÒ được (đọc, chưa từng bấm thật):
        Controls.POPUP_BTN_CAT_AUTO_ID = "[Toolbar : MainToolbar Tools] Tool : mnuSave - Index : 6 "
        CẢNH BÁO: hàm này khi dry_run=False sẽ GHI DỮ LIỆU THẬT vào sổ sách kế toán —
        chỉ gọi sau khi đã test kỹ toàn bộ luồng điền dữ liệu (grid, thuế, khách hàng...)
        trên dữ liệu demo, có xác nhận rõ ràng từ người dùng cho từng bước ghi.
        """
        if self.dry_run:
            log.info("[DRY-RUN] Sẽ bấm nút 'Cất' ở đây.")
            return True
        if not Controls.POPUP_BTN_CAT_AUTO_ID:
            raise NotImplementedError(
                "Chưa dò auto_id nút 'Cất' trong popup — không tự ý đoán để tránh bấm nhầm."
            )
        _click_element_center(
            popup.child_window(auto_id=Controls.POPUP_BTN_CAT_AUTO_ID, control_type="Button")
        )
        time.sleep(1.0)  # chờ MISA xử lý lưu + validate nội bộ

        # ---- Xử lý hộp thoại cảnh báo chênh lệch thuế suất (nếu có) ----
        # XÁC NHẬN THẬT (30/08/2026): MISA hiện hộp thoại "Tổng tiền thuế GTGT <X> khác
        # (Tiền hàng - Chiết khấu) * Thuế suất <Y>, chênh lệch <Z>. Bạn có muốn cất
        # chứng từ này không?" — ĐÂY LÀ CẢNH BÁO DỰ KIẾN theo đúng quy tắc nghiệp vụ #3
        # (MISA tự tính theo %VAT cố định 10% trong khi ta ghi đè tiền thuế thật theo
        # PDF, luôn lệch với 10%*Tiền hàng) — KHÔNG PHẢI lỗi, phải bấm "Có" để tiếp tục.
        # LỖI THẬT ĐÃ GẶP: self.app.window(title="MISA SME.NET 2019", ...) KHÔNG tìm ra
        # hộp thoại (để sót, phải bấm tay) — dù đã dò xác nhận đúng title/control_type
        # qua print_control_identifiers. Sửa: dùng CÁCH ĐÃ CHỨNG MINH ỔN ĐỊNH — tìm hwnd
        # thô qua win32 EnumWindows (giống cách tìm frmMain ban đầu), rồi connect() bằng
        # đúng hwnd đó, thay vì tin vào self.app.window() tìm theo tiêu đề.
        try:
            dialog_hwnd = self._find_window_by_exact_title("MISA SME.NET 2019", timeout=3.0)
            if dialog_hwnd:
                dialog_app = Application(backend="uia").connect(handle=dialog_hwnd)
                dialog_win = dialog_app.window(handle=dialog_hwnd)
                co_button = dialog_win.child_window(auto_id="6", control_type="Button")
                log.info("Hộp thoại cảnh báo chênh lệch thuế suất xuất hiện (đúng dự kiến) — bấm 'Có'.")
                _click_element_center(co_button)
                time.sleep(1.0)
            else:
                log.info("Không có hộp thoại cảnh báo nào xuất hiện — Cất thành công trực tiếp.")
        except Exception as e:
            log.warning("Lỗi khi xử lý hộp thoại xác nhận Cất (nếu có): %s — cần tự kiểm tra tay.", e)

        # ---- Chờ MISA xử lý xong (đồng bộ hóa đơn điện tử...) rồi đóng tab ----
        # THEO NGƯỜI DÙNG (30/08/2026): sau Cất, MISA mất khoảng 6-8 giây mới quay lại
        # được màn hình danh sách để thêm hóa đơn mới — cần chờ đủ trước khi thao tác
        # tiếp, và chủ động bấm "Đóng" để đóng tab chứng từ vừa lưu, tránh tích tụ nhiều
        # tab mở khi chạy hàng loạt (mỗi "Thêm" mở 1 tab MDI nội bộ mới, không tự đóng).
        time.sleep(7.0)
        try:
            _click_element_center(
                popup.child_window(auto_id=Controls.POPUP_BTN_DONG_AUTO_ID, control_type="Button")
            )
            time.sleep(1.0)
        except Exception as e:
            log.warning("Không đóng được tab hóa đơn vừa Cất (%s) — vẫn tiếp tục, có thể còn tab thừa.", e)

        return True

    @staticmethod
    def _find_window_by_exact_title(title: str, timeout: float = 3.0):
        """
        Tìm hwnd của 1 cửa sổ top-level có TIÊU ĐỀ KHỚP CHÍNH XÁC (win32 EnumWindows
        thô, không qua pywinauto/UIA) — cách đã xác nhận ổn định khi tìm frmMain lúc
        đầu. Thử lại trong `timeout` giây vì hộp thoại có thể xuất hiện có độ trễ sau
        khi bấm Cất. Trả về None nếu không tìm thấy trong thời gian chờ.
        """
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        deadline = time.time() + timeout
        while time.time() < deadline:
            found = []

            def foreach_window(hwnd, _lparam):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0 and user32.IsWindowVisible(hwnd):
                    buff = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buff, length + 1)
                    if buff.value == title:
                        found.append(hwnd)
                return True

            user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(foreach_window), 0)
            if found:
                return found[0]
            time.sleep(0.3)
        return None

    # --------------------------------------------------------------- #
    def close_popup_if_open(self, popup):
        """
        NGHI VẤN NGUYÊN NHÂN LỖI THẬT (29/08/2026): khi 1 hóa đơn bị dừng giữa chừng
        (VD tra cứu mã hàng thất bại), hàm này TRƯỚC ĐÂY gọi popup.close() để đóng form
        dở dang — nhưng MISA nhiều khả năng hiện hộp thoại "Bạn có muốn lưu thay đổi
        không?" khi đóng form đã sửa dở, và hộp thoại đó có thể đã bị xác nhận "Có" một
        cách ngoài ý muốn (do phím thừa còn sót lại trong hàng đợi input), khiến 1 hóa
        đơn RỖNG (chưa nhập mặt hàng) bị lưu thật vào sổ sách — xác nhận thực tế đã xảy
        ra khi test. Vì auto_id của hộp thoại xác nhận này CHƯA được dò, và rủi ro bấm
        nhầm quá cao với dữ liệu kế toán thật, hàm này KHÔNG TỰ ĐỘNG ĐÓNG POPUP NỮA —
        chỉ dừng lại, log cảnh báo, để người dùng tự kiểm tra và đóng tay (chọn "Không
        lưu" nếu MISA hỏi) trước khi làm gì tiếp theo.
        """
        if self.dry_run or popup is None:
            return
        log.warning(
            "Hóa đơn bị dừng giữa chừng — form ĐANG ĐỂ MỞ DỞ DANG trên MISA, KHÔNG tự "
            "đóng để tránh rủi ro lưu nhầm dữ liệu rỗng. Vui lòng tự kiểm tra trên màn "
            "hình MISA: nếu có hộp thoại hỏi lưu thay đổi, chọn 'Không lưu' / 'Cancel'."
        )

    def cancel_popup_safely(self, popup) -> bool:
        """
        Đóng popup MÀ KHÔNG LƯU — dùng khi 1 dòng hàng không đủ độ tin cậy để tự chọn
        (LowConfidenceMatchError, nghiệp vụ #4: không tự đoán) — để BATCH tự chạy tiếp
        sang hóa đơn kế tiếp, không cần dừng cả loạt chờ người can thiệp tay mỗi lần.
        Chỉ dùng ở điểm dừng "sạch": ngay trước khi raise, `_grid_search_ma_hang()` đã tự
        xoá ô Mã hàng của dòng đang lỗi -> phần dữ liệu đã gõ trước đó của các dòng khác
        vẫn còn, nhưng bản chất đây LÀ hóa đơn cần huỷ để nhập tay lại từ đầu (đúng quy
        tắc #6 — không được Cất hóa đơn còn thiếu/sai dữ liệu).
        KHÔNG dùng lại cho except Exception chung (lỗi không rõ dạng) — trạng thái form
        khi đó không chắc "sạch", vẫn nên dừng cả batch để người kiểm tra tay.
        Cách làm: bấm nút "Đóng" (đúng thao tác người dùng thật) -> MISA tự hỏi "Có muốn
        lưu thay đổi?" -> tìm cửa sổ hộp thoại qua win32 EnumWindows (đã chứng minh ổn
        định ở save()) -> bấm "Không" — auto_id="7" theo đúng quy ước Windows MessageBox
        chuẩn (IDNO=7), suy ra từ việc auto_id="6"=IDYES đã xác nhận đúng cho nút "Có" ở
        save(). CHƯA dò trực tiếp auto_id="7" trên máy thật — lần chạy ĐẦU TIÊN sau khi
        thêm tính năng này cần người dùng quan sát kỹ để xác nhận bấm đúng nút.
        """
        if self.dry_run or popup is None:
            return True
        try:
            _click_element_center(
                popup.child_window(auto_id=Controls.POPUP_BTN_DONG_AUTO_ID, control_type="Button")
            )
            time.sleep(1.0)
            dialog_hwnd = self._find_window_by_exact_title("MISA SME.NET 2019", timeout=3.0)
            if dialog_hwnd:
                dialog_app = Application(backend="uia").connect(handle=dialog_hwnd)
                dialog_win = dialog_app.window(handle=dialog_hwnd)
                khong_button = dialog_win.child_window(auto_id="7", control_type="Button")
                log.info(
                    "Hộp thoại hỏi lưu thay đổi xuất hiện (đúng dự kiến) — bấm 'Không' để HỦY, không lưu."
                )
                _click_element_center(khong_button)
                time.sleep(1.0)
            else:
                log.info("Đóng popup không có hộp thoại hỏi lưu nào xuất hiện — coi như đã đóng an toàn.")
            return True
        except Exception as e:
            log.warning(
                "Không tự đóng/hủy được popup an toàn (%s) — DỪNG batch, cần bạn tự kiểm tra tay.", e
            )
            return False


class LowConfidenceMatchError(Exception):
    """Ném ra khi mã hàng không đủ độ tin cậy để tự động chọn — cần người duyệt tay."""
    def __init__(self, ten_hang: str, score: float):
        super().__init__(f"Mã hàng cho '{ten_hang}' có độ tin cậy thấp ({score:.0f}) — cần xác nhận tay.")
        self.ten_hang = ten_hang
        self.score = score


"""
============================================================================
DÒ CONTROL — NHẬT KÝ THỰC TẾ đã thực hiện trên máy chạy MISA thật
============================================================================
(Cập nhật sau phiên dò control thật ngày 29/08/2026 — đơn vị CÔNG TY TNHH GUABAO ONE,
máy công ty qua UltraViewer. Toàn bộ auto_id trong class Controls ở trên lấy từ đây.)

BÀI HỌC QUAN TRỌNG NHẤT: phải chạy Python dưới quyền ADMINISTRATOR.
  - MISA SME.NET 2019 chạy elevated (Run as Administrator).
  - Windows UIPI (User Interface Privilege Isolation) chặn UI Automation đọc/ghi sâu
    vào tiến trình có quyền cao hơn từ tiến trình quyền thấp hơn — script không-admin
    chỉ thấy được lớp Dialog/TitleBar ngoài cùng, KHÔNG thấy control con nào cả.
  - Cách mở PowerShell Admin nhanh:
      Start-Process powershell -Verb RunAs -ArgumentList "-NoExit","-Command","cd '<path>'"
    (Windows sẽ hiện hộp thoại UAC, người ngồi máy phải bấm Yes.)

CÁCH KẾT NỐI ĐÚNG (đã xác nhận hoạt động):
  1. Tìm handle cửa sổ chính bằng win32 EnumWindows thô (ctypes) — không phụ thuộc
     pywinauto/UIA để liệt kê, tránh vấn đề Desktop(backend="uia").windows() đôi khi
     không enum ra được cửa sổ MISA.
  2. app = Application(backend="uia").connect(handle=<hwnd>)
     win = app.window(handle=<hwnd>)  # hoặc auto_id="frmMain"

KIẾN TRÚC MDI NỘI BỘ — điểm khác biệt lớn nhất so với giả định ban đầu:
  - "Bán hàng" (danh sách) và popup nhập/sửa chứng từ KHÔNG phải cửa sổ Windows
    riêng biệt (app.windows() không liệt kê ra) — chúng là các Window/Dialog NHÚNG
    bên trong frmMain, dạng cây con.
  - Cách tìm: main_win.descendants(control_type="Window") rồi lọc theo auto_id/title,
    hoặc trực tiếp main_win.child_window(auto_id="frmSAVoucherDetail", control_type="Window")
    sau khi đã Thêm/Xem 1 chứng từ.

GRID "HÀNG TIỀN" (grdDetail) — kiểu Janus GridEX (không phải DevExpress như đoán ban đầu):
  - Cột nhận diện qua auto_id="[Column Header] <FieldName>" (VD InventoryItemCode,
    Description, Quantity, UnitPrice, AmountOC...).
  - Mỗi dòng dữ liệu là DataItem nhưng auto_id="-1" GIỐNG NHAU ở MỌI DÒNG — không thể
    đánh địa chỉ theo auto_id, phải dùng chỉ số (index) trong descendants(control_type="DataItem").
  - QUAN TRỌNG (đã sửa hiểu lầm ban đầu): KHÔNG cần double-click để "vào chế độ edit"
    — mỗi ô (ComboBox/Edit) trong dòng đang active đã CÓ SẴN control con thật:
      + "[Editor] Edit Area" (Edit) — vùng nhập text thật, set_edit_text() thẳng vào đây.
      + Với ComboBox: thêm "[Editor] dropdown button" (Button) và danh sách con
        "[Editor] [valuelist] ValueListItem N" (ListItem) cho từng lựa chọn.
    Xác nhận thật bằng cách đọc cây con của dòng 0 (không double-click, không gõ):
    ví dụ ô "% thuế GTGT" có đúng 4 ListItem: 0%(0), 5%(1), 10%(2), KCT(3); ô
    "TK thuế GTGT" liệt kê đầy đủ ~190 tài khoản trong hệ thống tài khoản kế toán.
  - Vì vậy: row.children() (không phải child_window — row là phần tử cụ thể, không
    phải WindowSpecification) lọc theo window_text()==tên cột, rồi lấy Edit Area con.
  - Cột %VAT + Tiền thuế theo dòng (% thuế GTGT, Tiền thuế GTGT, TK thuế GTGT, Diễn giải
    thuế) NẰM TRONG CÙNG grid này, chỉ xuất hiện trong cây UIA khi tab "2. Thuế" đang
    active (WinForms tab lazy-load nội dung tab) — phải click đúng tab trước khi đọc/set
    các cột thuộc tab đó.

Hàm _grid_row/_grid_cell/_grid_set_cell/_grid_set_cell_dropdown trong class MisaAutomation
đã viết theo đúng cấu trúc xác nhận ở trên, nhưng CHƯA TEST GÕ GIÁ TRỊ THẬT (mới chỉ đọc
cấu trúc, chưa được phép nhập liệu) — bắt buộc test kỹ trên dữ liệu demo trước khi dùng thật.

CÒN THIẾU / CHƯA DÒ (việc tiếp theo):
  1. [ĐÃ DÒ] auto_id nút "Cất" trong popup detail = "[Toolbar : MainToolbar Tools] Tool
     : mnuSave - Index : 6 " (nằm trên MainToolbar của form nhập chứng từ, không phải
     ListToolbar của danh sách). Mới XÁC NHẬN qua đọc cây control — CHƯA từng bấm thật.
  2. [ĐÃ DÒ] "Chứng từ bán hàng" KHÔNG nằm trong MainMenu trên cùng — nó là 1 mục
     (DataItem) trong thanh điều hướng trái kiểu Outlook, auto_id="ebMain", thuộc
     Group "Bán hàng". [Group]/[Item] N là chỉ số TƯƠNG ĐỐI (lặp lại ở nhiều nhóm) nên
     phải tìm theo window_text()=="Chứng từ bán hàng", không dùng auto_id cố định — xem
     MisaAutomation.ensure_ban_hang_list_open(). CHƯA từng click mục nav này thật (lần
     dò trước người dùng đã tự mở sẵn danh sách bằng tay).
  3. Chưa test THẬT việc set_edit_text() vào Edit Area và click ListItem trong dropdown
     — mới xác nhận các control này TỒN TẠI, chưa xác nhận thao tác ghi thành công.
  4. Chưa xác nhận: ô tổng "Tiền thuế GTGT" (creTotalVATAmountOC) có cho sửa tay được
     không, hay MISA tự khoá/tự tính lại theo tổng các dòng (nếu khoá, phải ghi đè ở
     cấp DÒNG thay vì cấp TỔNG — xem cột "Tiền thuế GTGT" theo dòng trong grid).
  5. Chưa dò auto_id combobox "Khách hàng" khi gõ "Khách lẻ" — cần xác nhận đúng có
     item "Khách lẻ" sẵn trong danh mục khách hàng của GUABAO ONE hay phải tạo mới.

QUY TẮC AN TOÀN (không đổi):
  - LUÔN test trên máy/dữ liệu demo trước, dùng --dry-run trước, xác nhận từng bước
    điền đúng trước khi bấm Cất thật — vì MISA lưu trực tiếp vào sổ sách kế toán,
    sai sót ở đây ảnh hưởng số liệu tài chính thật.
  - Nếu cần dò thêm control ẩn trong grid, dùng Microsoft "Accessibility Insights for
    Windows" (hover từng ô) làm công cụ bổ trợ khi UIA cây tĩnh không đủ.
"""

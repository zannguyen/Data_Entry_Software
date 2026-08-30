"""
gui_app.py
----------
Giao diện Desktop (Tkinter — có sẵn trong Python, không cần cài thêm) để:
  1. Chọn thư mục chứa PDF hóa đơn + file danh mục hàng hóa MISA (CSV/Excel).
  2. Đọc & đối chiếu tất cả hóa đơn, hiển thị dạng HÀNG ĐỢI (giống mockup HTML đã
     thiết kế trước đây): mỗi dòng = 1 hóa đơn, trạng thái Sẵn sàng / Cần duyệt tay.
  3. Chọn 1 hóa đơn để xem chi tiết: từng dòng hàng, mã hàng đối chiếu (fuzzy match)
     kèm độ tin cậy, cho phép SỬA TAY mã hàng nếu độ tin cậy thấp.
  4. Chạy "Dry-run" (an toàn, chỉ log) hoặc "Nhập vào MISA" (ghi thật — có hộp thoại
     xác nhận rõ ràng, và bắt buộc chạy dưới quyền Administrator).

Chạy: python gui_app.py
Ứng dụng TỰ ĐỘNG kiểm tra quyền Administrator khi khởi động (bắt buộc vì MISA chạy
elevated — xem ghi chú UIPI trong misa_automation.py). Nếu chưa đủ quyền, tự bật
hộp thoại UAC (Windows sẽ hỏi xác nhận — người ngồi máy phải tự bấm Yes, đây là giới
hạn bảo mật của Windows, không có cách nào bỏ qua bước bấm Yes này) rồi tự khởi động
lại chính mình dưới quyền Admin và thoát tiến trình cũ (không chạy 2 cửa sổ chồng nhau).
"""

import ctypes
import os
import sys


def _set_dpi_aware():
    """
    XÁC NHẬN THẬT (30/08/2026): PHẢI gọi hàm này TRƯỚC KHI import tkinter — Windows chỉ
    cho khai báo DPI-awareness của tiến trình ĐÚNG 1 LẦN DUY NHẤT (mọi lời gọi sau lần
    đầu tiên đều bị bỏ qua). Việc `import tkinter` load thư viện Tcl/Tk (tcl86t.dll/
    tk86t.dll) dường như đã tự khai báo DPI-awareness ngay khi nạp DLL — nếu gọi hàm
    này SAU dòng import tkinter (như bản sửa lần đầu), lời gọi luôn thất bại âm thầm
    (tự kiểm chứng qua GetProcessDpiAwareness() vẫn trả về 0/Unaware dù code chạy không
    lỗi). Máy đang chạy scale màn hình 150% (144 DPI) — không khai báo DPI-aware khiến
    Windows "ảo hoá" toạ độ chuột theo hệ 96 DPI cho pywinauto's click_input(), trong
    khi UI Automation trả toạ độ phần tử theo pixel vật lý thật -> lệch 150%, click sai
    vị trí (càng xa góc màn hình càng lệch nặng — đã xác nhận: có lúc click trúng nhầm
    nút Task View trên taskbar Windows).
    """
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness.argtypes = [ctypes.c_int]
        ctypes.windll.shcore.SetProcessDpiAwareness.restype = ctypes.c_long
        hr = ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
        if hr == 0:  # S_OK
            return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


_set_dpi_aware()

import glob
import json
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# File log LƯU TRÊN ĐĨA (không chỉ trong bộ nhớ) các Số hóa đơn ĐÃ Cất thành công vào
# MISA thật. XÁC NHẬN THẬT (30/08/2026): nếu chỉ theo dõi bằng InvoiceRow.status trong
# bộ nhớ, mỗi lần app RESTART (đã xảy ra nhiều lần trong lúc sửa lỗi) sẽ mất hết dấu vết
# -> lần đọc PDF tiếp theo xử lý lại TỪ ĐẦU cả những hóa đơn đã Cất rồi, cất TRÙNG vào sổ
# sách thật (lỗi thật đã xảy ra với hóa đơn 00000631). File này khắc phục việc đó — sống
# sót qua mọi lần restart, vì lưu ngay trên đĩa ngay khi Cất thành công.
DA_CAT_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "da_cat_log.json")


def _load_da_cat_set() -> set:
    if not os.path.exists(DA_CAT_LOG_PATH):
        return set()
    try:
        with open(DA_CAT_LOG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return set(data.get("da_cat", []))
    except Exception:
        return set()


def _append_da_cat_log(so_hoa_don: str):
    existing = _load_da_cat_set()
    existing.add(so_hoa_don)
    _write_da_cat_set(existing)


def _remove_da_cat_log(so_hoa_don: str):
    existing = _load_da_cat_set()
    existing.discard(so_hoa_don)
    _write_da_cat_set(existing)


def _write_da_cat_set(values: set):
    tmp_path = DA_CAT_LOG_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump({"da_cat": sorted(values)}, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, DA_CAT_LOG_PATH)


def _relaunch_as_admin_if_needed():
    """
    Chỉ áp dụng trên Windows. Nếu tiến trình hiện tại KHÔNG chạy quyền Administrator,
    dùng ShellExecuteW với verb "runas" để Windows tự hiện hộp thoại UAC và khởi động
    lại đúng lệnh này (python gui_app.py ...) dưới quyền Admin, sau đó thoát tiến trình
    không có quyền hiện tại (sys.exit) để tránh mở 2 cửa sổ.
    """
    if os.name != "nt":
        return
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        is_admin = False
    if is_admin:
        return

    params = " ".join(f'"{a}"' for a in sys.argv)
    try:
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, params, None, 1
        )
        # ShellExecuteW trả về giá trị > 32 nếu thành công khởi chạy tiến trình mới.
        if int(ret) > 32:
            sys.exit(0)
        else:
            # Người dùng bấm "No" trên hộp thoại UAC, hoặc lỗi khác -> chạy tiếp KHÔNG
            # có quyền Admin (một số thao tác đọc vẫn hoạt động, nhưng nút "Nhập THẬT
            # vào MISA" sẽ báo lỗi khi thao tác thật vì bị Windows UIPI chặn).
            pass
    except Exception:
        pass


_relaunch_as_admin_if_needed()


from pdf_extractor import extract_invoices, Invoice
from product_matcher import ProductMatcher, build_sample_matcher, MatchResult
from misa_automation import (
    MisaAutomation, InvoiceToEnter, LineToEnter, LowConfidenceMatchError,
)


class SimplePathPicker(tk.Toplevel):
    """
    Bộ chọn thư mục/file TỰ VẼ bằng Tkinter thuần — KHÔNG dùng hộp thoại gốc của
    Windows (filedialog.askdirectory/askopenfilename). Lý do: hộp thoại kiểu mới của
    Windows (COM/IFileDialog) bị TREO ("Not Responding") khi chạy trong môi trường
    remote qua UltraViewer + quyền Administrator — xác nhận thực tế khi test. Tự vẽ
    bằng Tkinter tránh hoàn toàn phụ thuộc COM nên không bị treo trong môi trường này.

    select_file=False: chọn 1 THƯ MỤC. select_file=True: chọn 1 FILE khớp `extensions`.
    Kết quả trả về qua self.result (None nếu người dùng bấm Hủy).
    """

    def __init__(self, parent, title: str, start_dir: str, select_file: bool = False,
                 extensions: tuple = ()):
        super().__init__(parent)
        self.title(title)
        self.geometry("560x420")
        self.transient(parent)
        self.grab_set()
        self.select_file = select_file
        self.extensions = tuple(e.lower() for e in extensions)
        self.result = None

        start = start_dir if start_dir and os.path.isdir(start_dir) else os.path.expanduser("~")
        self.current_dir = tk.StringVar(value=start)

        path_bar = ttk.Frame(self, padding=6)
        path_bar.pack(fill="x")
        ttk.Label(path_bar, text="Đường dẫn:").pack(side="left")
        self.path_entry = ttk.Entry(path_bar, textvariable=self.current_dir)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(path_bar, text="Đi tới", command=self._navigate_to_entry).pack(side="left")

        self.listbox = tk.Listbox(self, font=("Consolas", 10))
        self.listbox.pack(fill="both", expand=True, padx=6, pady=4)
        self.listbox.bind("<Double-Button-1>", self._on_double_click)

        btn_bar = ttk.Frame(self, padding=6)
        btn_bar.pack(fill="x")
        label = "Chọn file này" if select_file else "Chọn thư mục này"
        ttk.Button(btn_bar, text=label, command=self._confirm).pack(side="right")
        ttk.Button(btn_bar, text="Hủy", command=self._cancel).pack(side="right", padx=6)

        self._refresh_listing()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _refresh_listing(self):
        self.listbox.delete(0, "end")
        d = self.current_dir.get()
        self.listbox.insert("end", ".. (lên thư mục cha)")
        try:
            entries = sorted(os.listdir(d), key=str.lower)
        except Exception as e:
            self.listbox.insert("end", f"(Lỗi đọc thư mục: {e})")
            return
        dirs = [e for e in entries if os.path.isdir(os.path.join(d, e))]
        for e in dirs:
            self.listbox.insert("end", f"[Thư mục] {e}")
        if self.select_file:
            files = [
                e for e in entries
                if os.path.isfile(os.path.join(d, e))
                and (not self.extensions or e.lower().endswith(self.extensions))
            ]
            for e in files:
                self.listbox.insert("end", e)

    def _navigate_to_entry(self):
        d = self.current_dir.get()
        if os.path.isdir(d):
            self._refresh_listing()
        else:
            messagebox.showwarning("Không hợp lệ", "Thư mục không tồn tại.", parent=self)

    def _on_double_click(self, _event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        text = self.listbox.get(sel[0])
        d = self.current_dir.get()
        if text.startswith(".. "):
            parent_dir = os.path.dirname(d.rstrip("\\/"))
            if parent_dir and parent_dir != d:
                self.current_dir.set(parent_dir)
                self._refresh_listing()
            return
        if text.startswith("[Thư mục] "):
            name = text[len("[Thư mục] "):]
            self.current_dir.set(os.path.join(d, name))
            self._refresh_listing()
            return
        # Là 1 file (chỉ có khi select_file=True)
        if self.select_file:
            self.result = os.path.join(d, text)
            self.destroy()

    def _confirm(self):
        if self.select_file:
            messagebox.showinfo(
                "Chọn file", "Double-click vào tên file trong danh sách để chọn.", parent=self
            )
            return
        d = self.current_dir.get()
        if not os.path.isdir(d):
            messagebox.showwarning("Không hợp lệ", "Thư mục không tồn tại.", parent=self)
            return
        self.result = d
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class InvoiceRow:
    """Gói 1 hóa đơn + kết quả đối chiếu mã hàng, dùng chung cho danh sách và chi tiết."""

    def __init__(self, invoice: Invoice):
        self.invoice = invoice
        self.matches: list[MatchResult] = []
        self.overrides: dict[int, str] = {}  # row_idx -> mã hàng do người dùng chọn tay
        self.status = "chua_doi_chieu"  # chua_doi_chieu | san_sang | can_duyet | da_cat | loi
        self.error_msg = ""

    @property
    def can_auto_process(self) -> bool:
        if self.invoice.warnings:
            return False
        if not self.matches:
            return False
        for i, m in enumerate(self.matches):
            if i in self.overrides:
                continue
            if not m.can_auto_accept:
                return False
        return True

    def resolved_ma_hang(self, idx: int) -> str:
        if idx in self.overrides:
            return self.overrides[idx]
        m = self.matches[idx]
        return m.ma_hang or ""


class App(tk.Tk):
    STATUS_LABEL = {
        "chua_doi_chieu": "Chưa đối chiếu",
        "san_sang": "✓ Sẵn sàng",
        "can_duyet": "⚠ Cần duyệt tay",
        "da_cat": "✔ Đã cất",
        "loi": "✗ Lỗi",
    }

    def __init__(self):
        super().__init__()
        try:
            is_admin = bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            is_admin = False
        admin_label = "Administrator" if is_admin else "KHÔNG có quyền Admin — bấm 'Nhập THẬT' sẽ lỗi"
        self.title(f"RPA Nhập hóa đơn MISA SME.NET 2019  [{admin_label}]")
        self.geometry("1280x760")

        self.pdf_dir = tk.StringVar()
        self.catalog_path = tk.StringVar()
        self.matcher: ProductMatcher | None = None
        self.rows: list[InvoiceRow] = []
        self.selected_row: InvoiceRow | None = None

        self._build_top_bar()
        self._build_main_area()
        self._build_log_area()

    def report_callback_exception(self, exc, val, tb):
        """
        Ghi đè handler mặc định của Tkinter — bình thường Tkinter chỉ in traceback ra
        stderr (console có thể bị ẩn/không tồn tại khi app chạy qua UAC relaunch), khiến
        lỗi trông như "không có gì xảy ra". Giờ mọi lỗi trong callback (bấm nút, chọn
        dòng...) sẽ hiện rõ bằng popup + ghi vào Log để không bị nuốt im lặng nữa.
        """
        import traceback
        msg = "".join(traceback.format_exception(exc, val, tb))
        try:
            self._log("LỖI (callback):\n" + msg)
        except Exception:
            pass
        messagebox.showerror("Lỗi", f"{val}\n\nXem chi tiết trong panel Log bên dưới.")

    # ------------------------------------------------------------------ #
    def _build_top_bar(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(side="top", fill="x")

        ttk.Label(bar, text="PDF hóa đơn:").pack(side="left")
        ttk.Entry(bar, textvariable=self.pdf_dir, width=40).pack(side="left", padx=4)
        ttk.Button(bar, text="Chọn file...", command=self._choose_pdf_file).pack(side="left")
        ttk.Button(bar, text="Chọn thư mục...", command=self._choose_pdf_dir).pack(side="left", padx=(2, 0))

        ttk.Label(bar, text="   Danh mục hàng hóa:").pack(side="left")
        ttk.Entry(bar, textvariable=self.catalog_path, width=35).pack(side="left", padx=4)
        ttk.Button(bar, text="Chọn...", command=self._choose_catalog).pack(side="left")

        ttk.Button(bar, text="Đọc & đối chiếu", command=self._load_and_match).pack(side="left", padx=(12, 4))

    def _build_main_area(self):
        main = ttk.Panedwindow(self, orient="horizontal")
        main.pack(side="top", fill="both", expand=True, padx=8, pady=4)

        # ---- Bên trái: hàng đợi hóa đơn ----
        left = ttk.Frame(main)
        main.add(left, weight=1)

        ttk.Label(left, text="Hàng đợi hóa đơn (theo số tăng dần)", font=("Segoe UI", 10, "bold")).pack(
            anchor="w"
        )
        columns = ("so_hd", "ngay", "so_dong", "tong_tt", "trang_thai")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", height=22, selectmode="browse")
        for col, label, width in [
            ("so_hd", "Số HĐ", 90),
            ("ngay", "Ngày", 90),
            ("so_dong", "Số dòng", 60),
            ("tong_tt", "Tổng TT", 110),
            ("trang_thai", "Trạng thái", 130),
        ]:
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, anchor="center" if col != "tong_tt" else "e")
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self._on_select_invoice)

        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=4)
        ttk.Button(btns, text="Dry-run tất cả (an toàn)", command=lambda: self._run_all(dry_run=True)).pack(
            side="left"
        )
        ttk.Button(
            btns, text="Nhập THẬT vào MISA...", command=lambda: self._run_all(dry_run=False)
        ).pack(side="left", padx=6)
        ttk.Button(
            btns, text="Nhập THẬT — chỉ hóa đơn đang chọn",
            command=lambda: self._run_all(dry_run=False, only_selected=True),
        ).pack(side="left")

        # ---- Bên phải: chi tiết hóa đơn đang chọn ----
        right = ttk.Frame(main)
        main.add(right, weight=2)

        self.detail_header = ttk.Label(right, text="Chọn 1 hóa đơn ở danh sách bên trái", font=("Segoe UI", 10, "bold"))
        self.detail_header.pack(anchor="w")

        self.warning_label = ttk.Label(right, text="", foreground="#b00020", wraplength=700, justify="left")
        self.warning_label.pack(anchor="w", pady=(2, 6))

        line_cols = ("ten_pdf", "sl", "dg", "ma_match", "score", "ma_final", "tien_thue")
        self.line_tree = ttk.Treeview(right, columns=line_cols, show="headings", height=12)
        for col, label, width in [
            ("ten_pdf", "Tên hàng (PDF)", 220),
            ("sl", "SL", 45),
            ("dg", "Đơn giá", 90),
            ("ma_match", "Mã gợi ý (MISA)", 130),
            ("score", "Điểm", 50),
            ("ma_final", "Mã sẽ nhập", 110),
            ("tien_thue", "Tiền thuế/dòng", 90),
        ]:
            self.line_tree.heading(col, text=label)
            self.line_tree.column(col, width=width, anchor="w" if col in ("ten_pdf", "ma_match") else "center")
        self.line_tree.pack(fill="both", expand=True, pady=4)
        self.line_tree.bind("<Double-1>", self._on_double_click_line)

        override_frame = ttk.Frame(right)
        override_frame.pack(fill="x", pady=4)
        ttk.Label(override_frame, text="Sửa mã hàng cho dòng đang chọn:").pack(side="left")
        self.override_var = tk.StringVar()
        ttk.Entry(override_frame, textvariable=self.override_var, width=20).pack(side="left", padx=4)
        ttk.Button(override_frame, text="Áp dụng", command=self._apply_override).pack(side="left")
        ttk.Button(
            override_frame, text="Đặt lại trạng thái -> Sẵn sàng (chạy lại)",
            command=self._reset_row_status,
        ).pack(side="left", padx=(12, 0))
        ttk.Button(
            override_frame, text="Gỡ 'Đã cất' (đã xóa hóa đơn này trên MISA)",
            command=self._unmark_da_cat,
        ).pack(side="left", padx=(12, 0))

        totals_frame = ttk.LabelFrame(right, text="Tổng & quy tắc nghiệp vụ cố định", padding=8)
        totals_frame.pack(fill="x", pady=8)
        self.totals_label = ttk.Label(totals_frame, text="", justify="left")
        self.totals_label.pack(anchor="w")

    def _build_log_area(self):
        frame = ttk.LabelFrame(self, text="Log", padding=4)
        frame.pack(side="bottom", fill="both", expand=False)
        self.log_text = tk.Text(frame, height=8, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)

    # ------------------------------------------------------------------ #
    def _log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _choose_pdf_file(self):
        cur = self.pdf_dir.get().strip()
        start = cur if os.path.isdir(cur) else (os.path.dirname(cur) or os.getcwd())
        picker = SimplePathPicker(
            self, "Chọn 1 file PDF hóa đơn", start, select_file=True, extensions=(".pdf",)
        )
        self.wait_window(picker)
        if picker.result:
            self.pdf_dir.set(picker.result)
            self._log(f"Đã chọn file PDF: {picker.result}")

    def _choose_pdf_dir(self):
        start = self.pdf_dir.get().strip() or os.getcwd()
        if os.path.isfile(start):
            start = os.path.dirname(start)
        picker = SimplePathPicker(self, "Chọn thư mục chứa nhiều PDF hóa đơn", start, select_file=False)
        self.wait_window(picker)
        if picker.result:
            self.pdf_dir.set(picker.result)
            self._log(f"Đã chọn thư mục PDF: {picker.result}")

    def _choose_catalog(self):
        start = os.path.dirname(self.catalog_path.get().strip()) or os.getcwd()
        picker = SimplePathPicker(
            self, "Chọn file danh mục hàng hóa MISA", start,
            select_file=True, extensions=(".csv", ".xlsx", ".xls"),
        )
        self.wait_window(picker)
        if picker.result:
            self.catalog_path.set(picker.result)
            self._log(f"Đã chọn danh mục: {picker.result}")

    # ------------------------------------------------------------------ #
    def _load_and_match(self):
        pdf_path = self.pdf_dir.get().strip()
        if not pdf_path:
            messagebox.showerror("Lỗi", "Chọn 1 file PDF hoặc 1 thư mục chứa PDF trước.")
            return
        if os.path.isfile(pdf_path):
            paths = [pdf_path] if pdf_path.lower().endswith(".pdf") else []
        elif os.path.isdir(pdf_path):
            paths = sorted(glob.glob(os.path.join(pdf_path, "*.pdf")))
        else:
            messagebox.showerror("Lỗi", f"Đường dẫn không tồn tại: {pdf_path}")
            return
        if not paths:
            messagebox.showwarning("Không có file", "Không tìm thấy file PDF hợp lệ.")
            return

        try:
            if self.catalog_path.get().strip():
                self.matcher = ProductMatcher(self.catalog_path.get().strip())
                self._log(f"Đã nạp danh mục hàng hóa: {self.catalog_path.get()}")
            else:
                self.matcher = build_sample_matcher()
                self._log("CHƯA chọn danh mục thật -> dùng danh mục MẪU (6 mã) để test.")
        except Exception as e:
            messagebox.showerror("Lỗi đọc danh mục", str(e))
            return

        self.rows = []
        for p in paths:
            try:
                for inv in extract_invoices(p):
                    self.rows.append(InvoiceRow(inv))
            except Exception as e:
                self._log(f"Lỗi đọc file {p}: {e}")

        self.rows.sort(key=lambda r: int(r.invoice.so_hoa_don) if r.invoice.so_hoa_don.isdigit() else 0)

        da_cat_set = _load_da_cat_set()
        so_bo_qua = 0
        for row in self.rows:
            row.matches = [self.matcher.match(l.ten_hang) for l in row.invoice.lines]
            if row.invoice.so_hoa_don in da_cat_set:
                row.status = "da_cat"
                so_bo_qua += 1
            elif row.invoice.warnings:
                row.status = "can_duyet"
            elif row.can_auto_process:
                row.status = "san_sang"
            else:
                row.status = "can_duyet"

        self._log(f"Đã đọc {len(self.rows)} hóa đơn từ {len(paths)} file PDF.")
        if so_bo_qua:
            self._log(
                f"{so_bo_qua} hóa đơn đã có trong log 'đã Cất' ({DA_CAT_LOG_PATH}) "
                "-> đánh dấu sẵn 'Đã cất', sẽ tự bỏ qua khi chạy (tránh nhập trùng)."
            )
        self._refresh_tree()

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.rows):
            inv = row.invoice
            self.tree.insert(
                "", "end", iid=str(i),
                values=(
                    inv.so_hoa_don, inv.ngay, len(inv.lines),
                    f"{inv.tong_thanh_toan:,.0f}", self.STATUS_LABEL[row.status],
                ),
            )

    # ------------------------------------------------------------------ #
    def _on_select_invoice(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        row = self.rows[idx]
        self.selected_row = row
        self._render_detail(row)

    def _render_detail(self, row: InvoiceRow):
        inv = row.invoice
        self.detail_header.configure(
            text=f"Hóa đơn số {inv.so_hoa_don} | Ngày {inv.ngay} | Người mua: {inv.nguoi_mua} (sẽ nhập = Khách lẻ)"
        )
        if inv.warnings:
            self.warning_label.configure(text="⚠ Cảnh báo đối chiếu số liệu:\n" + "\n".join(inv.warnings))
        else:
            self.warning_label.configure(text="")

        self.line_tree.delete(*self.line_tree.get_children())
        for i, (line, m) in enumerate(zip(inv.lines, row.matches)):
            ma_final = row.resolved_ma_hang(i)
            score_display = f"{m.score:.0f}" if not (i in row.overrides) else "tay"
            self.line_tree.insert(
                "", "end", iid=str(i),
                values=(
                    line.ten_hang, f"{line.so_luong:g}", f"{line.don_gia:,.0f}",
                    m.ten_hang_misa or "(không khớp)", score_display,
                    ma_final or "(chưa có — cần chọn tay)", f"{line.tien_thue:,.0f}",
                ),
                tags=("low_conf",) if (not m.can_auto_accept and i not in row.overrides) else (),
            )
        self.line_tree.tag_configure("low_conf", background="#fff3cd")

        self.totals_label.configure(
            text=(
                f"Khách hàng = Khách lẻ (cố định)   |   Trạng thái = Chưa thu tiền (cố định)\n"
                f"%VAT trong MISA = 10% (cố định, mọi dòng)   |   Tiền thuế mỗi dòng = ghi đè theo số thật trên PDF\n"
                f"Tổng tiền hàng: {inv.tong_tien_truoc_thue:,.0f}   |   "
                f"Tổng tiền thuế: {inv.tong_tien_thue:,.0f}   |   "
                f"Tổng thanh toán: {inv.tong_thanh_toan:,.0f}\n"
                f"Trạng thái xử lý: {self.STATUS_LABEL[row.status]}"
            )
        )

    def _on_double_click_line(self, _event=None):
        sel = self.line_tree.selection()
        if not sel:
            return
        self.override_var.set("")

    def _apply_override(self):
        if not self.selected_row:
            return
        sel = self.line_tree.selection()
        if not sel:
            messagebox.showinfo("Chọn dòng", "Chọn 1 dòng hàng trong bảng trước.")
            return
        idx = int(sel[0])
        new_code = self.override_var.get().strip()
        if not new_code:
            messagebox.showinfo("Thiếu mã", "Nhập mã hàng muốn dùng cho dòng này.")
            return
        self.selected_row.overrides[idx] = new_code
        # KHÔNG dùng self.selected_row.can_auto_process nữa để quyết định trạng thái —
        # thuộc tính đó chấm điểm dựa trên ProductMatcher đọc từ file CSV cũ (chỉ còn
        # dùng để XEM TRƯỚC cho đẹp), không phản ánh đúng việc tra cứu SỐNG trong MISA
        # lúc chạy thật (từ 29/08/2026). Người dùng đã tự sửa tay dòng này -> coi là sẵn
        # sàng chạy lại, miễn hóa đơn không có cảnh báo đối chiếu số liệu PDF.
        if not self.selected_row.invoice.warnings:
            self.selected_row.status = "san_sang"
        self._render_detail(self.selected_row)
        self._refresh_tree()
        self._log(f"Đã sửa tay mã hàng dòng {idx} -> {new_code} (hóa đơn {self.selected_row.invoice.so_hoa_don})")

    def _reset_row_status(self):
        """
        Cho phép người dùng CHỦ ĐỘNG đặt lại trạng thái hóa đơn đang chọn về "Sẵn sàng"
        để bấm "Nhập THẬT..." chạy lại — dùng sau khi đã sửa tay mã hàng cho (các) dòng
        bị "Cần duyệt tay", hoặc đơn giản muốn thử lại 1 hóa đơn từng bị dừng. Hóa đơn đã
        "Đã cất" thật thì KHÔNG cho đặt lại bằng nút này (tránh bấm nhầm dẫn tới nhập
        trùng vào sổ sách thật — muốn nhập lại hóa đơn đã Cất phải tự xoá dòng số hóa đơn
        đó khỏi da_cat_log.json, đây là rào chắn có chủ đích).
        """
        if not self.selected_row:
            messagebox.showinfo("Chưa chọn hóa đơn", "Chọn 1 hóa đơn trong hàng đợi trước.")
            return
        if self.selected_row.status == "da_cat":
            messagebox.showwarning(
                "Đã cất rồi",
                "Hóa đơn này đã Cất thành công vào MISA thật — không cho đặt lại trạng "
                "thái ở đây để tránh nhập trùng. Nếu bạn đã tự xóa chứng từ này khỏi MISA "
                "và muốn nhập lại, dùng nút 'Gỡ Đã cất (đã xóa hóa đơn này trên MISA)' "
                "ngay bên cạnh.",
            )
            return
        if self.selected_row.invoice.warnings:
            messagebox.showwarning(
                "Còn cảnh báo đối chiếu số liệu",
                f"Hóa đơn còn cảnh báo đối chiếu số liệu từ PDF: {self.selected_row.invoice.warnings}\n"
                "Cần xử lý cảnh báo này trước (không liên quan tới mã hàng), không thể đặt lại Sẵn sàng.",
            )
            return
        self.selected_row.status = "san_sang"
        self._render_detail(self.selected_row)
        self._refresh_tree()
        self._log(f"Đã đặt lại trạng thái hóa đơn {self.selected_row.invoice.so_hoa_don} -> Sẵn sàng (có thể chạy lại).")

    def _unmark_da_cat(self):
        """
        Gỡ 1 Số hóa đơn khỏi da_cat_log.json (file trên đĩa dùng để chống nhập trùng) —
        CHỈ dùng khi người dùng đã tự tay XÓA chứng từ tương ứng khỏi MISA thật (VD nhập
        sai cần làm lại từ đầu). Có hộp thoại xác nhận rõ ràng vì bấm nhầm sẽ mở đường
        cho lần chạy tự động tiếp theo Cất TRÙNG hóa đơn vẫn còn nguyên trong MISA.
        """
        if not self.selected_row:
            messagebox.showinfo("Chưa chọn hóa đơn", "Chọn 1 hóa đơn trong hàng đợi trước.")
            return
        if self.selected_row.status != "da_cat":
            messagebox.showinfo("Không cần thiết", "Hóa đơn này chưa ở trạng thái 'Đã cất', không cần gỡ.")
            return
        so_hd = self.selected_row.invoice.so_hoa_don
        confirmed = messagebox.askyesno(
            "Xác nhận đã xóa trên MISA thật",
            f"Hóa đơn {so_hd} đang được đánh dấu 'Đã cất' vì đã Cất thành công vào MISA "
            "thật trước đó.\n\n"
            "Bạn XÁC NHẬN đã tự xóa chứng từ này khỏi MISA (Bán hàng ▸ Chứng từ bán hàng) "
            "CHƯA?\n\n"
            "Nếu CHƯA xóa mà vẫn bấm Yes ở đây, lần chạy tự động tiếp theo sẽ Cất THÊM "
            "1 bản ghi TRÙNG LẶP vào sổ sách thật.\n\n"
            "Chỉ bấm Yes nếu chắc chắn đã xóa xong trên MISA.",
            icon="warning",
        )
        if not confirmed:
            return
        _remove_da_cat_log(so_hd)
        self.selected_row.status = "can_duyet" if self.selected_row.invoice.warnings else "san_sang"
        self._render_detail(self.selected_row)
        self._refresh_tree()
        self._log(
            f"Đã gỡ trạng thái 'Đã cất' cho hóa đơn {so_hd} (xác nhận đã xóa trên MISA thật) "
            "-> có thể chạy lại."
        )

    # ------------------------------------------------------------------ #
    def _run_all(self, dry_run: bool, only_selected: bool = False):
        if not self.rows:
            messagebox.showwarning("Chưa có dữ liệu", "Bấm 'Đọc & đối chiếu' trước.")
            return

        target_rows = self.rows
        if only_selected:
            sel = self.tree.selection()
            if not sel:
                messagebox.showinfo("Chưa chọn", "Chọn 1 hóa đơn trong danh sách bên trái trước.")
                return
            target_rows = [self.rows[int(sel[0])]]

        if not dry_run:
            so_hd_list = ", ".join(r.invoice.so_hoa_don for r in target_rows)
            confirm = messagebox.askyesno(
                "XÁC NHẬN GHI DỮ LIỆU THẬT",
                f"Thao tác này sẽ THAO TÁC THẬT trên MISA SME.NET 2019 đang mở "
                f"(bấm Thêm, điền dữ liệu, và CẤT — ghi thẳng vào sổ sách kế toán).\n\n"
                f"Hóa đơn sẽ nhập: {so_hd_list}\n\n"
                "Yêu cầu bắt buộc:\n"
                "- Script đang chạy dưới quyền Administrator (khớp quyền MISA).\n"
                "- MISA đã mở sẵn, đúng đơn vị, đúng màn hình.\n\n"
                "Bạn có chắc chắn muốn tiếp tục?",
                icon="warning",
            )
            if not confirm:
                return

        thread = threading.Thread(target=self._run_all_worker, args=(dry_run, target_rows), daemon=True)
        thread.start()

    def _run_all_worker(self, dry_run: bool, target_rows: list | None = None):
        rows_to_run = target_rows if target_rows is not None else self.rows
        automation = MisaAutomation(dry_run=dry_run)
        if not dry_run:
            try:
                automation.connect()
                automation.ensure_ban_hang_list_open()
            except Exception as e:
                self._log(f"LỖI kết nối/điều hướng MISA ({type(e).__name__}):\n" + traceback.format_exc())
                return

        # ---- Từ 29/08/2026: quyết định mã hàng chuyển sang tra cứu TRỰC TIẾP trong
        # MISA khi đang nhập (gõ tên hàng vào ô "Mã hàng", MISA tự tra cứu + Tab để
        # chốt), không còn phụ thuộc file danh mục CSV/matcher tính trước nữa. Vì vậy
        # CHỈ bỏ qua trước những hóa đơn có `inv.warnings` (lỗi đối chiếu số liệu đọc
        # từ PDF — quy tắc #6, không liên quan gì tới mã hàng). Độ tin cậy mã hàng giờ
        # được kiểm tra NGAY LÚC nhập thật (per dòng, so khớp với kết quả MISA trả về)
        # — hóa đơn nào bị chặn giữa chừng sẽ tự chuyển "can_duyet" như cũ.
        skip_count = sum(1 for r in rows_to_run if r.invoice.warnings)
        if skip_count:
            self._log(f"{skip_count} hóa đơn có cảnh báo đối chiếu số liệu -> sẽ bỏ qua, không tự nhập.")

        for row in rows_to_run:
            inv = row.invoice
            # An toàn khi CHẠY LẠI sau khi 1 batch bị dừng giữa chừng (VD lỗi ở hóa đơn
            # thứ N) — bỏ qua các hóa đơn ĐÃ Cất thành công ở lần chạy trước, tránh nhập
            # trùng 2 lần vào sổ sách thật.
            if row.status == "da_cat" or inv.so_hoa_don in _load_da_cat_set():
                # Kiểm tra lại LOG TRÊN ĐĨA (không chỉ row.status trong bộ nhớ) ngay
                # trước khi xử lý — chặn cả trường hợp app vừa mới khởi động lại và
                # danh sách hàng đợi chưa kịp đối chiếu qua log (VD chạy "chỉ hóa đơn
                # đang chọn" ngay sau khi đọc PDF).
                row.status = "da_cat"
                self._log(f"Bỏ qua hóa đơn {inv.so_hoa_don} — đã Cất thành công từ trước (tránh nhập trùng).")
                self.after(0, self._refresh_tree)
                continue
            if inv.warnings:
                row.status = "can_duyet"
                self._log(f"Bỏ qua hóa đơn {inv.so_hoa_don} — cảnh báo đối chiếu số liệu: {inv.warnings}")
                self.after(0, self._refresh_tree)
                continue

            lines = [
                LineToEnter(
                    ma_hang=row.overrides.get(i, l.ten_hang),  # text để MISA tra cứu (tên hàng PDF, hoặc override tay)
                    so_luong=l.so_luong, don_gia=l.don_gia, tien_thue_dong=l.tien_thue,
                )
                for i, l in enumerate(inv.lines)
            ]
            payload = InvoiceToEnter(
                ngay=inv.ngay, lines=lines, tien_thue_that=inv.tong_tien_thue,
                tong_tien_hang=inv.tong_tien_truoc_thue, tong_thanh_toan=inv.tong_thanh_toan,
                so_hoa_don_goc=inv.so_hoa_don,
            )
            popup = None
            try:
                if not dry_run:
                    automation.ensure_ban_hang_list_open()
                popup = automation.open_new_invoice_popup()
                automation.fill_invoice(popup, payload, row.matches)
                automation.save(popup)
                row.status = "da_cat"
                if not dry_run:
                    _append_da_cat_log(inv.so_hoa_don)
                self._log(f"Đã cất hóa đơn {inv.so_hoa_don}.")
            except LowConfidenceMatchError as e:
                row.status = "can_duyet"
                self._log(f"Hóa đơn {inv.so_hoa_don} BỎ QUA (mã hàng độ tin cậy thấp, cần duyệt tay): {e}")
                da_huy = True
                if not dry_run:
                    da_huy = automation.cancel_popup_safely(popup)
                self.after(0, self._refresh_tree)
                if not dry_run and not da_huy:
                    self._log(
                        "DỪNG toàn bộ vòng lặp — không tự hủy được form dở dang, cần bạn tự "
                        "kiểm tra/đóng (chọn 'Không lưu' nếu được hỏi) trước khi chạy tiếp."
                    )
                    break
                continue
            except Exception as e:
                row.status = "loi"
                row.error_msg = str(e)
                # Ghi ĐẦY ĐỦ traceback (không chỉ str(e)) — XÁC NHẬN THẬT (30/08/2026):
                # nhiều lần chẩn đoán lỗi bị chậm vì chỉ có thông báo lỗi ngắn gọn, thiếu
                # tên loại exception + dòng code gây lỗi, phải hỏi lại người dùng nhiều
                # lần mới xác định được nguyên nhân thật. Từ giờ log tự đủ chi tiết.
                self._log(
                    f"LỖI khi nhập hóa đơn {inv.so_hoa_don} ({type(e).__name__}):\n"
                    + traceback.format_exc()
                )
                automation.close_popup_if_open(popup)
                self.after(0, self._refresh_tree)
                if not dry_run:
                    self._log(
                        "DỪNG toàn bộ vòng lặp — form MISA đang để mở dở dang, cần bạn tự "
                        "kiểm tra/đóng (chọn 'Không lưu' nếu được hỏi) trước khi chạy tiếp."
                    )
                    break
                continue

            self.after(0, self._refresh_tree)

        self._log(f"=== Hoàn tất ({'DRY-RUN' if dry_run else 'THẬT'}) ===")


if __name__ == "__main__":
    app = App()
    app.mainloop()

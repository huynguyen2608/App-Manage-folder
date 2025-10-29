import os
import sys
import shutil
import re
import threading
import subprocess
import tempfile
from datetime import datetime

from PySide6.QtCore import Qt, QDir, QSize, Signal, QObject
from PySide6.QtGui import QPixmap, QImage, QIcon
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTreeView,
    QFileSystemModel, QHeaderView, QTextEdit, QSplitter, QFileDialog,
    QLineEdit, QListWidget, QListWidgetItem, QMessageBox, QSizePolicy,
    QScrollArea, QStyle, QApplication, QProgressBar
)

# Optional libraries
try:
    import fitz  # PyMuPDF for PDF rendering
except Exception:
    fitz = None

try:
    from docx import Document
except Exception:
    Document = None

try:
    import openpyxl
except Exception:
    openpyxl = None

# Optional converters
try:
    import docx2pdf  # only works on Windows where Word is available
except Exception:
    docx2pdf = None


# --- PDF rendering worker (emits QImage per page) ---
class PdfWorker(QObject):
    page_rendered = Signal(QImage)
    finished = Signal()
    error = Signal(str)

    def __init__(self, pdf_bytes, scale=1.0):
        super().__init__()
        self.pdf_bytes = pdf_bytes
        self.scale = scale

    def run(self):
        if fitz is None:
            self.error.emit("PyMuPDF (fitz) not installed")
            return
        try:
            doc = fitz.open(stream=self.pdf_bytes, filetype="pdf")
            for i in range(doc.page_count):
                page = doc.load_page(i)
                mat = fitz.Matrix(self.scale, self.scale)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                # Emit a copy to be safe
                self.page_rendered.emit(img.copy())
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


# --- DOCX rendering worker (emits PDF bytes) ---
class DocxToPdfWorker(QObject):
    pdf_ready = Signal(bytes)
    error = Signal(str)

    def __init__(self, docx_path: str, converter_func):
        super().__init__()
        self.docx_path = docx_path
        self.converter_func = converter_func 

    def run(self):
        try:
            pdf_bytes = self.converter_func(self.docx_path)
            if pdf_bytes:
                self.pdf_ready.emit(pdf_bytes)
            else:
                self.error.emit("Không tìm thấy công cụ chuyển đổi (MS Word/LibreOffice).")
        except Exception as e:
            self.error.emit(f"Lỗi chuyển đổi DOCX: {str(e)}")


# ------------------------ Custom File System Model ------------------------
class CustomFileSystemModel(QFileSystemModel):
    """Ẩn file tạm thời bắt đầu bằng '~$'."""
    def filterAcceptsRow(self, source_row, source_parent):
        index = self.index(source_row, 0, source_parent)
        file_name = self.fileName(index)

        # Ẩn file bắt đầu bằng "~$" (temporary/lock files)
        if file_name.startswith("~$"):
            return False

        return super().filterAcceptsRow(source_row, source_parent)


# ------------------------ Preview Widget ------------------------
class PreviewWidget(QWidget):
    """Hiển thị nội dung file: text, ảnh, pdf (từ bytes)."""

    def __init__(self):
        super().__init__()
        self.setStyleSheet("QWidget { background-color: white; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)

        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.text = QTextEdit(readOnly=True)

        self.img_label = QLabel(alignment=Qt.AlignCenter)
        self.img_scroll = QScrollArea()
        self.img_scroll.setWidgetResizable(True)
        self.img_scroll.setWidget(self.img_label)
        self.img_scroll.setStyleSheet("QScrollArea { border: 1px solid #ddd; }")

        self.pdf_scroll = QScrollArea()
        self.pdf_scroll.setWidgetResizable(True)
        self.pdf_container = QWidget()
        self.pdf_vlayout = QVBoxLayout(self.pdf_container)
        self.pdf_vlayout.setAlignment(Qt.AlignTop)
        self.pdf_scroll.setWidget(self.pdf_container)
        self.pdf_scroll.setStyleSheet("QScrollArea { border: 1px solid #ddd; }")

        layout.addWidget(self.text)
        layout.addWidget(self.img_scroll)
        layout.addWidget(self.pdf_scroll)

        # start hidden
        self.text.hide()
        self.img_scroll.hide()
        self.pdf_scroll.hide()

        self._pdf_worker_thread = None
        self._pdf_worker_obj = None

    def clear_pdf_pages(self):
        while self.pdf_vlayout.count():
            it = self.pdf_vlayout.takeAt(0)
            w = it.widget()
            if w:
                w.deleteLater()

    def show_text(self, s: str):
        self._stop_pdf()
        self.clear_pdf_pages()
        self.img_scroll.hide()
        self.pdf_scroll.hide()
        self.text.show()
        self.text.setPlainText(s)

    def show_image(self, path: str):
        self._stop_pdf()
        self.clear_pdf_pages()
        self.text.hide()
        self.pdf_scroll.hide()
        self.img_scroll.show()
        pix = QPixmap(path)
        if pix.isNull():
            self.show_text("Không thể hiển thị ảnh.")
            return
        max_w = 900
        if pix.width() > max_w:
            pix = pix.scaledToWidth(max_w, Qt.SmoothTransformation)
        self.img_label.setPixmap(pix)

    def show_pdf_bytes(self, pdf_bytes: bytes):
        """
        Render PDF (bytes) into images and show them.
        Uses PdfWorker which emits QImage pages.
        """
        self._stop_pdf()
        self.clear_pdf_pages()
        self.text.hide()
        self.img_scroll.hide()
        self.pdf_scroll.show()
        loading = QLabel("Đang tải PDF...")
        loading.setAlignment(Qt.AlignCenter)
        self.pdf_vlayout.addWidget(loading)

        worker = PdfWorker(pdf_bytes, scale=1.0)
        thread = threading.Thread(target=worker.run, daemon=True)

        def on_page(img: QImage):
            # remove loading label on first page rendered
            nonlocal loading
            if loading and loading.parent():
                loading.deleteLater()
                loading = None
            lbl = QLabel()
            lbl.setPixmap(QPixmap.fromImage(img))
            lbl.setAlignment(Qt.AlignCenter)
            self.pdf_vlayout.addWidget(lbl)
            QApplication.processEvents()

        def on_error(msg):
            self.show_text("Lỗi hiển thị PDF: " + msg)

        # connect signals (worker is QObject)
        worker.page_rendered.connect(on_page)
        worker.error.connect(on_error)

        self._pdf_worker_obj = worker
        self._pdf_worker_thread = thread
        thread.start()

    def _stop_pdf(self):
        # stop any previous rendering (we don't have cancel logic for the pure-thread worker,
        # but we clear UI and drop references)
        self.clear_pdf_pages()
        self._pdf_worker_obj = None
        self._pdf_worker_thread = None


# ------------------------ File Manager Widget ------------------------
class FileManagerWidget(QWidget):
    """Màn hình quản lý file"""

    def __init__(self, root_folder=r"D:\Client"):
        super().__init__()
        self.root_folder = root_folder if os.path.exists(root_folder) else QDir.rootPath()
        self.current_file = None

        # Thêm biến theo dõi worker DOCX
        self._docx_worker_thread = None
        self._docx_worker_obj = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        self.setStyleSheet("QWidget { background-color: #f7f9fc; }")

        style = QApplication.style()

        # --- Thanh tìm kiếm ---
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Tìm kiếm gần đúng tên file...")
        self.search_input.addAction(style.standardIcon(QStyle.SP_FileDialogContentsView), QLineEdit.LeadingPosition)
        self.search_input.textChanged.connect(self.on_search)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        self.search_results_list = QListWidget()
        self.search_results_list.setFixedHeight(160)
        self.search_results_list.hide()
        self.search_results_list.itemClicked.connect(self.select_from_search)
        layout.addWidget(self.search_results_list)

        # --- Nút thao tác file ---
        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("Thêm File")
        self.btn_add.setIcon(style.standardIcon(QStyle.SP_FileIcon))
        self.btn_add.clicked.connect(self.add_file)

        self.btn_delete = QPushButton("Xóa File/Folder")
        self.btn_delete.setIcon(style.standardIcon(QStyle.SP_TrashIcon))
        self.btn_delete.clicked.connect(self.delete_file)

        self.btn_root = QPushButton("Đổi Folder Gốc")
        self.btn_root.setIcon(style.standardIcon(QStyle.SP_DirHomeIcon))
        self.btn_root.clicked.connect(self.change_root)

        self.btn_refresh = QPushButton()
        self.btn_refresh.setIcon(style.standardIcon(QStyle.SP_BrowserReload))
        self.btn_refresh.setFixedWidth(35)
        self.btn_refresh.clicked.connect(self.refresh_tree)

        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_delete)
        btn_row.addWidget(self.btn_root)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_refresh)
        layout.addLayout(btn_row)

        # --- Cây thư mục & khu vực xem trước ---
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        self.model = CustomFileSystemModel()
        self.model.setRootPath(self.root_folder)
        self.model.setFilter(QDir.AllEntries | QDir.NoDotAndDotDot)

        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(self.root_folder))
        self.tree.hideColumn(1)
        self.tree.hideColumn(2)
        self.tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tree.header().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.tree.clicked.connect(self.on_tree_clicked)
        splitter.addWidget(self.tree)

        # --- Preview ---
        preview_box = QWidget()
        preview_layout = QVBoxLayout(preview_box)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        header = QHBoxLayout()
        self.lbl_filename = QLabel("Chọn một tệp để xem trước")
        header.addWidget(self.lbl_filename)
        header.addStretch()
        self.btn_open = QPushButton("Mở File")
        self.btn_open.setIcon(style.standardIcon(QStyle.SP_DialogOpenButton))
        self.btn_open.setEnabled(False)
        self.btn_open.clicked.connect(self.open_file)
        header.addWidget(self.btn_open)
        preview_layout.addLayout(header)

        self.preview = PreviewWidget()
        preview_layout.addWidget(self.preview)

        self.lbl_file_info = QLabel("Ngày: N/A")
        self.lbl_file_info.setStyleSheet("color: gray; font-size: 9pt; padding: 4px; border-top: 1px solid #ccc;")
        self.lbl_file_info.setAlignment(Qt.AlignRight)
        self.lbl_file_info.hide()
        preview_layout.addWidget(self.lbl_file_info)
        splitter.addWidget(preview_box)

        splitter.setSizes([400, 800])

    # ------------------------ Helpers: docx -> pdf bytes ------------------------
    def _convert_docx_to_pdf_bytes(self, docx_path: str) -> bytes | None:
        """
        Try multiple strategies to convert a .docx to PDF bytes:
        1. docx2pdf (Windows + MS Word installed)
        2. LibreOffice / soffice command-line conversion (cross-platform if installed)
        Returns PDF bytes on success, or None on failure.
        """
        # 1) Try docx2pdf (Windows / MS Word)
        try:
            if docx2pdf and sys.platform.startswith("win"):
                tmp_out = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                tmp_out.close()
                try:
                    # docx2pdf.convert accepts input, output paths
                    docx2pdf.convert(docx_path, tmp_out.name)
                    with open(tmp_out.name, "rb") as f:
                        data = f.read()
                    return data
                finally:
                    try:
                        os.unlink(tmp_out.name)
                    except Exception:
                        pass
        except Exception:
            # ignore and fallback
            pass

        # 2) Try LibreOffice / soffice
        try:
            # create temp dir for output
            with tempfile.TemporaryDirectory() as tmpdir:
                # Use soffice or libreoffice
                soffice_cmds = ["soffice", "libreoffice"]
                chosen = None
                for cmd in soffice_cmds:
                    if shutil.which(cmd):
                        chosen = cmd
                        break
                if chosen is None:
                    # no soffice found
                    raise EnvironmentError("LibreOffice/soffice not found")

                # run conversion: soffice --headless --convert-to pdf --outdir tmpdir docx_path
                subprocess.run([chosen, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, docx_path],
                               check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                base = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
                out_pdf = os.path.join(tmpdir, base)
                if os.path.exists(out_pdf):
                    with open(out_pdf, "rb") as f:
                        data = f.read()
                    return data
        except Exception:
            pass

        # 3) Fallback: cannot convert
        return None

    # ------------------------ Các hành động ------------------------
    def _stop_docx_conversion(self):
        """Dừng/dọn dẹp các worker chuyển đổi DOCX đang chạy."""
        # Dù không thể hủy luồng, chúng ta xóa các tham chiếu worker để tránh kết nối tín hiệu
        self._docx_worker_obj = None
        self._docx_worker_thread = None

    def on_tree_clicked(self, index):
        # Dừng mọi luồng chuyển đổi DOCX cũ trước khi xử lý mới
        self._stop_docx_conversion()

        file_path = self.model.filePath(index)
        if not file_path:
            return

        if os.path.isdir(file_path):
            self.lbl_filename.setText(os.path.basename(file_path) or file_path)
            self.preview.show_text("Thư mục được chọn.")
            self.lbl_file_info.hide()
            return

        ext = os.path.splitext(file_path)[1].lower()
        self.lbl_filename.setText(os.path.basename(file_path))
        self.lbl_file_info.show()
        self.current_file = file_path
        self.btn_open.setEnabled(True)

        try:
            m_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            self.lbl_file_info.setText(f"Ngày lưu cuối: {m_time:%d/%m/%Y %H:%M}")
        except Exception:
            self.lbl_file_info.hide()

        # TEXT
        if ext in (".txt", ".py", ".csv", ".log"):
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                self.preview.show_text(f.read())

        # IMAGE
        elif ext in (".jpg", ".png", ".jpeg", ".bmp", ".gif"):
            self.preview.show_image(file_path)

        # PDF
        elif ext == ".pdf":
            # read bytes and preview
            try:
                with open(file_path, "rb") as f:
                    pdf_bytes = f.read()
                self.preview.show_pdf_bytes(pdf_bytes)
            except Exception as e:
                self.preview.show_text(f"Lỗi đọc PDF: {e}")

        # DOCX -> convert to PDF then preview PDF (preserve formatting)
        elif ext == ".docx":
            # Hiển thị trạng thái Loading trước
            self.preview.show_text("Đang chuyển đổi DOCX sang PDF... Vui lòng chờ.")
            
            # Khởi tạo Worker chuyển đổi trong luồng riêng
            worker = DocxToPdfWorker(file_path, self._convert_docx_to_pdf_bytes)
            thread = threading.Thread(target=worker.run, daemon=True)

            def on_pdf_ready(pdf_bytes: bytes):
                # Khi PDF bytes đã sẵn sàng, hiển thị nó
                self.preview.show_pdf_bytes(pdf_bytes)

            def on_error(msg):
                # Nếu chuyển đổi thất bại, thử hiển thị văn bản thô
                if Document:
                    try:
                        doc = Document(file_path)
                        self.preview.show_text(f"Chuyển đổi PDF thất bại ({msg}). Đang hiển thị văn bản thô:\n\n" + "\n\n".join(p.text for p in doc.paragraphs))
                    except Exception as e:
                        self.preview.show_text(f"Chuyển đổi PDF thất bại: {msg}\nKhông thể đọc văn bản thô: {e}")
                else:
                    self.preview.show_text(f"Chuyển đổi PDF thất bại: {msg}\n(Thiếu thư viện python-docx để đọc văn bản thô).")

            # Kết nối tín hiệu
            worker.pdf_ready.connect(on_pdf_ready)
            worker.error.connect(on_error)

            self._docx_worker_obj = worker
            self._docx_worker_thread = thread
            thread.start()

        # EXCEL preview as text (existing behavior)
        elif ext in (".xlsx", ".xls") and openpyxl:
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                text = "\n".join("\t".join(str(c or "") for c in r) for r in ws.iter_rows(values_only=True))
                self.preview.show_text(text)
            except Exception as e:
                self.preview.show_text(f"Lỗi đọc Excel: {e}")

        # OTHER
        else:
            self.preview.show_text("Không hỗ trợ định dạng này.")


    def open_file(self):
        if not self.current_file:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.current_file)
            elif sys.platform == "darwin":
                subprocess.call(["open", self.current_file])
            else:
                subprocess.call(["xdg-open", self.current_file])
        except Exception as e:
            QMessageBox.warning(self, "Lỗi", str(e))

    def on_search(self):
        q = self.search_input.text().strip()
        self.search_results_list.clear()
        if not q:
            self.search_results_list.hide()
            return

        results = []
        for root, _, files in os.walk(self.root_folder):
            for f in files:
                if q.lower() in f.lower():
                    results.append(os.path.join(root, f))

        if not results:
            self.search_results_list.hide()
            return

        for path in results:
            item = QListWidgetItem(os.path.basename(path))
            item.setData(Qt.UserRole, path)
            self.search_results_list.addItem(item)
        self.search_results_list.show()

    def select_from_search(self, item):
        file_path = item.data(Qt.UserRole)
        if os.path.exists(file_path):
            idx = self.model.index(file_path)
            if idx.isValid():
                self.tree.setCurrentIndex(idx)
                self.tree.scrollTo(idx)
                self.on_tree_clicked(idx)
        self.search_results_list.hide()

    def add_file(self):
        idx = self.tree.currentIndex()
        target = self.model.filePath(idx) if idx.isValid() else self.root_folder
        if not os.path.isdir(target):
            target = os.path.dirname(target)
        src, _ = QFileDialog.getOpenFileName(self, "Chọn File để thêm", "", "Tất cả (*.*)")
        if not src:
            return
        dst = os.path.join(target, os.path.basename(src))
        if os.path.exists(dst):
            if QMessageBox.question(self, "Ghi đè", f"File {dst} đã tồn tại. Ghi đè?") != QMessageBox.Yes:
                return
        shutil.copy2(src, dst)
        self.refresh_tree()

    def delete_file(self):
        idx = self.tree.currentIndex()
        if not idx.isValid():
            QMessageBox.warning(self, "Lỗi", "Chưa chọn file hoặc thư mục.")
            return
        path = self.model.filePath(idx)
        if not os.path.exists(path):
            return
        if QMessageBox.question(self, "Xác nhận", f"Xóa '{path}'?") != QMessageBox.Yes:
            return
        try:
            shutil.rmtree(path) if os.path.isdir(path) else os.remove(path)
            self.refresh_tree()
        except Exception as e:
            QMessageBox.critical(self, "Lỗi Xóa", str(e))

    def change_root(self):
        new_root = QFileDialog.getExistingDirectory(self, "Chọn Thư mục Gốc", self.root_folder)
        if new_root:
            self.root_folder = new_root
            self.model.setRootPath(new_root)
            self.tree.setRootIndex(self.model.index(new_root))

    def refresh_tree(self):
        root = self.model.rootPath()
        self.model.setRootPath(root)
        self.tree.setRootIndex(self.model.index(root))
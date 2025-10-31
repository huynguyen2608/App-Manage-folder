import os
import sys
import shutil
import re
import threading
import subprocess
import tempfile
import os.path
from datetime import datetime
from typing import List, Dict, Any # THÊM import typing

from PySide6.QtCore import Qt, QDir, QSize, Signal, QObject, QThread, QModelIndex # THÊM QThread, QObject, Signal
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
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        if fitz is None:
            self.error.emit("PyMuPDF (fitz) chưa được cài đặt. Không thể xem trước PDF.")
            self.finished.emit()
            return
            
        try:
            # Sử dụng PyMuPDF để xử lý PDF từ bytes
            doc = fitz.open(stream=self.pdf_bytes, filetype="pdf")
            
            for i in range(doc.page_count):
                if self._is_cancelled:
                    break

                page = doc.load_page(i)
                # Render page to Pixmap/Image
                pix = page.get_pixmap(matrix=fitz.Matrix(self.scale, self.scale))
                
                # Chuyển đổi Pixmap sang QImage
                img_format = QImage.Format.Format_RGB32 if pix.alpha else QImage.Format.Format_RGB888
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, img_format)
                
                self.page_rendered.emit(img)
                QApplication.processEvents() # Cho phép UI cập nhật

            doc.close()
        except Exception as e:
            self.error.emit(f"Lỗi khi đọc file PDF: {e}")
        finally:
            self.finished.emit()

# --- NEW: Loading Overlay ---
class LoadingOverlay(QWidget):
    """Một lớp overlay đơn giản hiển thị thông báo 'Loading'."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        # Không dùng FramelessWindowHint vì nó là child của FileManagerWidget
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Container cho thông báo loading
        loading_container = QWidget()
        loading_container.setStyleSheet("background-color: rgba(0, 0, 0, 150); border-radius: 10px;")
        
        container_layout = QVBoxLayout(loading_container)
        
        self.label = QLabel("Đang tải dữ liệu, vui lòng chờ...")
        self.label.setStyleSheet("color: white; font-size: 16pt; padding: 20px;")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0) # Chế độ busy/indeterminate
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar { 
                border: 1px solid white; 
                border-radius: 5px; 
                text-align: center; 
                height: 10px;
                margin: 0 50px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
        """)
        
        container_layout.addWidget(self.label, alignment=Qt.AlignmentFlag.AlignCenter)
        container_layout.addWidget(self.progress_bar)
        
        main_layout.addWidget(loading_container, alignment=Qt.AlignmentFlag.AlignCenter)
        
    def show_overlay(self):
        # Thiết lập kích thước bằng với widget cha
        if self.parentWidget():
            self.setGeometry(self.parentWidget().rect())
        self.raise_()
        self.show()

    def hide_overlay(self):
        self.hide()
        
    def resizeEvent(self, event):
        # Đảm bảo overlay luôn che phủ toàn bộ widget cha khi cha thay đổi kích thước
        if self.parentWidget():
            self.setGeometry(self.parentWidget().rect())
        super().resizeEvent(event)


# --- NEW: File Scan Worker (Luồng chạy nền để quét file) ---
class FileScanWorker(QObject):
    """Quét thư mục gốc và lưu trữ thông tin file vào cache."""
    finished = Signal(list) # Gửi list thông tin file đã cache
    error = Signal(str)

    def __init__(self, root_path: str):
        super().__init__()
        self.root_path = root_path
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        file_cache: List[Dict[str, Any]] = []
        try:
            # os.walk là blocking, nhưng nó chạy trong QThread nên không block UI
            for root, _, files in os.walk(self.root_path):
                if self._is_cancelled:
                    return

                for file_name in files:
                    if self._is_cancelled:
                        return
                    
                    full_path = os.path.join(root, file_name)
                    
                    # Chỉ lưu trữ các thông tin cần thiết để tìm kiếm và hiển thị
                    file_info = {
                        'name': file_name,
                        'path': full_path,
                        'lower_name': file_name.lower(),
                        'dir': root # Thư mục chứa file
                    }
                    file_cache.append(file_info)

            self.finished.emit(file_cache)
        except Exception as e:
            self.error.emit(f"Lỗi quét thư mục: {str(e)}")
            
            
# --- Custom Model & Preview Widgets (Giữ nguyên) ---
class CustomFileSystemModel(QFileSystemModel):
    # Giữ nguyên CustomFileSystemModel
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(False)

class PreviewWidget(QWidget):
    # Giữ nguyên PreviewWidget
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.image_preview = QLabel()
        self.image_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_preview.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.image_preview.setScaledContents(True)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.image_preview)

        self.layout.addWidget(self.text_preview)
        self.layout.addWidget(self.scroll_area)

        self.show_text("Chọn file để xem trước...")
        
        self.pdf_thread = None
        self.pdf_worker = None
        
    # Thêm hàm show_text, show_image, clear_preview (Giữ nguyên)
    def show_text(self, text):
        self.text_preview.setText(text)
        self.text_preview.show()
        self.scroll_area.hide()

    def show_image(self, image: QImage):
        # Dừng worker PDF cũ nếu có
        self.cancel_pdf_worker()
        
        # Chỉ hiển thị hình ảnh đầu tiên hoặc một hình ảnh duy nhất
        self.image_preview.setPixmap(QPixmap.fromImage(image))
        self.text_preview.hide()
        self.scroll_area.show()

    def show_multi_page_images(self, pdf_bytes, mime_type):
        """Xử lý xem trước PDF/DOCX (cần chuyển sang PDF) bằng luồng riêng."""
        self.cancel_pdf_worker()
        self.clear_preview()
        self.show_text("Đang tải xem trước...")
        
        # Sử dụng tạm file để chuyển DOCX sang PDF nếu cần
        temp_pdf_path = None
        
        if mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' and docx2pdf is not None:
             try:
                 temp_pdf_path = os.path.join(tempfile.gettempdir(), f"preview_{os.getpid()}_{datetime.now().microsecond}.pdf")
                 # Ghi docx bytes ra file tạm để docx2pdf có thể đọc
                 with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                    tmp_docx.write(pdf_bytes)
                    temp_docx_path = tmp_docx.name
                 
                 docx2pdf.convert(temp_docx_path, temp_pdf_path)
                 
                 # Đọc lại PDF bytes
                 with open(temp_pdf_path, 'rb') as f:
                     pdf_bytes = f.read()
                 
             except Exception as e:
                 self.show_text(f"Lỗi chuyển đổi DOCX sang PDF: {e}")
                 # Dọn dẹp file tạm
                 if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
                 if os.path.exists(temp_docx_path): os.remove(temp_docx_path)
                 return
             finally:
                 # Dọn dẹp file tạm docx
                 if os.path.exists(temp_docx_path): os.remove(temp_docx_path)


        self.pdf_thread = QThread()
        # Scale 1.5 cho hình ảnh sắc nét hơn
        self.pdf_worker = PdfWorker(pdf_bytes, scale=1.5)
        self.pdf_worker.moveToThread(self.pdf_thread)

        self.pdf_thread.started.connect(self.pdf_worker.run)
        self.pdf_worker.page_rendered.connect(self._add_page_to_preview)
        self.pdf_worker.finished.connect(self._handle_pdf_finished)
        self.pdf_worker.error.connect(self._handle_pdf_error)
        
        self.pdf_thread.start()

    def _add_page_to_preview(self, image: QImage):
        """Thêm hình ảnh trang vào layout của image_preview."""
        # Chuyển layout của image_preview sang QVBoxLayout để xếp chồng các trang
        if not isinstance(self.image_preview.parentWidget().layout(), QVBoxLayout):
            container = QWidget()
            old_layout = self.image_preview.parentWidget().layout()
            if old_layout:
                old_layout.removeWidget(self.image_preview.parentWidget())
            
            new_layout = QVBoxLayout(container)
            new_layout.setSpacing(10) # Khoảng cách giữa các trang
            new_layout.setContentsMargins(0, 0, 0, 0)
            new_layout.addWidget(self.image_preview)
            self.image_preview.setParent(container)
            self.scroll_area.setWidget(container)
            self.image_preview.hide() # Ẩn label cũ
            
        # Tạo QLabel mới cho trang hiện tại
        lbl = QLabel()
        lbl.setPixmap(QPixmap.fromImage(image))
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setMinimumWidth(image.width()) 
        
        # Thêm QLabel mới vào layout của container
        self.scroll_area.widget().layout().addWidget(lbl)
        
        self.text_preview.hide()
        self.scroll_area.show()

    def _handle_pdf_finished(self):
        # Dọn dẹp sau khi render xong
        self.pdf_thread.quit()
        self.pdf_thread.wait()
        
    def _handle_pdf_error(self, message: str):
        self.show_text(f"Lỗi xem trước đa trang: {message}")
        self.pdf_thread.quit()
        self.pdf_thread.wait()

    def cancel_pdf_worker(self):
        if self.pdf_worker and self.pdf_thread and self.pdf_thread.isRunning():
            self.pdf_worker.cancel()
            self.pdf_thread.quit()
            self.pdf_thread.wait()

    def clear_preview(self):
        # Dọn dẹp hình ảnh cũ
        self.cancel_pdf_worker()
        self.image_preview.clear()
        
        # Nếu đang ở chế độ nhiều trang, dọn dẹp các QLabel con
        if isinstance(self.scroll_area.widget().layout(), QVBoxLayout):
            layout = self.scroll_area.widget().layout()
            # Xóa tất cả widget con (các trang PDF)
            while layout.count() > 0:
                item = layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            # Đặt lại widget cho scroll area nếu cần (chỉ để tránh lỗi)
            self.image_preview = QLabel()
            self.image_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.image_preview.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
            self.image_preview.setScaledContents(True)
            self.scroll_area.setWidget(self.image_preview)

        self.show_text("Chọn file để xem trước...")

class DocxWorker(QObject):
    # Giữ nguyên DocxWorker
    finished = Signal(str)
    error = Signal(str)
    
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        if Document is None:
            self.error.emit("Thư viện 'python-docx' chưa được cài đặt.")
            return

        try:
            document = Document(self.file_path)
            full_text = []
            for para in document.paragraphs:
                full_text.append(para.text)
            self.finished.emit('\n'.join(full_text))
        except Exception as e:
            self.error.emit(f"Lỗi đọc file DOCX: {e}")

class ExcelWorker(QObject):
    # Giữ nguyên ExcelWorker
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        if openpyxl is None:
            self.error.emit("Thư viện 'openpyxl' chưa được cài đặt.")
            return

        try:
            workbook = openpyxl.load_workbook(self.file_path, read_only=True)
            output = ["--- Xem trước Excel ---"]
            # Chỉ xem trước sheet đầu tiên
            sheet = workbook.active
            
            for row in sheet.iter_rows(max_row=10, max_col=5): # Chỉ lấy 10 hàng đầu tiên, 5 cột đầu tiên
                row_data = [str(cell.value) if cell.value is not None else "" for cell in row]
                output.append('\t'.join(row_data))
                
            workbook.close()
            self.finished.emit('\n'.join(output))
        except Exception as e:
            self.error.emit(f"Lỗi đọc file Excel: {e}")


# --- FileManagerWidget (ĐÃ CẬP NHẬT) ---
class FileManagerWidget(QWidget):
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # --- File Manager Panel ---
        self.file_panel = QWidget()
        file_panel_layout = QVBoxLayout(self.file_panel)
        file_panel_layout.setContentsMargins(0, 0, 0, 0)

        # 1. Control Bar
        control_bar = QHBoxLayout()
        self.btn_root = QPushButton("Thư mục Gốc")
        self.btn_root.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.btn_root.clicked.connect(self.change_root)
        control_bar.addWidget(self.btn_root)
        
        self.btn_refresh = QPushButton("Làm mới")
        self.btn_refresh.setIcon(QApplication.style().standardIcon(QStyle.SP_BrowserReload))
        self.btn_refresh.clicked.connect(self.refresh_tree)
        control_bar.addWidget(self.btn_refresh)
        
        self.btn_add_file = QPushButton("Thêm File")
        self.btn_add_file.setIcon(QApplication.style().standardIcon(QStyle.SP_FileIcon))
        self.btn_add_file.clicked.connect(self.add_file)
        control_bar.addWidget(self.btn_add_file)

        self.btn_delete = QPushButton("Xóa")
        self.btn_delete.setIcon(QApplication.style().standardIcon(QStyle.SP_TrashIcon))
        self.btn_delete.clicked.connect(self.delete_file)
        control_bar.addWidget(self.btn_delete)
        
        file_panel_layout.addLayout(control_bar)
        
        # 2. Search Box
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Tìm kiếm file (trong cache)...")
        file_panel_layout.addWidget(self.search_box)
        
        # 3. Tree View (Hiển thị Cây thư mục)
        self.tree = QTreeView()
        self.tree.setHeaderHidden(True)
        self.tree.setSelectionMode(QTreeView.SelectionMode.SingleSelection)
        file_panel_layout.addWidget(self.tree)
        
        # 4. Search Results List (Hiển thị kết quả tìm kiếm từ Cache)
        self.search_results_list = QListWidget()
        self.search_results_list.setToolTip("Kết quả tìm kiếm từ bộ nhớ cache")
        self.search_results_list.setVisible(False) # Ẩn mặc định
        file_panel_layout.addWidget(self.search_results_list)

        # Kết nối tín hiệu
        self.tree.doubleClicked.connect(self.open_file_with_preview)
        self.search_results_list.itemDoubleClicked.connect(self.open_file_from_list)
        self.search_box.textChanged.connect(self._toggle_tree_list) # Chuyển đổi hiển thị
        self.search_box.textChanged.connect(self._perform_cache_search) # Kích hoạt tìm kiếm

        # --- Model Setup ---
        self.root_folder = os.path.abspath(QDir.rootPath()) # Mặc định là thư mục gốc
        self.current_file = None
        
        self.model = CustomFileSystemModel()
        # VÔ HIỆU HÓA LỌC MẶC ĐỊNH TRÊN MODEL:
        # self.model.setFilter(QDir.NoDotAndDotDot | QDir.AllEntries)
        # self.model.setRootPath(self.root_folder)
        
        # Vẫn dùng model để hiển thị cây thư mục cơ bản (vẫn nhanh), nhưng search sẽ dùng cache
        self.model.setRootPath(self.root_folder)
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(self.root_folder))

        # Ẩn các cột không cần thiết
        for i in range(1, self.model.columnCount()):
            self.tree.hideColumn(i)
        
        # --- NEW: Thuộc tính Cache và Worker ---
        self.file_cache: List[Dict[str, Any]] = []
        self.scan_worker: FileScanWorker | None = None
        self.scan_thread: QThread | None = None
        
        # --- NEW: Thêm Loading Overlay ---
        self.loading_overlay = LoadingOverlay(self)
        self.loading_overlay.hide() # Ẩn mặc định

        # Ẩn/Hiện List/Tree ban đầu
        self._toggle_tree_list("")

        # Bắt đầu tải dữ liệu ban đầu
        self.load_root_folder(self.root_folder) 

        # --- Preview Panel ---
        self.preview_panel = PreviewWidget()
        
        # --- Splitter (chia đôi) ---
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.file_panel)
        splitter.addWidget(self.preview_panel)
        splitter.setSizes([300, 700]) # Kích thước ban đầu

        main_layout.addWidget(splitter)
        
        # --- Message Label ---
        self.message_label = QLabel("Sẵn sàng.")
        self.message_label.setStyleSheet("padding: 2px; color: gray;")
        file_panel_layout.addWidget(self.message_label)


    # --- QUẢN LÝ CACHE VÀ THREADING ---

    def start_scan_worker(self, root_path: str):
        """Khởi tạo và chạy luồng quét file."""
        # Dọn dẹp luồng cũ nếu đang chạy
        if self.scan_worker and self.scan_thread and self.scan_thread.isRunning():
            self.scan_worker.cancel()
            self.scan_thread.quit()
            self.scan_thread.wait()

        self.scan_thread = QThread()
        self.scan_worker = FileScanWorker(root_path)
        
        self.scan_worker.moveToThread(self.scan_thread)

        self.scan_thread.started.connect(self.scan_worker.run)
        self.scan_worker.finished.connect(self._handle_scan_finished)
        self.scan_worker.error.connect(self._handle_scan_error)
        
        self.scan_thread.start()

    def _handle_scan_finished(self, file_cache: List[Dict[str, Any]]):
        """Xử lý khi worker quét file xong."""
        self.file_cache = file_cache
        self.loading_overlay.hide_overlay()
        self.show_message(f"Tải {len(self.file_cache)} file vào cache thành công.")
        
        if self.scan_thread:
            self.scan_thread.quit()
            self.scan_thread.wait()
        
        # Sau khi tải xong, chạy tìm kiếm nếu có từ khóa
        self._perform_cache_search(self.search_box.text())

    def _handle_scan_error(self, message: str):
        """Xử lý khi worker quét file gặp lỗi."""
        self.loading_overlay.hide_overlay()
        QMessageBox.critical(self, "Lỗi Quét File", message)
        if self.scan_thread:
            self.scan_thread.quit()
            self.scan_thread.wait()

    def _perform_cache_search(self, search_text: str):
        """Sử dụng cache để tìm kiếm và cập nhật ListWidget."""
        search_text = search_text.strip().lower()
        self.search_results_list.clear()

        if not search_text:
            self.search_results_list.addItem("Nhập từ khóa để tìm kiếm...")
            return

        if not self.file_cache:
            self.search_results_list.addItem("Đang tải cache... Vui lòng chờ.")
            return

        # Lọc kết quả từ cache (Search by name or containing directory)
        results = [
            info for info in self.file_cache
            if search_text in info['lower_name'] or search_text in info['dir'].lower()
        ]
        
        if not results:
            self.search_results_list.addItem(f"Không tìm thấy file nào cho '{search_text}'.")
            return

        # Thêm kết quả vào ListWidget
        for info in results:
            item = QListWidgetItem(info['name'])
            item.setToolTip(info['path'])
            item.setData(Qt.ItemDataRole.UserRole, info['path']) # Lưu path vào UserRole
            self.search_results_list.addItem(item)
            
        self.show_message(f"Tìm thấy {len(results)} kết quả.")

    def _toggle_tree_list(self, search_text):
        """Ẩn/Hiện Tree View hoặc List View dựa trên nội dung Search Box."""
        if search_text.strip():
            self.tree.setVisible(False)
            self.search_results_list.setVisible(True)
        else:
            self.tree.setVisible(True)
            self.search_results_list.setVisible(False)
            self.search_results_list.clear() # Dọn dẹp khi không dùng

    # --- HÀM HỆ THỐNG FILE ---

    def load_root_folder(self, new_root_folder: str):
        if not os.path.isdir(new_root_folder):
            QMessageBox.critical(self, "Lỗi", f"Thư mục không hợp lệ: {new_root_folder}")
            return
            
        self.root_folder = new_root_folder
        self.model.setRootPath(self.root_folder)
        self.tree.setRootIndex(self.model.index(self.root_folder))
        
        # Bắt đầu quét file và hiển thị loading
        self.loading_overlay.show_overlay() 
        self.start_scan_worker(self.root_folder)
        self.show_message(f"Đang quét thư mục: {os.path.basename(new_root_folder)}...")


    def refresh_tree(self):
        # Force refresh QFileSystemModel
        self.model.setRootPath(self.root_folder)
        # Tải lại cache
        self.load_root_folder(self.root_folder) 
        
    def add_file(self):
        idx = self.tree.currentIndex()
        if not idx.isValid():
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn thư mục đích.")
            return
        target = self.model.filePath(idx)
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
            self.load_root_folder(new_root)

    # --- HÀM XEM TRƯỚC VÀ MỞ FILE ---

    def open_file_with_preview(self, index: QModelIndex):
        """Mở file khi double click trên Tree View."""
        path = self.model.filePath(index)
        if os.path.isfile(path):
            self.current_file = path
            self.preview_file(path)

    def open_file_from_list(self, item: QListWidgetItem):
        """Mở file khi double click trên Search Results List."""
        path = item.data(Qt.ItemDataRole.UserRole)
        if path and os.path.isfile(path):
            self.current_file = path
            self.preview_file(path)

    def preview_file(self, path):
        self.preview_panel.clear_preview()
        
        # Mở file và lấy bytes
        try:
            with open(path, 'rb') as f:
                file_bytes = f.read()
        except Exception as e:
            self.preview_panel.show_text(f"Không thể đọc file: {e}")
            return
            
        file_extension = os.path.splitext(path)[1].lower()
        
        # Xử lý xem trước PDF/DOCX
        if file_extension in ['.pdf']:
            self.preview_panel.show_multi_page_images(file_bytes, 'application/pdf')
            
        elif file_extension in ['.docx', '.doc']:
             # Nếu có docx2pdf, cố gắng chuyển sang PDF để xem đa trang
             if docx2pdf is not None and fitz is not None:
                self.preview_panel.show_multi_page_images(file_bytes, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                return
             
             # Nếu không có docx2pdf, dùng DocxWorker để lấy text
             self.docx_thread = QThread()
             self.docx_worker = DocxWorker(path)
             self.docx_worker.moveToThread(self.docx_thread)
             
             self.docx_thread.started.connect(self.docx_worker.run)
             self.docx_worker.finished.connect(lambda text: self.preview_panel.show_text(text))
             self.docx_worker.error.connect(lambda msg: self.preview_panel.show_text(msg))
             self.preview_panel.show_text("Đang tải file Word...")
             self.docx_thread.start()
        
        # Xử lý xem trước TXT/CSV/CODE
        elif file_extension in ['.txt', '.csv', '.py', '.js', '.html', '.css', '.json']:
            try:
                # Cố gắng decode bằng utf-8
                text_content = file_bytes.decode('utf-8')
                self.preview_panel.show_text(text_content)
            except UnicodeDecodeError:
                # Nếu không được, thử latin-1
                text_content = file_bytes.decode('latin-1')
                self.preview_panel.show_text(f"--- Dữ liệu được decode bằng Latin-1 ---\n{text_content}")
            except Exception as e:
                self.preview_panel.show_text("Lỗi đọc file văn bản: " + str(e))
                
        # Xử lý xem trước Image
        elif file_extension in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
            image = QImage.fromData(file_bytes)
            self.preview_panel.show_image(image)
        
        # Xử lý xem trước Excel (XLSX, XLS)
        elif file_extension in ['.xlsx', '.xls']:
            self.excel_thread = QThread()
            self.excel_worker = ExcelWorker(path)
            self.excel_worker.moveToThread(self.excel_thread)
            
            self.excel_thread.started.connect(self.excel_worker.run)
            self.excel_worker.finished.connect(lambda text: self.preview_panel.show_text(text))
            self.excel_worker.error.connect(lambda msg: self.preview_panel.show_text(msg))
            self.preview_panel.show_text("Đang tải file Excel...")
            self.excel_thread.start()

        else:
            self.preview_panel.show_text("Không hỗ trợ xem trước định dạng này.")

    def show_message(self, message: str):
        """Hiển thị thông báo ở Message Label."""
        self.message_label.setText(message)

# Đặt LoadingOverlay là lớp con của FileManagerWidget
# Điều này được xử lý trong __init__ khi self.loading_overlay = LoadingOverlay(self)

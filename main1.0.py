# main.py
import sys
import os
import threading
import subprocess
import shutil 
from datetime import datetime
import re
from typing import List, Dict, Any

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTreeView, QFileSystemModel, QSplitter, QTextEdit, QLabel, QPushButton,
    QTabWidget, QMessageBox, QTableWidget, QTableWidgetItem, QLineEdit,
    QHeaderView, QSizePolicy, QScrollArea, QStyle, QFileDialog, QListWidget, QListWidgetItem 
)
from PySide6.QtCore import Qt, QDir, QSize, Signal, QObject, QModelIndex, QTimer, QRect 
from PySide6.QtGui import QPixmap, QImage, QIcon 

# optional libs
try:
    import fitz # PyMuPDF for PDF rendering
except Exception:
    fitz = None

try:
    from docx import Document
except Exception:
    Document = None

import openpyxl

# ----------------- Config -----------------
# THAY ĐỔI: Thư mục gốc được đọc từ D:\Client
ROOT_FOLDER = r"D:\Client" if os.path.exists(r"D:\Client") else (r"D:\\" if os.path.exists(r"D:\\") else QDir.rootPath())
DEFAULT_EXCEL_PATH = r"D:\Data\report.xlsx"
CONTRACT_TEMPLATE_PATH = r"D:\Template"
DEBOUNCE_DELAY_MS = 500 # Độ trễ 0.5 giây cho tính toán công nợ
# ------------------------------------------

# THAY ĐỔI: Headers cho cột công nợ (Giả định)
COL_G_HEADER = "GIÁ TRỊ NGHIỆM THU THEO TIẾN ĐỘ"
COL_H_HEADER = "KH THANH TOÁN TẠM ỨNG"
COL_J_HEADER = "CÒN PHẢI THU"

# --- Custom File System Model to hide temporary files ---
class CustomFileSystemModel(QFileSystemModel):
    def filterAcceptsRow(self, source_row, source_parent):
        index = self.index(source_row, 0, source_parent)
        file_name = self.fileName(index)
        
        # Ẩn file bắt đầu bằng "~$" (temporary/lock files)
        if file_name.startswith("~$"):
            return False
            
        return super().filterAcceptsRow(source_row, source_parent)

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
                self.page_rendered.emit(img.copy())
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

# --- Preview widget supporting text, image, pdf pages ---
class PreviewWidget(QWidget):
    def __init__(self):
        super().__init__()
        # THÊM: Đồng nhất màu nền
        self.setStyleSheet("QWidget { background-color: white; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4,4,4,4)
        
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.text = QTextEdit()
        self.text.setReadOnly(True)

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
            if loading and loading.parent():
                loading.deleteLater()
            lbl = QLabel()
            lbl.setPixmap(QPixmap.fromImage(img))
            lbl.setAlignment(Qt.AlignCenter)
            self.pdf_vlayout.addWidget(lbl)
            QApplication.processEvents()

        def on_error(msg):
            self.show_text("Lỗi hiển thị PDF: " + msg)

        worker.page_rendered.connect(on_page)
        worker.error.connect(on_error)

        self._pdf_worker_obj = worker
        self._pdf_worker_thread = thread
        thread.start()

    def _stop_pdf(self):
        self.clear_pdf_pages()

# --- Excel Editor widget ---
class ExcelEditor(QWidget):
    def __init__(self, excel_path: str):
        super().__init__()
        # THÊM: Đồng nhất màu nền
        self.setStyleSheet("QWidget { background-color: #f7f9fc; }")
        self.excel_path = excel_path
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6,6,6,6)

        self.sheets_tables = {} # Lưu trữ table theo tên sheet
        self.sheets_names = []
        self.col_map = {} # THÊM: Lưu trữ index cột cho tính toán công nợ
        self.debounce_timer = QTimer() # THÊM: Timer cho debounce
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.timeout.connect(self._perform_calculation)
        self.pending_calc_row = -1

        control_row = QHBoxLayout()
        self.msg_label = QLabel("")
        control_row.addWidget(self.msg_label)
        control_row.addStretch()
        
        # [NEW FEATURE] THÊM NÚT THAY ĐỔI FILE
        self.btn_change_file = QPushButton("Thay Đổi File")
        self.btn_change_file.setIcon(QApplication.style().standardIcon(QStyle.SP_FileDialogStart)) 
        self.btn_change_file.setFixedWidth(120)
        self.btn_change_file.clicked.connect(self._change_excel_file) 
        control_row.addWidget(self.btn_change_file)

        # THAY ĐỔI: THÊM NÚT HỦY THAY ĐỔI
        self.btn_undo = QPushButton("Hủy Thay Đổi")
        self.btn_undo.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogResetButton)) 
        self.btn_undo.setFixedWidth(120)
        self.btn_undo.clicked.connect(self._confirm_and_undo) 
        control_row.addWidget(self.btn_undo) # Đặt bên trái nút Lưu Excel
        
        self.btn_save_excel = QPushButton("Lưu Excel")
        self.btn_save_excel.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogSaveButton)) 
        self.btn_save_excel.setFixedWidth(120)
        self.btn_save_excel.clicked.connect(self._confirm_and_save) 
        control_row.addWidget(self.btn_save_excel)
        layout.addLayout(control_row)

        # Sử dụng QTabWidget thay vì QVBoxLayout
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        # THÊM: Giảm font size của Tabs
        self.tab_widget.setStyleSheet("QTabBar::tab { font-size: 9pt; }")

        self.load_excel()

    # [NEW METHOD] Thay đổi file excel
    def _change_excel_file(self):
        new_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Chọn File Excel Mới", 
            os.path.dirname(self.excel_path) if self.excel_path else QDir.homePath(),
            "Excel Files (*.xlsx *.xls)"
        )

        if new_path:
            # Kiểm tra tránh trường hợp chọn lại chính file đó
            if os.path.normcase(new_path) == os.path.normcase(self.excel_path):
                self.msg_label.setText("Đã chọn lại file hiện tại.")
                return

            self.excel_path = new_path
            self.load_excel() # Tải lại dữ liệu từ file mới
            self.msg_label.setText(f"Đã tải file Excel mới: {os.path.basename(new_path)}")

    def _confirm_and_save(self):
        reply = QMessageBox.question(self, "Xác nhận", "Bạn muốn lưu thay đổi và ghi đè lên file Excel gốc?", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            if self.save_overwrite():
                # Nếu lưu thành công, cập nhật original_sheet_data cho các sheet
                self.load_excel() 

    # THAY ĐỔI: Thêm hàm xác nhận và hủy thay đổi
    def _confirm_and_undo(self):
        current_table_widget = self.tab_widget.currentWidget()
        if not isinstance(current_table_widget, QTableWidget):
            self.msg_label.setText("Không có bảng tính để hủy thay đổi.")
            return

        # Yêu cầu xác nhận
        reply = QMessageBox.question(self, "Xác nhận", 
                                     "Bạn có chắc chắn muốn hủy các thay đổi?", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self._undo_changes(current_table_widget)

    def _undo_changes(self, table: QTableWidget):
        """Khôi phục dữ liệu bảng về trạng thái khi tải."""
        sheet_name = self.tab_widget.tabText(self.tab_widget.currentIndex())
        if not hasattr(table, 'original_sheet_data'):
            self.msg_label.setText("Không có dữ liệu gốc để khôi phục.")
            return

        original_data = table.original_sheet_data
        data_rows = original_data['data_rows']
        
        # Thiết lập lại số lượng hàng và cột
        table.setRowCount(len(data_rows))
        table.setColumnCount(len(original_data['header_labels']))
        
        # Điền lại dữ liệu gốc
        for ri, row in enumerate(data_rows):
            for ci, val in enumerate(row):
                table.setItem(ri, ci, QTableWidgetItem(val))
        
        self.msg_label.setText(f"Đã hủy thay đổi trong sheet '{sheet_name}'.")
        # Không cần emit cellChanged(0,0) vì load_excel() sẽ gọi _on_cell_changed cho tất cả các dòng
        # Nên ta chỉ cần kích hoạt lại tính toán công nợ cho tất cả các dòng
        current_sheet_name = self.tab_widget.tabText(self.tab_widget.currentIndex())
        if current_sheet_name in self.col_map:
            col_indices = self.col_map[current_sheet_name]
            g_idx = col_indices['G']
            if g_idx != -1:
                for r in range(table.rowCount()):
                    # Kích hoạt lại tính toán cho từng dòng
                    self._on_cell_changed(r, g_idx) 

    # THAY ĐỔI: Logic tính toán công nợ và debounce
    def _perform_calculation(self):
        if self.pending_calc_row != -1:
            row = self.pending_calc_row
            table = self.tab_widget.currentWidget()
            sheet_name = self.tab_widget.tabText(self.tab_widget.currentIndex())
            
            # Chỉ áp dụng tính toán nếu là sheet công nợ
            if sheet_name in self.col_map:
                col_indices = self.col_map[sheet_name]
                g_idx, h_idx, j_idx = col_indices['G'], col_indices['H'], col_indices['J']
                
                # Chỉ tính khi các cột G, H, J tồn tại
                if g_idx != -1 and h_idx != -1 and j_idx != -1:
                    try:
                        # Lấy giá trị từ cột G và H
                        item_g = table.item(row, g_idx)
                        item_h = table.item(row, h_idx)
                        
                        # Chuyển đổi giá trị sang float, thay thế dấu phẩy nếu có
                        def parse_float(item):
                            if item and item.text().strip():
                                # Loại bỏ định dạng số ngàn và khoảng trắng
                                return float(item.text().replace(',', '').strip()) 
                            return 0.0
                            
                        val_g = parse_float(item_g)
                        val_h = parse_float(item_h)
                        
                        # Tính toán: J = G - H
                        con_phai_thu = val_g - val_h
                        
                        # Cập nhật cột J
                        item_j = QTableWidgetItem(f"{con_phai_thu:,.2f}") # Định dạng lại
                        item_j.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled) # Không cho chỉnh sửa
                        table.setItem(row, j_idx, item_j)

                        # THAY ĐỔI: Thông báo đã tính toán
                        # Đã sửa lỗi: Cột J -> CỘT CÒN PHẢI THU
                        self.msg_label.setText(f"Đã tính toán: CỘT CÒN PHẢI THU={con_phai_thu:,.2f} tại dòng {row+1}")
                        
                    except Exception as e:
                        self.msg_label.setText(f"Lỗi tính toán tại dòng {row+1}: {e}")

            self.pending_calc_row = -1 # Đặt lại

    def _on_cell_changed(self, row, column):
        current_table = self.tab_widget.currentWidget()
        sheet_name = self.tab_widget.tabText(self.tab_widget.currentIndex())
        
        # Kiểm tra xem thay đổi có nằm trong cột G hoặc H của sheet công nợ không
        if sheet_name in self.col_map:
            col_indices = self.col_map[sheet_name]
            
            if col_indices['G'] != -1 and col_indices['H'] != -1:
                if column == col_indices['G'] or column == col_indices['H']:
                    self.pending_calc_row = row
                    # Bắt đầu/Reset timer 0.5s
                    self.debounce_timer.start(DEBOUNCE_DELAY_MS) 

    # THAY ĐỔI: Cập nhật load_excel để lưu dữ liệu gốc và thiết lập tính toán công nợ
    def load_excel(self):
        # Dọn dẹp tab hiện tại
        self.sheets_tables.clear()
        self.sheets_names.clear()
        self.col_map.clear() # Xóa map cột cũ
        while self.tab_widget.count() > 0:
            self.tab_widget.removeTab(0)
            
        self.msg_label.setText("") 

        if not os.path.exists(self.excel_path):
            self.msg_label.setText(f"Không tìm thấy: {self.excel_path}")
            return
            
        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            for name in wb.sheetnames:
                ws = wb[name]
                
                all_rows_raw = list(ws.iter_rows(values_only=True))
                
                # THAY ĐỔI: Lưu trữ dữ liệu thô gốc (dùng cho Undo)
                original_data = {
                    'header_labels': [],
                    'data_rows': [],
                    'header_index': -1 # THÊM: Lưu index hàng header gốc trong excel (tính từ 0)
                }
                
                if not all_rows_raw: continue
                    
                data_rows = []
                header_labels = []
                max_col = 0
                header_row_index = -1
                
                # --- Bước 1: Tìm hàng Header và xác định max_col ---
                for r_idx, row in enumerate(all_rows_raw):
                    cleaned_row = ["" if v is None else str(v).strip() for v in row]
                    
                    if any(c for c in cleaned_row if c):
                        
                        if header_row_index == -1:
                            header_labels = cleaned_row
                            header_row_index = r_idx
                            max_col = len(header_labels)
                            if max_col == 0:
                                max_col = len(row)
                                header_labels = cleaned_row

                        else:
                            data_rows.append(cleaned_row[:max_col])
                            
                
                if not header_labels:
                    continue
                
                # THAY ĐỔI: Lưu dữ liệu đã cắt vào original_data
                original_data['header_labels'] = header_labels
                original_data['data_rows'] = data_rows
                original_data['header_index'] = header_row_index # Lưu index hàng header

                # --- Bước 2: Tạo Table ---
                table = QTableWidget()
                
                # THAY ĐỔI: Gắn dữ liệu gốc vào table object
                table.original_sheet_data = original_data
                
                # Bỏ cột Index (Header dọc)
                table.verticalHeader().setVisible(False)
                table.setStyleSheet("font-size: 9pt;")
                
                table.setRowCount(len(data_rows))
                table.setColumnCount(max_col)
                
                table.setHorizontalHeaderLabels(header_labels)
                
                for col in range(max_col):
                    header_text = header_labels[col] if col < len(header_labels) else f"Column {col+1}"
                    
                    # SỬA LỖI: Sử dụng setHeaderData với Qt.ToolTipRole để đặt Tooltip cho header
                    table.horizontalHeader().model().setHeaderData(
                        col, 
                        Qt.Horizontal, 
                        header_text, 
                        Qt.ToolTipRole
                    )
                
                header = table.horizontalHeader()
                header.setStyleSheet("""
                    QHeaderView::section { 
                        background-color: #f0f0f0; 
                        color: #333; 
                        padding: 6px; 
                        border: 1px solid #ddd;
                        font-weight: bold;
                        font-size: 9pt;
                    }
                """)
                
                # --- Bước 3: Ánh xạ cột tính toán (Công nợ) ---
                g_idx, h_idx, j_idx = -1, -1, -1
                for col_idx, label in enumerate(header_labels):
                    # Sử dụng phương pháp tìm kiếm gần đúng hơn (chứa keyword)
                    label_upper = label.strip().upper()
                    if COL_G_HEADER in label_upper:
                        g_idx = col_idx
                    elif COL_H_HEADER in label_upper:
                        h_idx = col_idx
                    elif COL_J_HEADER in label_upper:
                        j_idx = col_idx

                self.col_map[name] = {'G': g_idx, 'H': h_idx, 'J': j_idx}
                
                # Điền dữ liệu vào table (bắt đầu từ hàng 0)
                for ri, row in enumerate(data_rows):
                    for ci, val in enumerate(row):
                        if ci < max_col:
                            # Thiết lập giá trị mặc định cho cột J
                            if ci == j_idx:
                                item = QTableWidgetItem(val)
                                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled) # Không cho chỉnh sửa
                                table.setItem(ri, ci, item)
                            else:
                                table.setItem(ri, ci, QTableWidgetItem(val))
                        
                # --- Bước 4: Kết nối Signal nếu là sheet công nợ ---
                if g_idx != -1 and h_idx != -1 and j_idx != -1:
                    # Kích hoạt tính toán ban đầu cho tất cả các dòng
                    for r in range(table.rowCount()):
                        self._on_cell_changed(r, g_idx) 
                    
                    # Kết nối sự kiện thay đổi cell để kích hoạt tính toán debounce
                    table.cellChanged.connect(self._on_cell_changed)

                table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                table.setSortingEnabled(True)

                # Thêm table vào tab widget
                self.tab_widget.addTab(table, name)
                self.sheets_tables[name] = table
                self.sheets_names.append(name)
                
        except Exception as e:
            self.msg_label.setText("Lỗi đọc Excel: " + str(e))

    def save_overwrite(self):
        if not self.sheets_tables:
            self.msg_label.setText("Không có sheet để lưu.")
            return False
        try:
            # TẢI LẠI WORKBOOK GỐC để duy trì các công thức và định dạng khác
            wb = openpyxl.load_workbook(self.excel_path, data_only=False) # data_only=False để ghi đè
            
            for name, table in self.sheets_tables.items():
                
                if name not in wb.sheetnames:
                    # Tạo sheet mới nếu không tồn tại
                    ws = wb.create_sheet(title=name)
                    # Giả sử header row là 1
                    header_row_in_excel = 1 
                else:
                    ws = wb[name]
                    # THAY ĐỔI: Tìm lại hàng header đã lưu trong original_sheet_data để ghi đè đúng vị trí
                    # Lưu ý: header_index lưu từ 0, Excel index từ 1.
                    header_row_in_excel = table.original_sheet_data.get('header_index', 0) + 1 

                # Lấy header labels từ table
                header = table.horizontalHeader()
                header_labels = [header.model().headerData(c, Qt.Horizontal) for c in range(table.columnCount())]
                
                # Ghi hàng Header
                for c, label in enumerate(header_labels):
                    ws.cell(row=header_row_in_excel, column=c+1, value=label)
                    
                rows = table.rowCount()
                cols = table.columnCount()
                
                # Ghi dữ liệu (bắt đầu từ hàng sau header)
                for r in range(rows):
                    for c in range(cols):
                        item = table.item(r, c)
                        # Ghi vào vị trí r+1 sau hàng header (header_row_in_excel)
                        excel_row = header_row_in_excel + 1 + r 
                        
                        # Ghi giá trị (không phải công thức). THAY DẤU PHẨY NẾU CÓ TRONG SỐ
                        value_to_write = item.text().replace(',', '') if item else ""
                        
                        # Cần kiểm tra xem có phải là số không để ghi đúng kiểu
                        if isinstance(value_to_write, str) and value_to_write.replace('.', '', 1).isdigit():
                            try:
                                value_to_write = float(value_to_write)
                            except ValueError:
                                pass # Giữ nguyên là chuỗi nếu không phải số hợp lệ

                        ws.cell(row=excel_row, column=c+1, value=value_to_write)
                        
                # Xóa các dòng thừa dưới dữ liệu mới
                # Chỉ xóa nếu số hàng mới ít hơn số hàng cũ
                if header_row_in_excel + 1 + rows < ws.max_row:
                    rows_to_delete = ws.max_row - (header_row_in_excel + 1 + rows) + 1
                    ws.delete_rows(header_row_in_excel + 1 + rows, rows_to_delete)
                        
            os.makedirs(os.path.dirname(self.excel_path), exist_ok=True)
            wb.save(self.excel_path)
            self.msg_label.setText("Đã lưu thành công.")
            return True
        except Exception as e:
            self.msg_label.setText(f"Lỗi lưu: {str(e)}")
            return False

# --- Contract Template Widget ---
class ContractTemplateWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        # THAY ĐỔI: Giữ nguyên màu nền này vì nó đã hợp lý
        self.setStyleSheet("background-color: #f7f9fc;") 
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(6, 6, 6, 6)

        # [NEW FEATURE] Control Row for Reload and Add Template
        control_row = QHBoxLayout()
        control_row.addStretch() # Push buttons to the right
        
        # 1. Reload Button
        self.btn_reload = QPushButton("Reload")
        self.btn_reload.setIcon(QApplication.style().standardIcon(QStyle.SP_BrowserReload)) 
        self.btn_reload.clicked.connect(self.load_templates)
        control_row.addWidget(self.btn_reload)
        
        # 2. Add Template Button
        self.btn_add_template = QPushButton("Thêm Mẫu")
        self.btn_add_template.setIcon(QApplication.style().standardIcon(QStyle.SP_FileIcon)) 
        self.btn_add_template.clicked.connect(self.add_new_template)
        control_row.addWidget(self.btn_add_template)

        main_layout.addLayout(control_row) # Add control row first

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")
        h_container = QWidget()
        h_container.setStyleSheet("QWidget { background-color: transparent; }") 
        self.layout = QHBoxLayout(h_container) 
        self.layout.setAlignment(Qt.AlignLeft | Qt.AlignTop) 
        scroll_area.setWidget(h_container)
        
        main_layout.addWidget(scroll_area) # Then add scroll area
        
        self.load_templates()

    def load_templates(self):
        while self.layout.count():
            item = self.layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        if not os.path.isdir(CONTRACT_TEMPLATE_PATH):
            self.layout.addWidget(QLabel(f"Không tìm thấy thư mục mẫu hợp đồng: {CONTRACT_TEMPLATE_PATH}"))
            return

        all_files = os.listdir(CONTRACT_TEMPLATE_PATH) 
        template_files = [
            f for f in all_files 
            if f.lower().endswith(('.doc', '.docx')) 
            and os.path.isfile(os.path.join(CONTRACT_TEMPLATE_PATH, f))
            and not f.startswith("~$")
        ]

        if not template_files:
            self.layout.addWidget(QLabel("Không có file mẫu (.doc, .docx) nào trong thư mục."))
            return

        for filename in template_files:
            file_path = os.path.join(CONTRACT_TEMPLATE_PATH, filename)
            self.layout.addWidget(self._create_template_item(filename, file_path))
            
    # [NEW METHOD] Thêm mẫu hợp đồng mới
    def add_new_template(self):
        target_path = CONTRACT_TEMPLATE_PATH

        if not os.path.isdir(target_path):
            QMessageBox.critical(self, "Lỗi Thư mục", f"Thư mục mẫu không tồn tại:\n{target_path}")
            return

        file_path, _ = QFileDialog.getOpenFileName(self, 
                                                   "Chọn File Mẫu để Thêm vào", 
                                                   QDir.homePath(), 
                                                   "Tài liệu (*.doc *.docx);;Tất cả File (*.*)")

        if file_path:
            file_name = os.path.basename(file_path)
            destination = os.path.join(target_path, file_name)
            
            if os.path.exists(destination):
                reply = QMessageBox.question(self, "Xác nhận Ghi đè", 
                                             f"File '{file_name}' đã tồn tại trong thư mục mẫu.\nBạn có muốn ghi đè?",
                                             QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    return

            try:
                # Sử dụng shutil.copy2 để sao chép cả metadata (ngày tháng)
                shutil.copy2(file_path, destination) 
                self.load_templates() # Tải lại danh sách mẫu
                QMessageBox.information(self, "Thành công", f"Đã sao chép '{file_name}' thành công vào thư mục mẫu.")
                
            except Exception as e:
                QMessageBox.critical(self, "Lỗi Sao chép", f"Không thể sao chép file:\n{e}")

    def _create_template_item(self, filename, file_path):
        try:
            file_stats = os.stat(file_path)
            size_bytes = file_stats.st_size
            size_kb = size_bytes / 1024
            m_time = datetime.fromtimestamp(file_stats.st_mtime).strftime('%d/%m/%Y %H:%M')
            size_str = f"{size_kb:.2f} KB" if size_kb < 1024 else f"{size_kb / 1024:.2f} MB"
        except Exception:
            m_time = "N/A"
            size_str = "N/A"

        class ClickableCard(QWidget):
            clicked = Signal(str)
            def __init__(self, path, parent=None):
                super().__init__(parent)
                self.file_path = path
                self.setCursor(Qt.PointingHandCursor)
                self.setFixedSize(260, 180) 
                self.default_style = self._get_default_style()
                self.hover_style = self._get_hover_style()
                self.setStyleSheet(self.default_style)
            
            def _get_default_style(self):
                return """
                    ClickableCard { 
                        border: 1px solid #ddd; 
                        border-radius: 12px; 
                        background-color: white; /* THAY ĐỔI: Màu nền thẻ là trắng */
                        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); 
                        margin-top: 5px; 
                        margin-bottom: 5px;
                    }
                """
            
            def _get_hover_style(self):
                return """
                    ClickableCard {
                        border: 1px solid #ddd; 
                        border-radius: 12px; 
                        background-color: #e6f0ff; /* THAY ĐỔI: Màu hover */
                        box-shadow: 0 10px 20px rgba(0, 0, 0, 0.2);
                        margin-top: 1px;
                        margin-bottom: 9px; 
                    }
                """

            def enterEvent(self, event):
                self.setStyleSheet(self.hover_style)
                super().enterEvent(event)

            def leaveEvent(self, event):
                self.setStyleSheet(self.default_style)
                super().leaveEvent(event)

            def mousePressEvent(self, event):
                if event.button() == Qt.LeftButton:
                    self.clicked.emit(self.file_path)
                    
        item_widget = ClickableCard(file_path)
        item_layout = QVBoxLayout(item_widget) 
        item_layout.setContentsMargins(15, 15, 15, 15)
        item_layout.setSpacing(8) 
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        icon_label = QLabel()
        style = QApplication.style()
        icon_label.setPixmap(style.standardIcon(QStyle.SP_FileIcon).pixmap(QSize(40, 40))) 
        
        file_label = QLabel(f"<b>{filename}</b>")
        file_label.setStyleSheet("font-size: 10pt; color: #1f2937;")
        file_label.setWordWrap(True)
        file_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        header_layout.addWidget(icon_label)
        header_layout.addWidget(file_label)
        header_layout.addStretch()

        item_layout.addLayout(header_layout)
        date_label = QLabel(f"<span style='font-size: 9pt; color:#6b7280;'>Date modified:</span> {m_time}")
        size_label = QLabel(f"<span style='font-size: 9pt; color:#6b7280;'>File size:</span> {size_str}")

        item_layout.addWidget(date_label)
        item_layout.addWidget(size_label)
        item_layout.addStretch() 
        item_widget.clicked.connect(self.open_file_with_default_app)
        
        return item_widget
    
    def open_file_with_default_app(self, file_path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(file_path)
            elif sys.platform == "darwin":
                subprocess.call(["open", file_path])
            else:
                subprocess.call(["xdg-open", file_path])
        except Exception as e:
            QMessageBox.critical(self, "Lỗi Mở File", f"Không thể mở file '{os.path.basename(file_path)}':\n{e}")

# ---------------- Main Window ----------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Manage Folder")
        self.resize(1200, 800)

        style = self.style()
        
        # THAY ĐỔI: Đặt Icon ứng dụng từ file đã upload (image_6af425.png)
        icon_path = "image_6af425.png"
        if os.path.exists(icon_path):
            custom_folder_icon = QIcon(icon_path)
            self.setWindowIcon(custom_folder_icon)
            QApplication.setWindowIcon(custom_folder_icon)
        else:
            print(f"CẢNH BÁO: Không tìm thấy icon tùy chỉnh tại '{icon_path}'. Sử dụng icon mặc định.")
            default_folder_icon = style.standardIcon(QStyle.SP_DirIcon)
            self.setWindowIcon(default_folder_icon)
            QApplication.setWindowIcon(default_folder_icon)

        # THAY ĐỔI: CSS để đồng nhất màu nền
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f7f9fc; /* Màu nền chính của cửa sổ */
            }
            QWidget, QLabel, QPushButton, QLineEdit, QTableWidget, QTreeView {
                font-size: 10pt;
            }
            QTabWidget::pane {
                border-top: 1px solid #ccc;
                background-color: #f7f9fc; /* Màu nền của nội dung tab */
            }
            QTabBar::tab { 
                font-size: 11pt;
                padding: 6px 10px;
                min-height: 30px;
                background-color: #f0f0f0; /* Màu nền tab không chọn */
            }
            QTabBar::tab:selected {
                background-color: #f7f9fc; /* Màu nền tab chọn */
                color: #004d99;
                border-bottom: 2px solid #004d99;
            }
            QPushButton {
                padding: 6px; 
            }
            #btn_refresh { 
                padding: 6px;
            }
            #lbl_filename {
                font-size: 17pt; 
                font-weight: bold;
            }
            /* THAY ĐỔI: Để list search không bị tách biệt quá nhiều */
            QListWidget { 
                border-top: none; 
            }
        """)

        tabs = QTabWidget()
        self.setCentralWidget(tabs)
        
        
        # --- Tab 1: Quản lý file ---
        t_file = QWidget()
        # THÊM: Đồng nhất màu nền cho widget chứa nội dung tab
        t_file.setStyleSheet("QWidget { background-color: #f7f9fc; }") 
        tfile_layout = QVBoxLayout(t_file)
        tfile_layout.setContentsMargins(6,6,6,6)
        
        tree_container = QWidget()
        tree_container.setStyleSheet("QWidget { background-color: #f7f9fc; }")
        self.tree_vlayout = QVBoxLayout(tree_container) 
        self.tree_vlayout.setContentsMargins(0, 0, 0, 0)
        
        # Search bar
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Tìm kiếm gần đúng tên file...") 
        self.search_input.setFixedHeight(30)
        
        search_icon = style.standardIcon(QStyle.SP_FileDialogContentsView)
        self.search_input.addAction(search_icon, QLineEdit.LeadingPosition)
        # THAY ĐỔI: Sử dụng textChanged thay vì returnPressed để tìm kiếm hiển thị modal
        self.search_input.textChanged.connect(self.on_search) 
        
        search_layout.addWidget(self.search_input)
        self.tree_vlayout.addLayout(search_layout) 
        
        # THAY ĐỔI: List Widget hiển thị kết quả Search (Tích hợp, không dùng Popup)
        self.search_results_list = QListWidget()
        self.search_results_list.setFocusPolicy(Qt.NoFocus) 
        self.search_results_list.setStyleSheet("QListWidget { border: 1px solid #ccc; background-color: white; }")
        # THAY ĐỔI (QUAN TRỌNG): Cố định chiều cao 200px
        self.search_results_list.setFixedHeight(200) 
        self.search_results_list.hide() # Ẩn mặc định
        
        self.search_results_list.itemClicked.connect(self._select_file_from_search)
        
        # Thêm thẳng vào layout (ngay dưới search_input)
        self.tree_vlayout.addWidget(self.search_results_list) 
        
        # NÚT THÊM FILE & NÚT XÓA FILE & NÚT THAY ĐỔI ROOT & NÚT REFRESH
        file_actions_layout = QHBoxLayout()
        
        self.btn_add_file = QPushButton("Thêm File")
        self.btn_add_file.setIcon(style.standardIcon(QStyle.SP_FileIcon)) 
        self.btn_add_file.setToolTip("Sao chép file từ máy tính vào thư mục đang chọn.")
        self.btn_add_file.clicked.connect(self.add_file_to_current_folder)
        file_actions_layout.addWidget(self.btn_add_file)
        
        self.btn_delete_file = QPushButton("Xóa File/Folder")
        self.btn_delete_file.setIcon(style.standardIcon(QStyle.SP_TrashIcon))
        self.btn_delete_file.setToolTip("Xóa mục đang chọn.")
        self.btn_delete_file.clicked.connect(self.delete_selected_file)
        file_actions_layout.addWidget(self.btn_delete_file)
        
        # THÊM: NÚT THAY ĐỔI ROOT FILE
        self.btn_change_root = QPushButton("Thay đổi Folder gốc")
        self.btn_change_root.setIcon(style.standardIcon(QStyle.SP_DirHomeIcon))
        self.btn_change_root.setToolTip("Thay đổi thư mục gốc của cây thư mục.")
        self.btn_change_root.clicked.connect(self.change_root_folder)
        file_actions_layout.addWidget(self.btn_change_root)
        
        file_actions_layout.addStretch() 
        
        self.btn_refresh = QPushButton()
        self.btn_refresh.setObjectName("btn_refresh") 
        self.btn_refresh.setIcon(style.standardIcon(QStyle.SP_BrowserReload)) 
        self.btn_refresh.setToolTip("Làm mới cây thư mục.")
        self.btn_refresh.setFixedWidth(35) 
        self.btn_refresh.clicked.connect(self.refresh_tree)
        file_actions_layout.addWidget(self.btn_refresh)
        
        self.tree_vlayout.addLayout(file_actions_layout) 
        
        # SPLITTER VÀ TREE
        splitter = QSplitter(Qt.Horizontal)
        tfile_layout.addWidget(splitter)
        
        self.model = CustomFileSystemModel()
        self.model.setRootPath(ROOT_FOLDER)
        self.model.setFilter(QDir.NoDotAndDotDot | QDir.AllEntries)
        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(ROOT_FOLDER))
        
        self.tree.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tree.setHeaderHidden(False) 
        # THÊM: Đồng nhất màu nền cho Tree View
        self.tree.setStyleSheet("QTreeView { background-color: white; border: 1px solid #ccc; }")

        self.tree.hideColumn(1) 
        self.tree.hideColumn(2) 
        
        header = self.tree.header()
        header.setSectionResizeMode(0, QHeaderView.Stretch) 
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents) 
        header.setStretchLastSection(False) 

        self.tree.clicked.connect(self.on_tree_clicked)
        self.tree_vlayout.addWidget(self.tree)
        splitter.addWidget(tree_container)

        # Preview area (right)
        right_widget = QWidget()
        right_widget.setStyleSheet("QWidget { background-color: #f7f9fc; }")
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0,0,0,0)
        
        header_row = QHBoxLayout()
        
        self.lbl_filename = QLabel("Chọn một tệp để xem trước")
        self.lbl_filename.setObjectName("lbl_filename")
        header_row.addWidget(self.lbl_filename)
        header_row.addStretch()
        
        self.btn_open_default = QPushButton("Mở File") 
        self.btn_open_default.setIcon(style.standardIcon(QStyle.SP_DialogOpenButton)) 
        self.btn_open_default.setEnabled(False)
        self.btn_open_default.clicked.connect(self.open_with_default_app)
        header_row.addWidget(self.btn_open_default)
        right_layout.addLayout(header_row)

        self.preview = PreviewWidget()
        right_layout.addWidget(self.preview)
        
        self.lbl_file_info = QLabel("Ngày: N/A")
        self.lbl_file_info.setStyleSheet("color: gray; font-size: 9pt; padding: 4px; border-top: 1px solid #ccc;")
        self.lbl_file_info.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.lbl_file_info.hide()
        right_layout.addWidget(self.lbl_file_info)

        splitter.addWidget(right_widget)
        splitter.setSizes([400, 800]) 

        tabs.addTab(t_file, style.standardIcon(QStyle.SP_DirIcon), "Quản lý file")

        # --- Tab 2: Theo dõi tiến độ (Excel editor) ---
        t_progress = QWidget()
        t_progress.setStyleSheet("QWidget { background-color: #f7f9fc; }") # Đồng nhất màu nền
        progress_layout = QVBoxLayout(t_progress)
        progress_layout.setContentsMargins(6,6,6,6)

        self.excel_editor = ExcelEditor(DEFAULT_EXCEL_PATH)
        progress_layout.addWidget(self.excel_editor)
        
        tabs.addTab(t_progress, style.standardIcon(QStyle.SP_BrowserReload), "Theo dõi tiến độ")
        
        # --- Tab 3: Mẫu hợp đồng ---
        t_contract = QWidget()
        t_contract.setStyleSheet("QWidget { background-color: #f7f9fc; }") # Đồng nhất màu nền
        contract_layout = QVBoxLayout(t_contract)
        contract_layout.setContentsMargins(0, 0, 0, 0)
        self.contract_widget = ContractTemplateWidget()
        contract_layout.addWidget(self.contract_widget)
        
        tabs.addTab(t_contract, style.standardIcon(QStyle.SP_FileLinkIcon), "Mẫu hợp đồng")

        self.current_file = None

    # ----------------------------------------------
    # PHƯƠNG THỨC: THAY ĐỔI ROOT FOLDER (MỚI)
    # ----------------------------------------------
    def change_root_folder(self):
        global ROOT_FOLDER
        
        # Mở hộp thoại chọn thư mục
        new_root = QFileDialog.getExistingDirectory(
            self, 
            "Chọn Thư mục Gốc mới", 
            ROOT_FOLDER,
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )

        if new_root and os.path.isdir(new_root):
            if new_root == ROOT_FOLDER:
                self.statusBar().showMessage("Thư mục gốc không thay đổi.", 3000)
                return
            
            # Cập nhật biến global (chỉ dùng để tham chiếu mặc định lần sau)
            ROOT_FOLDER = new_root
            
            # Cập nhật Model và Tree View
            self.model.setRootPath(new_root)
            self.tree.setRootIndex(self.model.index(new_root))
            self.tree.collapseAll() # Thu gọn tất cả cho thư mục mới

            self.preview.show_text(f"Thư mục gốc đã được thay đổi thành:\n{new_root}")
            self.lbl_filename.setText(os.path.basename(new_root) or new_root)
            self.statusBar().showMessage(f"Thư mục gốc đã thay đổi: {new_root}", 5000)
        else:
            self.statusBar().showMessage("Không thay đổi thư mục gốc.", 3000)


    # ----------------------------------------------
    # PHƯƠNG THỨC: REFRESH TREEVIEW
    # ----------------------------------------------
    def refresh_tree(self):
        current_index = self.tree.currentIndex()
        old_root_path = self.model.rootPath()
        
        self.model.setRootPath(old_root_path) 
        self.tree.setRootIndex(self.model.index(old_root_path))
        
        if current_index.isValid():
            current_path = self.model.filePath(current_index)
            new_index = self.model.index(current_path)
            if new_index.isValid():
                self.tree.setCurrentIndex(new_index)
                self.tree.scrollTo(new_index)
                
        self.statusBar().showMessage("Đã làm mới cây thư mục", 3000)

    # ----------------------------------------------
    # PHƯƠNG THỨC: TÌM KIẾM FILE VÀ HIGHLIGHT
    # ----------------------------------------------
    def search_files_and_prepare_highlight(self, search_term: str) -> List[Dict[str, str]]:
        """Tìm kiếm file gần đúng, trả về tên file gốc và tên đã được highlight (chứa tag span)."""
        search_results = []
        search_term = search_term.strip()
        
        if not search_term:
            return []

        # Pattern: '(?i)' để tìm kiếm không phân biệt chữ hoa/thường (case-insensitive)
        # Sử dụng ROOT_FOLDER hiện tại
        current_root = self.model.rootPath()
        regex_pattern = f"(?i)({re.escape(search_term)})"
        
        for root, _, files in os.walk(current_root):
            for file_name in files:
                
                if search_term.lower() in file_name.lower():
                    
                    # Tạo chuỗi HTML để highlight.
                    highlighted_name = re.sub(
                        regex_pattern, 
                        r'<span style="background-color: yellow;">\1</span>',
                        file_name
                    )
                    
                    search_results.append({
                        "original_name": file_name, 
                        "highlighted_name": highlighted_name, 
                        "full_path": os.path.join(root, file_name)
                    })
                    
        return search_results

    def on_search(self):
        """THAY ĐỔI: Xử lý tìm kiếm và hiển thị trong QListWidget tích hợp.
           - Không sử dụng Qt.Popup để tránh mất focus.
        """
        q = self.search_input.text().strip()
        
        # Nếu ô tìm kiếm trống, ẩn danh sách
        if not q:
            self.search_results_list.hide() 
            self.statusBar().clearMessage()
            return

        results = self.search_files_and_prepare_highlight(q)
        
        # Tắt tín hiệu tạm thời để tránh kích hoạt các sự kiện không cần thiết khi dọn dẹp list
        self.search_results_list.blockSignals(True) 
        self.search_results_list.clear()
        self.search_results_list.blockSignals(False)
        
        if results:
            for item_data in results:
                item = QListWidgetItem()
                
                # Hiển thị tên file gốc (original_name)
                item.setText(item_data["original_name"]) 
                item.setData(Qt.UserRole, item_data["full_path"])
                item.setData(Qt.UserRole + 1, item_data["original_name"]) 
                self.search_results_list.addItem(item)
                
            # THAY ĐỔI: Chỉ cần hiển thị QListWidget
            self.search_results_list.show()
            
            # Giữ focus trên QLineEdit (nên giữ lại dù không còn dùng Popup)
            self.search_input.setFocus() 
            
            self.statusBar().showMessage(f"Tìm thấy {len(results)} kết quả cho '{q}'", 5000)
        else:
            # THAY ĐỔI: Ẩn QListWidget trực tiếp
            self.search_results_list.hide()
            self.search_input.setFocus() 
            self.statusBar().showMessage(f"Không tìm thấy kết quả nào cho '{q}'", 5000)

    def _select_file_from_search(self, item: QListWidgetItem):
        """Xử lý khi người dùng click vào một file trong danh sách tìm kiếm."""
        file_path = item.data(Qt.UserRole)
        if file_path and os.path.exists(file_path):
            idx = self.model.index(file_path)
            if idx.isValid():
                self.tree.setCurrentIndex(idx)
                self.tree.scrollTo(idx)
                self.on_tree_clicked(idx) # Kích hoạt hàm preview
            
            # THAY ĐỔI: Ẩn QListWidget trực tiếp
            self.search_results_list.hide()
            self.search_input.setText("") # Xóa nội dung tìm kiếm sau khi chọn


    # ----------------------------------------------
    # PHƯƠNG THỨC: XÓA FILE (Đã chỉnh sửa nút xác nhận)
    # ----------------------------------------------
    def delete_selected_file(self):
        current_index = self.tree.currentIndex()
        if not current_index.isValid():
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn một file hoặc thư mục để xóa.")
            return

        file_path = self.model.filePath(current_index)
        file_name = os.path.basename(file_path)
        is_dir = os.path.isdir(file_path)

        if not os.path.exists(file_path):
             QMessageBox.warning(self, "Lỗi", "Đường dẫn không tồn tại.")
             return
        
        parent_index = current_index.parent()
        
        type_str = "thư mục" if is_dir else "file"
        
        # THAY ĐỔI (QUAN TRỌNG): Sử dụng custom QMessageBox với nút "Có" và "Không"
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Xác nhận Xóa")
        msg_box.setText(f"Bạn có chắc muốn xóa vĩnh viễn {type_str}:\n{file_name}?")
        
        # Custom Buttons
        btn_yes = msg_box.addButton("Có", QMessageBox.YesRole)
        btn_no = msg_box.addButton("Không", QMessageBox.NoRole)
        msg_box.setDefaultButton(btn_no) # Mặc định chọn Không
        
        msg_box.exec()
        
        if msg_box.clickedButton() != btn_yes:
            return # Người dùng chọn Không hoặc đóng hộp thoại

        # Logic xóa file
        try:
            if is_dir:
                shutil.rmtree(file_path) 
            else:
                os.remove(file_path) 
            
            self.preview.show_text(f"Đã xóa {type_str}: {file_name}")
            self.lbl_filename.setText("Chọn một tệp để xem trước")
            
            self.refresh_tree()
            self.tree.setCurrentIndex(parent_index)

            QMessageBox.information(self, "Thành công", f"Đã xóa {type_str} '{file_name}' thành công.")
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi Xóa", f"Không thể xóa {type_str}:\n{e}")
    # ----------------------------------------------
    # CÁC PHƯƠNG THỨC KHÁC
    # ----------------------------------------------
        
    def add_file_to_current_folder(self):
        current_index = self.tree.currentIndex()
        if not current_index.isValid():
            target_path = self.model.rootPath()
        else:
            selected_path = self.model.filePath(current_index)
            if os.path.isdir(selected_path):
                target_path = selected_path
            else:
                target_path = os.path.dirname(selected_path)

        if not os.path.isdir(target_path):
             QMessageBox.warning(self, "Lỗi Thư mục", f"Thư mục đích không hợp lệ: {target_path}")
             return

        file_path, _ = QFileDialog.getOpenFileName(self, 
                                                   "Chọn File để Thêm vào", 
                                                   QDir.homePath(), 
                                                   "Tất cả File (*.*)")

        if file_path:
            file_name = os.path.basename(file_path)
            destination = os.path.join(target_path, file_name)
            
            if os.path.exists(destination):
                reply = QMessageBox.question(self, "Xác nhận", 
                                             f"Bạn có muốn lưu thay đổi?",
                                             QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    return

            try:
                shutil.copy2(file_path, destination)
                self.refresh_tree()
                target_index = self.model.index(target_path)
                self.tree.expand(target_index)
                QMessageBox.information(self, "Thành công", f"Đã sao chép '{file_name}' thành công vào:\n{target_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "Lỗi Sao chép", f"Không thể sao chép file:\n{e}")


    def on_tree_clicked(self, index):
        # Giữ nguyên logic cũ
        file_path = self.model.filePath(index)
        file_name = self.model.fileName(index)

        if file_name.startswith("~$"):
            self.preview.show_text("Không thể xem trước file tạm thời.")
            self.lbl_filename.setText(file_name)
            self.btn_open_default.setEnabled(False)
            self.current_file = None
            return

        date_info = "Ngày: N/A"
        self.current_file = None
        self.btn_open_default.setEnabled(False)
        self.lbl_file_info.hide()

        if os.path.exists(file_path) and not os.path.isdir(file_path):
            try:
                m_time = os.path.getmtime(file_path)
                date_info = f"Ngày lưu cuối: {datetime.fromtimestamp(m_time).strftime('%d/%m/%Y %H:%M:%S')}"
                self.lbl_file_info.show()
            except Exception:
                pass

        self.lbl_file_info.setText(date_info)

        if os.path.isdir(file_path):
            if self.tree.isExpanded(index):
                self.tree.collapse(index)
            else:
                self.tree.expand(index)
            self.lbl_filename.setText(os.path.basename(file_path) or file_path)
            self.preview.show_text("Thư mục được chọn")
            return

        self.current_file = file_path
        self.lbl_filename.setText(os.path.basename(file_path)) 
        self.btn_open_default.setEnabled(True)

        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext in (".txt", ".md", ".py", ".csv", ".log"):
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    txt = f.read()
                self.preview.show_text(txt)
            elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif"):
                self.preview.show_image(file_path)
            elif ext == ".pdf":
                with open(file_path, "rb") as f:
                    b = f.read()
                self.preview.show_pdf_bytes(b)
            elif ext == ".docx":
                if Document is None:
                    self.preview.show_text("python-docx chưa cài. Cài 'python-docx' để xem .docx.")
                else:
                    doc = Document(file_path)
                    text = "\n\n".join(p.text for p in doc.paragraphs)
                    self.preview.show_text(text)
            elif ext in (".xlsx", ".xls"):
                try:
                    wb = openpyxl.load_workbook(file_path, data_only=True)
                    ws = wb.active
                    rows = []
                    for r in ws.iter_rows(values_only=True):
                        rows.append([("" if c is None else str(c)) for c in r])
                    out = "\n".join("\t".join(row) for row in rows[:200])
                    self.preview.show_text(out)
                except Exception as e:
                    self.preview.show_text("Không thể mở Excel: " + str(e))
            else:
                self.preview.show_text("Không hỗ trợ xem trước định dạng này.")
        except Exception as e:
            self.preview.show_text("Lỗi xem trước: " + str(e))

    def open_with_default_app(self):
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
            QMessageBox.warning(self, "Mở file", f"Không thể mở file: {e}")

def main():
    if fitz is None:
        print("CẢNH BÁO: Thư viện 'PyMuPDF' (fitz) chưa được cài đặt. Không thể xem trước file PDF.")
    if Document is None:
        print("CẢNH BẢO: Thư viện 'python-docx' chưa được cài đặt. Không thể xem trước file DOCX.")
    
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
import os
import shutil
import subprocess
import sys
from datetime import datetime

from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QApplication, QScrollArea, QFileDialog, QMessageBox,
    QSizePolicy, QGridLayout, QStyle, QLineEdit
)

# Thư mục mẫu hợp đồng
CONTRACT_TEMPLATE_PATH = r"D:\Template"

# Thử import googletrans và python-docx
try:
    from googletrans import Translator
except Exception:
    Translator = None

try:
    from docx import Document
except Exception:
    Document = None


class ContractTemplateWidget(QWidget):
    translation_start_signal = Signal(str, str)  # file_path, save_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color: #f7f9fc;")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(6, 6, 6, 6)

        self.translator = Translator() if Translator else None

        control_row = QHBoxLayout()

        # --- Ô Tìm kiếm (Search Box) ---
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Tìm kiếm mẫu hợp đồng...")
        # CSS cho QLineEdit để đẹp và cùng chiều cao với QPushButton
        self.search_box.setStyleSheet(
            "QLineEdit { "
            "   padding: 5px; "
            "   border: 1px solid #ccc; "
            "   border-radius: 5px; "
            "   font-size: 14px; "
            "   max-width: 350px; " # Giới hạn chiều rộng tối đa
            "}"
            "QLineEdit:focus { "
            "   border: 1px solid #4a90e2; "
            "   border-width: 1px; "
            "}"
        )
        self.search_box.textChanged.connect(self.load_templates)
        control_row.addWidget(self.search_box)
        
        control_row.addStretch() # Thêm stretch để đẩy các button sang phải

        self.btn_reload = QPushButton("Reload")
        self.btn_reload.setIcon(QApplication.style().standardIcon(QStyle.SP_BrowserReload))
        self.btn_reload.clicked.connect(self.load_templates)
        control_row.addWidget(self.btn_reload)

        self.btn_add_template = QPushButton("Thêm Mẫu")
        self.btn_add_template.setIcon(QApplication.style().standardIcon(QStyle.SP_FileIcon))
        self.btn_add_template.clicked.connect(self.add_new_template)
        control_row.addWidget(self.btn_add_template)

        self.btn_translate_contract = QPushButton("Tạo Hợp đồng Translate")
        self.btn_translate_contract.setIcon(QApplication.style().standardIcon(QStyle.SP_MessageBoxQuestion))
        self.btn_translate_contract.setToolTip("Chọn file Word mẫu để dịch sang tiếng Anh.")
        self.btn_translate_contract.clicked.connect(self.create_translated_contract)
        control_row.addWidget(self.btn_translate_contract)

        main_layout.addLayout(control_row)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        # Giữ nguyên background trong suốt cho QScrollArea
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")

        h_container = QWidget()
        h_container.setStyleSheet("QWidget { background-color: transparent; }")

        self.layout = QGridLayout(h_container)
        self.layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.layout.setSpacing(10) # Thêm khoảng cách giữa các card
        scroll_area.setWidget(h_container)
        main_layout.addWidget(scroll_area)

        # Lần đầu load templates (search box rỗng)
        self.load_templates()

    def load_templates(self):
        # Lấy nội dung tìm kiếm
        filter_text = self.search_box.text().lower()
        
        # Xóa các widget cũ
        for i in reversed(range(self.layout.count())):
            item = self.layout.itemAt(i)
            if item.widget():
                item.widget().deleteLater()

        if not os.path.isdir(CONTRACT_TEMPLATE_PATH):
            self.layout.addWidget(QLabel(f"Không tìm thấy thư mục: {CONTRACT_TEMPLATE_PATH}"))
            return

        # Lọc danh sách file theo nội dung tìm kiếm
        files = [
            f for f in os.listdir(CONTRACT_TEMPLATE_PATH)
            if f.lower().endswith(('.doc', '.docx')) and not f.startswith("~$")
            and filter_text in f.lower() # Thêm điều kiện lọc
        ]
        
        if not files:
            if filter_text:
                self.layout.addWidget(QLabel(f"Không tìm thấy mẫu hợp đồng cho: '{self.search_box.text()}'"))
            else:
                self.layout.addWidget(QLabel("Không có file mẫu hợp đồng."))
            return

        # Đặt số cột tối đa là 5
        col_count = 5 
        for i, name in enumerate(files):
            path = os.path.join(CONTRACT_TEMPLATE_PATH, name)
            row, col = divmod(i, col_count)
            # Thêm widget ClickableCard (chứa tất cả thông tin) vào vị trí (row, col)
            self.layout.addWidget(self._create_card(name, path), row, col)

        # *** Cập nhật quan trọng để Full màn hình với 6 card: Thiết lập Column Stretch ***
        for col in range(col_count):
            self.layout.setColumnStretch(col, 1)


    def _create_card(self, filename, file_path):
        try:
            st = os.stat(file_path)
            size_kb = st.st_size / 1024
            mtime = datetime.fromtimestamp(st.st_mtime).strftime("%d/%m/%Y %H:%M")
            size_str = f"{size_kb:.2f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
        except Exception:
            mtime = "N/A"
            size_str = "N/A"

        class ClickableCard(QWidget):
            clicked = Signal(str)
            def __init__(self, path):
                super().__init__()
                self.path = path
                self.setCursor(Qt.PointingHandCursor)
                # Đã loại bỏ setMinimumHeight(80) để card có thể linh hoạt hơn

                # CSS mới: Thẻ có border và background trắng (tạo sự phân tách rõ ràng)
                self.default = """
                    QWidget { 
                        background-color: white; 
                        border: 1px solid #ddd; 
                        border-radius: 5px; 
                        padding: 8px;
                    }
                """
                # CSS hover: Thay đổi background và border nhưng không thêm đường ngăn cách
                self.hover = """
                    QWidget { 
                        background-color: #e6f0ff; /* Light blue background for hover */
                        border: 1px solid #4a90e2; /* Blue border for focus/hover */
                        border-radius: 5px; 
                        padding: 8px;
                    }
                """
                self.setStyleSheet(self.default)

            def enterEvent(self, e): self.setStyleSheet(self.hover)
            def leaveEvent(self, e): self.setStyleSheet(self.default)
            def mousePressEvent(self, e):
                if e.button() == Qt.LeftButton: self.clicked.emit(self.path)

        card = ClickableCard(file_path)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 10, 10, 10)

        
        lbl_filename = QLabel(f"<b>{filename}</b><br\><p>Date: {mtime}</p><br\><p>Size: {size_str}</p>")
        lbl_filename.setWordWrap(True)
        # Thiết lập QSizePolicy.Expanding cho chiều ngang để card giãn ra tối đa trong ô grid
        lbl_filename.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred) 

        layout.addWidget(lbl_filename)
      
        card.clicked.connect(self.open_file)
        return card

    def add_new_template(self):
        if not os.path.isdir(CONTRACT_TEMPLATE_PATH):
            QMessageBox.critical(self, "Lỗi", f"Không tìm thấy thư mục: {CONTRACT_TEMPLATE_PATH}")
            return
        src, _ = QFileDialog.getOpenFileName(self, "Chọn file mẫu", "", "Word Files (*.doc *.docx)")
        if not src: return
        dst = os.path.join(CONTRACT_TEMPLATE_PATH, os.path.basename(src))
        if os.path.exists(dst):
            r = QMessageBox.question(self, "Ghi đè", f"File {dst} đã tồn tại. Ghi đè?")
            if r != QMessageBox.Yes: return
        shutil.copy2(src, dst)
        self.load_templates()

    def create_translated_contract(self):
        if Document is None or self.translator is None:
            QMessageBox.critical(self, "Thiếu thư viện", "Cần cài python-docx và googletrans.")
            return
        src, _ = QFileDialog.getOpenFileName(self, "Chọn File Mẫu", CONTRACT_TEMPLATE_PATH, "Word Files (*.doc *.docx)")
        if not src: return
        save, _ = QFileDialog.getSaveFileName(self, "Lưu Hợp đồng Dịch",
            os.path.join(CONTRACT_TEMPLATE_PATH, f"Translated_{os.path.basename(src)}"),
            "Word Files (*.docx)")
        if not save: return
        self.translation_start_signal.emit(src, save)

    def open_file(self, path):
        try:
            if sys.platform.startswith("win"): os.startfile(path)
            elif sys.platform == "darwin": subprocess.call(["open", path])
            else: subprocess.call(["xdg-open", path])
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở file: {e}")
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
    QSizePolicy, QGridLayout, QStyle, QLineEdit, QDialog,
    QComboBox, QFormLayout 
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


# --- START: TranslationDialog (Popup trước khi dịch) ---
class TranslationDialog(QDialog):
    """
    Hộp thoại Popup cho phép người dùng chọn file mẫu, 
    chọn công cụ dịch, và nhập tên file lưu mới.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Tạo Hợp đồng Dịch")
        self.setFixedSize(450, 250)
        self.setStyleSheet("background-color: #f7f9fc;")

        self.src_file_path = None
        self.save_file_path = None
        
        main_layout = QVBoxLayout(self)

        # --- Form Layout ---
        form_layout = QFormLayout()
        form_layout.setRowWrapPolicy(QFormLayout.WrapAllRows)
        form_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        
        # 1. Chọn File
        self.file_path_label = QLineEdit("Chưa chọn file")
        self.file_path_label.setReadOnly(True)
        self.file_path_label.setStyleSheet("background-color: #e0e0e0; padding: 5px;")
        
        btn_browse = QPushButton("Chọn File Mẫu")
        btn_browse.clicked.connect(self._select_file)
        
        file_selection_layout = QHBoxLayout()
        file_selection_layout.addWidget(self.file_path_label)
        file_selection_layout.addWidget(btn_browse)
        
        form_layout.addRow("File Mẫu:", file_selection_layout)

        # 2. Chọn Công cụ Dịch
        self.combo_translator = QComboBox()
        self.combo_translator.addItem("Google Translate", "google")
        self.combo_translator.addItem("OpenAI (chưa hỗ trợ)", "openai")
        
        # Vô hiệu hóa OpenAI và Google nếu thiếu thư viện
        if not Document: # Thiếu docx
             QMessageBox.warning(self, "Thiếu Thư viện", "Cần cài 'python-docx' để xử lý file Word.", QMessageBox.Ok)

        if Translator is None:
             self.combo_translator.setItemData(0, Qt.ItemFlag.NoItemFlags, Qt.ItemDataRole.UserRole - 1)
        
        # Luôn vô hiệu hóa OpenAI tạm thời
        self.combo_translator.setItemData(1, Qt.ItemFlag.NoItemFlags, Qt.ItemDataRole.UserRole - 1)
        
        form_layout.addRow("Công cụ Dịch:", self.combo_translator)

        # 3. Tên File Lưu Mới
        self.save_name_input = QLineEdit()
        self.save_name_input.setPlaceholderText("Ví dụ: HopDongDich_EN.docx")
        form_layout.addRow("Tên File Lưu:", self.save_name_input)

        main_layout.addLayout(form_layout)
        main_layout.addStretch()

        # --- Nút hành động ---
        control_layout = QHBoxLayout()
        self.btn_create = QPushButton("Tạo File")
        self.btn_create.setEnabled(False) # Mặc định vô hiệu hóa
        self.btn_create.setStyleSheet("padding: 6px; background-color: #4CAF50; color: white; border-radius: 5px;")
        self.btn_create.clicked.connect(self.accept_dialog)
        
        btn_cancel = QPushButton("Hủy")
        btn_cancel.clicked.connect(self.reject)
        
        control_layout.addStretch()
        control_layout.addWidget(self.btn_create)
        control_layout.addWidget(btn_cancel)

        main_layout.addLayout(control_layout)

        # Kiểm tra điều kiện ban đầu
        self._update_create_button_state()
        self.combo_translator.currentIndexChanged.connect(self._update_create_button_state)
        self.save_name_input.textChanged.connect(self._update_create_button_state)

    def _select_file(self):
        """Mở hộp thoại để chọn file Word mẫu."""
        src, _ = QFileDialog.getOpenFileName(self, "Chọn File Mẫu", CONTRACT_TEMPLATE_PATH, "Word Files (*.doc *.docx)")
        if src:
            self.src_file_path = src
            self.file_path_label.setText(os.path.basename(src))
            # Gợi ý tên file lưu
            if not self.save_name_input.text():
                 base_name = os.path.splitext(os.path.basename(src))[0]
                 self.save_name_input.setText(f"{base_name}_Translated.docx")
            
            self._update_create_button_state()

    def _update_create_button_state(self):
        """Cập nhật trạng thái của nút 'Tạo File'."""
        is_file_selected = self.src_file_path is not None
        is_save_name_valid = bool(self.save_name_input.text().strip())
        
        is_docx_available = Document is not None
        is_translator_available = (self.combo_translator.currentData() == "google" and Translator is not None)
        
        self.btn_create.setEnabled(is_file_selected and is_save_name_valid and is_docx_available and is_translator_available)

    def accept_dialog(self):
        """Xử lý khi nút 'Tạo File' được bấm."""
        if not self.src_file_path: return

        save_filename = self.save_name_input.text().strip()
        if not save_filename.lower().endswith(('.doc', '.docx')):
             save_filename += ".docx" 
             
        self.save_file_path = os.path.join(CONTRACT_TEMPLATE_PATH, save_filename)
        
        if os.path.exists(self.save_file_path):
            r = QMessageBox.question(self, "Ghi đè", f"File {os.path.basename(self.save_file_path)} đã tồn tại. Ghi đè?")
            if r != QMessageBox.Yes: return
        
        self.accept()
        
    def get_translation_info(self):
        """Trả về thông tin dịch được chọn."""
        if self.result() == QDialog.Accepted:
            return {
                "src_path": self.src_file_path,
                "save_path": self.save_file_path,
                "translator": self.combo_translator.currentData(), # 'google' hoặc 'openai'
            }
        return None
# --- END: TranslationDialog ---


class ContractTemplateWidget(QWidget):
    translation_start_signal = Signal(str, str, str)  # file_path, save_path, translator_engine 

    def __init__(self, translation_handler, parent=None): 
        super().__init__(parent)
        self.translation_handler = translation_handler 
        self.setStyleSheet("background-color: #f7f9fc;")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(6, 6, 6, 6)

        control_row = QHBoxLayout()

        # --- Ô Tìm kiếm (Search Box) ---
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Tìm kiếm mẫu hợp đồng...")
        self.search_box.setStyleSheet(
            "QLineEdit { "
            "   padding: 5px; "
            "   border: 1px solid #ccc; "
            "   border-radius: 5px; "
            "   font-size: 14px; "
            "   max-width: 350px; " 
            "}"
            "QLineEdit:focus { "
            "   border: 1px solid #4a90e2; "
            "   border-width: 1px; "
            "}"
        )
        self.search_box.textChanged.connect(self.load_templates)
        control_row.addWidget(self.search_box)
        
        control_row.addStretch() 

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
        self.btn_translate_contract.setToolTip("Mở popup để chọn file Word mẫu và công cụ dịch.")
        self.btn_translate_contract.clicked.connect(self.open_translate_dialog)
        control_row.addWidget(self.btn_translate_contract)

        main_layout.addLayout(control_row)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")

        h_container = QWidget()
        h_container.setStyleSheet("QWidget { background-color: transparent; }")

        self.layout = QGridLayout(h_container)
        self.layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.layout.setSpacing(10) 
        scroll_area.setWidget(h_container)
        main_layout.addWidget(scroll_area)

        self.load_templates()

    def load_templates(self):
        filter_text = self.search_box.text().lower()
        
        for i in reversed(range(self.layout.count())):
            item = self.layout.itemAt(i)
            if item.widget():
                item.widget().deleteLater()

        if not os.path.isdir(CONTRACT_TEMPLATE_PATH):
            self.layout.addWidget(QLabel(f"Không tìm thấy thư mục: {CONTRACT_TEMPLATE_PATH}"))
            return

        files = [
            f for f in os.listdir(CONTRACT_TEMPLATE_PATH)
            if f.lower().endswith(('.doc', '.docx')) and not f.startswith("~$")
            and filter_text in f.lower() 
        ]
        
        if not files:
            if filter_text:
                self.layout.addWidget(QLabel(f"Không tìm thấy mẫu hợp đồng cho: '{self.search_box.text()}'"))
            else:
                self.layout.addWidget(QLabel("Không có file mẫu hợp đồng."))
            return

        col_count = 5 
        for i, name in enumerate(files):
            path = os.path.join(CONTRACT_TEMPLATE_PATH, name)
            row, col = divmod(i, col_count)
            self.layout.addWidget(self._create_card(name, path), row, col)

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
                # FIX: Sửa lỗi Runtime bằng cách gọi hàm khởi tạo lớp cơ sở chính xác
                super().__init__() 
                self.path = path
                self.setCursor(Qt.PointingHandCursor)
                self.default = """
                    QWidget { 
                        background-color: white; 
                        border: 1px solid #ddd; 
                        border-radius: 5px; 
                        padding: 8px;
                    }
                """
                self.hover = """
                    QWidget { 
                        background-color: #e6f0ff; 
                        border: 1px solid #4a90e2; 
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

    def open_translate_dialog(self):
        """Mở popup TranslationDialog và xử lý kết quả."""
        if Document is None:
            QMessageBox.critical(self, "Thiếu thư viện", "Cần cài python-docx để xử lý file Word.")
            return

        # Tạo và hiển thị dialog
        dialog = TranslationDialog(self)
        if dialog.exec() == QDialog.Accepted:
            info = dialog.get_translation_info()
            if info:
                # Phát tín hiệu với 3 tham số, kích hoạt TranslationHandler
                self.translation_start_signal.emit(info['src_path'], info['save_path'], info['translator'])

    def open_file(self, path):
        try:
            if sys.platform.startswith("win"): os.startfile(path)
            elif sys.platform == "darwin": subprocess.call(["open", path])
            else: subprocess.call(["xdg-open", path])
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở file: {e}")
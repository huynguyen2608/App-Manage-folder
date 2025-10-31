import os
import shutil
import sys
import re
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QProgressBar, 
    QMessageBox, QDialog, QHBoxLayout 
)
from PySide6.QtCore import Qt, Signal, QObject, QCoreApplication, QThread, QPoint
from PySide6.QtGui import QColor, QPalette

# CÔNG CỤ 1: GOOGLETRANS
try:
    from googletrans import Translator
except Exception:
    Translator = None 

# CÔNG CỤ 2: OPENAI
try:
    from openai import OpenAI
    import openai
except Exception:
    OpenAI = None 
    openai = None 

# --- KHỐI IMPORT PYTHON-DOCX (ĐÃ CẬP NHẬT ĐỂ BẮT LỖI CỤ THỂ) ---
try:
    # Lỗi Rgb thường do phiên bản python-docx quá cũ. Lệnh nâng cấp sẽ sửa lỗi này.
    from docx import Document
    from docx.shared import RGBColor, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    # THÊM: Màu Xám nhạt cho chữ dịch (Light Gray)
    TRANSLATION_COLOR = RGBColor(128, 128, 128)
    
except ImportError as e:
    # Bắt lỗi khi không thể import Rgb
    print(f"LỖI IMPORT PYTHON-DOCX: {e}. Vui lòng chạy 'pip install --upgrade python-docx'.", file=sys.stderr)
    Document = None
    RGBColor = None
    Pt = None 
    Rgb = None
    TRANSLATION_COLOR = None 
except Exception as e: 
    # Bắt các lỗi khác (ví dụ: lỗi nếu Rgb vẫn là None)
    print(f"LỖI KHÁC KHI IMPORT DOCX: {e}", file=sys.stderr)
    Document = None
    RGBColor = None
    Pt = None 
    Rgb = None
    TRANSLATION_COLOR = None
# ------------------------------------------------------------------

# Thiết lập một hệ thống nhắc (Prompt) rõ ràng cho AI
SYSTEM_PROMPT = (
    "Bạn là một chuyên gia dịch thuật pháp lý. Nhiệm vụ của bạn là dịch chính xác "
    "từng đoạn văn bản Hợp đồng từ Tiếng Việt sang Tiếng Anh, duy trì cấu trúc "
    "pháp lý và định dạng nguyên bản. CHỈ trả về đoạn văn bản đã dịch."
)


# --- Translation Worker Thread (CHỨC NĂNG DỊCH TRONG LUỒNG RIÊNG) ---
class TranslationWorker(QObject):
    finished = Signal(bool) # True if successful, False if cancelled
    error = Signal(str)
    progress_update = Signal(int, int, str) # current, total, text

    def __init__(self, file_path: str, save_path: str, method: str, client_or_translator, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.save_path = save_path
        self.method = method # 'ai' hoặc 'google'
        self.client_or_translator = client_or_translator 
        self._is_cancelled = False
        self.total_items = 0
        self.translated_count = 0
        self.model = "gpt-3.5-turbo" 

    def cancel(self):
        self._is_cancelled = True

    # ------------------ Logic dịch OpenAI ------------------
    def _call_openai_api(self, text_to_translate: str) -> str:
        """Thực hiện cuộc gọi API đến OpenAI."""
        if self._is_cancelled:
            return ""

        try:
            response = self.client_or_translator.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": f"Dịch sang Tiếng Anh đoạn sau: {text_to_translate}"}
                ],
                temperature=0.1, 
            )
            return response.choices[0].message.content.strip()
        except openai.AuthenticationError:
            raise Exception("Lỗi Xác thực OpenAI: Kiểm tra OPENAI_API_KEY.")
        except Exception as e:
            print(f"Lỗi API OpenAI: {e}")
            raise Exception(f"Lỗi API OpenAI: {str(e)}")

    # ------------------ Logic dịch Google Translate ------------------
    def _call_googletrans_api(self, text_to_translate: str) -> str:
        """Thực hiện dịch bằng Google Translate."""
        if self._is_cancelled:
            return ""
        
        try:
            # Dịch từ Tiếng Việt (vi) sang Tiếng Anh (en)
            result = self.client_or_translator.translate(text_to_translate, src='vi', dest='en')
            return result.text
        except Exception as e:
            print(f"Lỗi Google Translate: {e}")
            raise Exception(f"Lỗi Google Translate: Không kết nối được hoặc dữ liệu lớn.")


    def run(self):
        if Document is None:
            # Sẽ chỉ xảy ra nếu import docx thất bại
            self.error.emit("Thiếu thư viện python-docx.")
            self.finished.emit(False)
            return
        if self.client_or_translator is None:
            self.error.emit(f"Không tìm thấy Client/Translator cho phương thức {self.method}.")
            self.finished.emit(False)
            return

        # Chọn hàm dịch
        if self.method == 'ai':
            translation_func = self._call_openai_api
        elif self.method == 'google':
            translation_func = self._call_googletrans_api
        else:
            self.error.emit("Phương thức dịch không hợp lệ.")
            self.finished.emit(False)
            return

        try:
            # 1. Đọc file gốc
            src_doc = Document(self.file_path)
            new_doc = Document() # Tài liệu mới để xây dựng nội dung

            # Lọc các đoạn có nội dung
            all_paragraphs = [p for p in src_doc.paragraphs if p.text.strip()]
            self.total_items = len(all_paragraphs)
            self.translated_count = 0
            
            for i, p_original in enumerate(all_paragraphs):
                if self._is_cancelled:
                    self.finished.emit(False)
                    return

                original_text = p_original.text
                
                # 2. Thêm đoạn gốc (Tiếng Việt) với định dạng gốc
                p_vn = new_doc.add_paragraph()
                p_vn.text = original_text
                # Copy định dạng căn lề và style
                p_vn.paragraph_format.alignment = p_original.paragraph_format.alignment
                p_vn.style = p_original.style
                
                # 3. Thông báo tiến độ
                self.translated_count += 1
                self.progress_update.emit(
                    self.translated_count, 
                    self.total_items, 
                    f"Đang dịch đoạn {self.translated_count}/{self.total_items} bằng {self.method.upper()}: {original_text[:50]}..."
                )
                QCoreApplication.processEvents() # Cập nhật UI

                # 4. Gọi hàm dịch đã chọn
                translated_text = translation_func(original_text)
                
                if translated_text:
                    # 5. Thêm đoạn dịch (Tiếng Anh) ngay bên dưới
                    p_en = new_doc.add_paragraph()
                    p_en.style = p_original.style  # Giữ nguyên style đoạn gốc

                    # Sao chép định dạng căn lề, khoảng cách...
                    pf = p_original.paragraph_format
                    p_en.paragraph_format.alignment = pf.alignment
                    p_en.paragraph_format.left_indent = pf.left_indent
                    p_en.paragraph_format.right_indent = pf.right_indent
                    p_en.paragraph_format.first_line_indent = pf.first_line_indent
                    p_en.paragraph_format.space_before = pf.space_before
                    p_en.paragraph_format.space_after = pf.space_after
                    p_en.paragraph_format.line_spacing = pf.line_spacing

                    # --- Sao chép định dạng chữ từ các "run" trong đoạn gốc ---
                    if p_original.runs:
                        # Nếu có nhiều run, copy font định dạng từ run đầu tiên
                        for run_orig in p_original.runs:
                            run_new = p_en.add_run(translated_text)
                            # Sao chép định dạng gốc
                            run_new.bold = run_orig.bold
                            run_new.underline = run_orig.underline
                            run_new.font.name = run_orig.font.name
                            run_new.font.size = run_orig.font.size
                            # Chỉ khác biệt: nghiêng và màu xám
                            run_new.italic = True
                            if TRANSLATION_COLOR:
                                run_new.font.color.rgb = TRANSLATION_COLOR
                            break  # chỉ cần lấy format từ run đầu tiên
                    else:
                        # Nếu đoạn gốc không có run
                        run_new = p_en.add_run(translated_text)
                        run_new.italic = True
                        if TRANSLATION_COLOR:
                            run_new.font.color.rgb = TRANSLATION_COLOR

                        
                
            # 6. Lưu file
            new_doc.save(self.save_path)

            self.finished.emit(True)

        except Exception as e:
            self.error.emit(str(e))
            self.finished.emit(False)


# --- CustomProgressDialog và các lớp UI liên quan (GIỮ NGUYÊN) ---

class CustomProgressDialog(QDialog):
    canceled = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(400, 150)
        self.setWindowTitle("Xử lý Dịch Thuật")
        
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        
        self.lbl_title = QLabel("Xử lý Dịch Hợp đồng")
        self.lbl_title.setStyleSheet("font-size: 14pt; font-weight: bold; color: #333;")
        
        self.lbl_status = QLabel("Đang chuẩn bị...")
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setStyleSheet("font-size: 10pt; color: #666;")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setFormat("%p%") 
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setTextVisible(True)
        
        self.btn_cancel = QPushButton("Hủy") 
        self.btn_cancel.setFixedWidth(100)
        self.btn_cancel.clicked.connect(self.reject) 
        
        layout.addWidget(self.lbl_title, alignment=Qt.AlignCenter)
        layout.addWidget(self.lbl_status)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.btn_cancel, alignment=Qt.AlignRight)
        
    def update_progress(self, current, total, text):
        if total > 0:
            self.progress_bar.setRange(0, total)
            self.progress_bar.setValue(current)
            
            percentage = (current / total) * 100 if total > 0 else 0
            self.progress_bar.setFormat(f"{percentage:.1f}%")

        self.lbl_status.setText(text)

    def closeEvent(self, event):
        event.ignore() 

class LoadingOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAutoFillBackground(True)
        self.setPalette(QPalette(QColor(0, 0, 0, 128)))

    def resizeEvent(self, event):
        if self.parentWidget():
            self.resize(self.parentWidget().size())

    def showEvent(self, event):
        if self.parentWidget():
            self.setGeometry(self.parentWidget().rect())
        super().showEvent(event)

class TranslationHandler(QObject):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.worker_thread = None
        self.worker = None
        
        self.openai_client = None
        self.google_translator = None

        try:
            if OpenAI:
                self.openai_client = OpenAI()
        except Exception as e:
            print(f"CẢNH BÁO: Không thể khởi tạo OpenAI Client. Lỗi: {e}")

        try:
            if Translator:
                self.google_translator = Translator()
        except Exception as e:
            print(f"CẢNH BÁO: Không thể khởi tạo Google Translator. Lỗi: {e}")


        self.progress_dialog = CustomProgressDialog(main_window)

    def is_available(self, method: str):
        if Document is None:
            return False, "Thiếu thư viện python-docx."
            
        if method == 'ai':
            if self.openai_client is not None:
                return True, ""
            return False, "Thiếu OpenAI Client (OPENAI_API_KEY chưa thiết lập)."
        
        elif method == 'google':
            if self.google_translator is not None:
                return True, ""
            return False, "Thiếu Google Translator (googletrans chưa được cài)."
        
        return False, "Phương thức dịch không hợp lệ."

    def start_translation(self, file_path: str, save_path: str, method: str):
        is_ok, msg = self.is_available(method)
        if not is_ok:
            QMessageBox.critical(self.main_window, "Thiếu Thiết lập", f"Lỗi: {msg}")
            return
            
        self.progress_dialog.rejected.connect(self._handle_cancellation) 

        if method == 'ai':
            client_or_translator = self.openai_client
        elif method == 'google':
            client_or_translator = self.google_translator
        
        self.worker_thread = QThread()
        self.worker = TranslationWorker(file_path, save_path, method, client_or_translator)
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._handle_completion)
        self.worker.error.connect(self._handle_error)
        self.worker.progress_update.connect(self.progress_dialog.update_progress)
        
        self.worker_thread.start()
        
        self.progress_dialog.exec() 

    def _handle_cancellation(self):
        if self.worker:
            self.worker.cancel()
            QMessageBox.information(self.main_window, "Hủy", "Đang chờ luồng dịch kết thúc...")

    def _handle_completion(self, success: bool):
        self._cleanup_worker()
        self.progress_dialog.hide() 
        if success:
            QMessageBox.information(self.main_window, "Hoàn thành", f"Hợp đồng đã được dịch và lưu thành công.")
        elif self.worker is not None and not self.worker._is_cancelled:
            QMessageBox.warning(self.main_window, "Bị Hủy", "Quá trình dịch đã bị hủy.")


    def _handle_error(self, message: str):
        self._cleanup_worker()
        self.progress_dialog.hide() 
        QMessageBox.critical(self.main_window, "Lỗi Dịch Thuật", message)

    def _cleanup_worker(self):
        if self.worker_thread:
            self.worker_thread.quit()
            self.worker_thread.wait() 
        self.worker_thread = None
        self.worker = None
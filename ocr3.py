import sys
import cv2
import pytesseract
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QLabel,
                             QPushButton, QFileDialog, QTextEdit, QVBoxLayout,
                             QHBoxLayout, QSplitter, QScrollArea, QComboBox, QLabel)
from PyQt6.QtGui import QPixmap, QImage
from PyQt6.QtCore import Qt
from docx import Document
from fpdf import FPDF

pytesseract.pytesseract.tesseract_cmd = r'E:\Tesseract\tesseract.exe'


class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.image_path = None
        self.psm_value = 6  # Значение по умолчанию

    def initUI(self):
        self.setWindowTitle('OCR Программа')
        self.setGeometry(100, 100, 1400, 900)  # Изменен размер окна

        # Основной виджет и layout
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(5, 5, 5, 5)

        # Панель инструментов
        toolbar = QHBoxLayout()
        self.load_button = QPushButton('Загрузить изображение')
        self.load_button.clicked.connect(self.loadImage)

        self.ocr_button = QPushButton('Распознать текст')
        self.ocr_button.clicked.connect(self.performOCR)

        self.save_txt_button = QPushButton('Сохранить в TXT')
        self.save_txt_button.clicked.connect(lambda: self.saveText('txt'))

        self.save_docx_button = QPushButton('Сохранить в DOCX')
        self.save_docx_button.clicked.connect(lambda: self.saveText('docx'))

        self.save_pdf_button = QPushButton('Сохранить в PDF')
        self.save_pdf_button.clicked.connect(lambda: self.saveText('pdf'))

        # Добавление выпадающего списка для выбора psm
        self.psm_label = QLabel("Выберите PSM:")
        self.psm_combo = QComboBox()
        for i in range(14):  # от 0 до 13
            self.psm_combo.addItem(f"PSM {i}", i)
        self.psm_combo.setCurrentIndex(6)  # Значение по умолчанию — PSM 6
        self.psm_combo.currentIndexChanged.connect(self.updatePSM)

        toolbar.addWidget(self.load_button)
        toolbar.addWidget(self.ocr_button)
        toolbar.addWidget(self.save_txt_button)
        toolbar.addWidget(self.save_docx_button)
        toolbar.addWidget(self.save_pdf_button)
        toolbar.addWidget(self.psm_label)
        toolbar.addWidget(self.psm_combo)

        main_layout.addLayout(toolbar)

        # Разделение на две области (70/30)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Левая панель - изображение (70%)
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        scroll_area_left = QScrollArea()
        scroll_area_left.setWidget(self.image_label)
        scroll_area_left.setWidgetResizable(True)
        scroll_area_left.setMinimumWidth(800)  # Минимальная ширина
        splitter.addWidget(scroll_area_left)

        # Правая панель - результат (30%)
        self.text_edit = QTextEdit()
        scroll_area_right = QScrollArea()
        scroll_area_right.setWidget(self.text_edit)
        scroll_area_right.setWidgetResizable(True)
        scroll_area_right.setMinimumWidth(400)  # Минимальная ширина
        splitter.addWidget(scroll_area_right)

        # Установка пропорций 70/30
        splitter.setStretchFactor(0, 7)
        splitter.setStretchFactor(1, 3)

        main_layout.addWidget(splitter)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def loadImage(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Выберите изображение', '',
                                                   'Изображения (*.png *.jpg *.jpeg *.bmp)')
        if file_path:
            self.image_path = file_path
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap.scaled(
                self.image_label.width(),
                self.image_label.height(),
                Qt.AspectRatioMode.KeepAspectRatio))

    def updatePSM(self):
        # Обновление выбранного значения PSM
        self.psm_value = self.psm_combo.currentData()

    def performOCR(self):
        if not self.image_path:
            return

        image = cv2.imread(self.image_path)

        # Преобразование изображения в оттенки серого
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Простая бинаризация для улучшения качества
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # Использование выбранного значения PSM для Tesseract
        config = f'--psm {self.psm_value}'  # Параметр psm изменяется в зависимости от выбора пользователя
        text = pytesseract.image_to_string(binary, lang='eng+rus', config=config)

        self.text_edit.setText(text)

    def saveText(self, format):
        text = self.text_edit.toPlainText()
        if not text:
            return

        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getSaveFileName(self, 'Сохранить файл', '', f'Файл (*.{format})')

        if file_path:
            if format == 'txt':
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(text)
            elif format == 'docx':
                doc = Document()
                doc.add_paragraph(text)
                doc.save(file_path)
            elif format == 'pdf':
                pdf = FPDF()
                pdf.add_page()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.set_font('Arial', size=12)
                pdf.multi_cell(0, 10, text)
                pdf.output(file_path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = OCRApp()
    ex.show()
    sys.exit(app.exec())

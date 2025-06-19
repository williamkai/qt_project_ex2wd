import os
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFontMetrics
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QApplication, QLabel, QFileDialog,
    QSizePolicy, QLineEdit, QHBoxLayout, QFormLayout, QMessageBox,QProgressBar
)

class BaseFuncWindow(QWidget):
    def __init__(self, title):
        super().__init__()
        self.setWindowTitle(title)
        self.setMinimumSize(600, 400)

        self.excel_path = None
        self.word_path = None
        self.output_folder = None

        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        def add_labeled_button(label_text, button_text, click_func):
            container = QHBoxLayout()

            label = QLabel(label_text)
            label.setFixedWidth(100)
            label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            container.addWidget(label)

            button = QPushButton(button_text)
            button.setFixedSize(160, 30)
            button.clicked.connect(click_func)
            container.addWidget(button)

            wrapper = QWidget()
            wrapper.setLayout(container)
            layout.addWidget(wrapper, alignment=Qt.AlignmentFlag.AlignCenter)

            return button

        self.btn_select_excel = add_labeled_button("Excel 檔案：", "選擇 Excel 檔案", self.select_excel_file)
        self.btn_select_word = add_labeled_button("Word 模板：", "選擇 Word 模板", self.select_word_file)
        self.btn_select_output = add_labeled_button("輸出資料夾：", "選擇輸出資料夾", self.select_output_folder)

        self.setLayout(layout)

    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 Excel 檔案", "", "Excel 檔案 (*.xls *.xlsx)")
        if path:
            self.excel_path = path
            filename = os.path.basename(path)
            self.btn_select_excel.setText(f"Excel: {filename}")

    def select_word_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 Word 模板", "", "Word 文件 (*.docx)")
        if path:
            self.word_path = path
            filename = os.path.basename(path)
            self.btn_select_word.setText(filename)

    def select_pdf_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 pdf 模板", "", "pdf 文件 (*.pdf)")
        if path:
            self.word_path = path
            filename = os.path.basename(path)
            self.btn_select_word.setText(filename)

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "選擇輸出資料夾")
        if folder:
            self.output_folder = folder
            metrics = QFontMetrics(self.btn_select_output.font())
            max_width = 300
            elided_text = metrics.elidedText(folder, Qt.TextElideMode.ElideMiddle, max_width)
            self.btn_select_output.setText(elided_text)
            self.btn_select_output.setMaximumWidth(max_width)
            self.btn_select_output.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)


    def setup_ui_pdf(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        def add_labeled_button(label_text, button_text, click_func):
            container = QHBoxLayout()

            label = QLabel(label_text)
            label.setFixedWidth(100)
            label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            container.addWidget(label)

            button = QPushButton(button_text)
            button.setFixedSize(160, 30)
            button.clicked.connect(click_func)
            container.addWidget(button)

            wrapper = QWidget()
            wrapper.setLayout(container)
            layout.addWidget(wrapper, alignment=Qt.AlignmentFlag.AlignCenter)

            return button

        self.btn_select_excel = add_labeled_button("Excel 檔案：", "選擇 Excel 檔案", self.select_excel_file)
        self.btn_select_word = add_labeled_button("Pdf 模板：", "選擇 pdf 模板", self.select_pdf_file)
        self.btn_select_output = add_labeled_button("輸出資料夾：", "選擇輸出資料夾", self.select_output_folder)

        self.setLayout(layout)
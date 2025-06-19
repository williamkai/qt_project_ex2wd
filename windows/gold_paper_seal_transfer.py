import os
from docx import Document
from copy import deepcopy
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QPushButton, QApplication, QLabel, QComboBox,
    QLineEdit, QHBoxLayout,  QMessageBox,QProgressBar
)
from core.base_windows import BaseFuncWindow
from core.conversion_utils import (
    read_data_auto,
    duplicate_table_and_insert,
    map_all_placeholders,
    fill_data_to_table_v2,
    get_dynamic_font_size_func
)

class GoldPaperSealTransferWindow(BaseFuncWindow):
    def __init__(self, title="試算表轉金紙封條"):
        super().__init__(title)
        self.is_closing = False
        self.setMinimumSize(600, 620)

    def closeEvent(self, event):
        if not self.btn_run.isEnabled():
            self.is_closing = True
            QMessageBox.information(self, "提醒", "正在轉換中，已中止處理。")
        else:
            self.is_closing = True
        event.accept()

    def setup_ui(self):
        super().setup_ui_pdf()
        layout = self.layout()

        def add_labeled_input(label_text, placeholder):
            row = QHBoxLayout()
            label = QLabel(label_text)
            label.setFixedWidth(150)
            label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            row.addWidget(label)

            line_edit = QLineEdit()
            line_edit.setPlaceholderText(placeholder)
            line_edit.setFixedWidth(200)
            row.addWidget(line_edit)

            wrapper = QWidget()
            wrapper.setLayout(row)
            layout.addWidget(wrapper, alignment=Qt.AlignmentFlag.AlignCenter)

            return line_edit

        def add_labeled_combobox(label_text, options, default_index=0):
            row = QHBoxLayout()
            label = QLabel(label_text)
            label.setFixedWidth(150)
            label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            row.addWidget(label)

            combo = QComboBox()
            combo.addItems([str(opt) for opt in options])
            combo.setCurrentIndex(default_index)
            combo.setFixedWidth(80)
            row.addWidget(combo)

            wrapper = QWidget()
            wrapper.setLayout(row)
            self.layout().addWidget(wrapper, alignment=Qt.AlignmentFlag.AlignCenter)

            return combo


        # 主要 UI 控件建立（照第一個功能邏輯）
        self.input_required = add_labeled_input("必須欄位：", "請輸入欄位，例如: b,c,e")
        self.input_optional = add_labeled_input("選擇欄位：", "例如: f,g,h")

        self.limit_rows_combo = add_labeled_combobox(
            "筆數選擇：", ["15","30","45","90","200", "400", "600", "800", "1000", "2000", "4000", "全部"], 2
        )

        # 執行按鈕
        self.btn_run = QPushButton("執行轉換")
        self.btn_run.setFixedSize(200, 40)
        self.btn_run.clicked.connect(self.run_conversion)
        layout.addWidget(self.btn_run, alignment=Qt.AlignmentFlag.AlignCenter)

        # 進度條
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(300)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar, alignment=Qt.AlignmentFlag.AlignCenter)

    def run_conversion(self):
        """
        讀取 Excel 檔案，依指定欄位將資料寫進去;
        pdf設定位子，然後將資料填入 pdf 模板中。
        這個功能是將試算表轉換為金紙封條的功能
        """
        pass
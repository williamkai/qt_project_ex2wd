from PyQt6.QtCore import Qt ,QThreadPool
from PyQt6.QtWidgets import (
    QWidget, QPushButton,  QLabel, QComboBox,
    QLineEdit, QHBoxLayout,  QMessageBox,QProgressBar
)
from core.base_windows import BaseFuncWindow
from core.conversion_utils import read_data_auto

from core.kai_thread_pool import WordExportWorker  # 假設你放這邊

class ExcelToInvitationWindow(BaseFuncWindow):
    def __init__(self, title="試算表轉召請文功能"):
        super().__init__(title)
        self.is_closing = False
        self.thread_pool = QThreadPool()
        self.setMinimumSize(600, 620)

    def closeEvent(self, event):
        if not self.btn_run.isEnabled():
            self.is_closing = True
            QMessageBox.information(self, "提醒", "正在轉換中，已中止處理。")
        else:
            self.is_closing = True
        event.accept()

    def setup_ui(self):
        super().setup_ui()

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

        self.input_required = add_labeled_input("必須欄位：", "請輸入欄位，例如: c,d")
        self.input_optional = add_labeled_input("選擇欄位：", "請輸入欄位，例如: e,f")
        # 在 setup_ui 中加入
        self.font_size_1 = add_labeled_combobox("8字以內字體大小：", [8,10,12,14,16,18,20,22,25,28,30], default_index=7)  # 預設22
        self.font_size_2 = add_labeled_combobox("9~20字字體大小：", [8,10,12,14,16,18,20,22,25,28,30], default_index=5)  # 預設18
        self.font_size_3 = add_labeled_combobox("21字以上字體大小：", [6,8,10,12,14,16,18,20], default_index=3)       # 預設12
        self.limit_rows_combo = add_labeled_combobox(
            "筆數選擇：", ["200", "400", "600","800","1000","2000","4000", "全部"], default_index=0
        )

        self.btn_run = QPushButton("執行轉換")
        self.btn_run.setFixedSize(200, 40)
        self.btn_run.clicked.connect(self.run_conversion)
        layout.addWidget(self.btn_run, alignment=Qt.AlignmentFlag.AlignCenter)

        # ✅ 新增進度條
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(300)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar, alignment=Qt.AlignmentFlag.AlignCenter)

    def run_conversion(self):
        self.btn_run.setEnabled(False)
        self.btn_run.setText("轉換中...")
        self.progress_bar.setValue(0)

        try:
            # 驗證參數
            if not self.excel_path or not self.word_path or not self.output_folder:
                raise ValueError("請確認 Excel、Word 模板與輸出資料夾皆已選擇")

            required_text = self.input_required.text().strip()
            optional_text = self.input_optional.text().strip()
            if not required_text:
                raise ValueError("請輸入必須欄位。")

            required_cols = [c.strip().upper() for c in required_text.split(",") if c.strip()]
            optional_cols = [c.strip().upper() for c in optional_text.split(",") if c.strip()] if optional_text else []

            limit_text = self.limit_rows_combo.currentText()
            limit_rows = None if limit_text == "全部" else int(limit_text)

            # 計算有效筆數並設定進度條最大值
            max_progress = self.count_valid_rows(self.excel_path, required_cols, limit_rows)
            self.progress_bar.setMaximum(max_progress)

            font_size_rules = [
                (8, int(self.font_size_1.currentText())),
                (20, int(self.font_size_2.currentText())),
                (9999, int(self.font_size_3.currentText())),
            ]

            worker = WordExportWorker(
                excel_path=self.excel_path,
                word_path=self.word_path,
                output_folder=self.output_folder,
                required_cols=required_cols,
                optional_cols=optional_cols,
                font_size_rules=font_size_rules,
                limit_rows=limit_rows,
            )
            worker.signals.progress.connect(self.progress_bar.setValue)
            worker.signals.finished.connect(self.export_done)

            self.thread_pool.start(worker)

        except ValueError as e:
            QMessageBox.warning(self, "輸入錯誤", str(e))
            self.btn_run.setEnabled(True)
            self.btn_run.setText("轉換")

    def count_valid_rows(self, excel_path, required_cols, limit_rows):
        import pandas as pd
        df = read_data_auto(excel_path).fillna('')
        if limit_rows:
            df = df.head(limit_rows)
        col_letter_map = {chr(65+i): name for i, name in enumerate(df.columns)}
        valid_rows = [row for idx, row in df.iterrows() if all(str(row[col_letter_map[c]]) for c in required_cols)]
        return len(valid_rows)


    def export_done(self, success, message):
        self.btn_run.setEnabled(True)
        self.btn_run.setText("轉換")
        if success:
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.critical(self, "錯誤", message)


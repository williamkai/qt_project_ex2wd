import os
from docx import Document
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

class InvitationTransferWindow(BaseFuncWindow):
    def __init__(self, title="試算表轉拔薦文疏-本"):
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

        # 主要 UI 控件建立（照第一個功能邏輯）
        self.input_required = add_labeled_input("必須欄位：", "請輸入欄位，例如: b,c,e")
        self.input_optional = add_labeled_input("拔薦合併欄：", "例如: f,g,h")

        self.font_size_1 = add_labeled_combobox("6字以內字體大小：", [4,6,8,10,12,14,16,18,20,22,25,28,30], 2)
        self.font_size_2 = add_labeled_combobox("7~15字字體大小：", [4,6,8,10,12,14,16,18,20,22,25,28,30], 2)
        self.font_size_3 = add_labeled_combobox("16字以上字體大小：", [4,6,8,10,12,14,16,18,20], 1)

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

    def get_font_size_func_from_ui(self):
        size_1 = int(self.font_size_1.currentText())
        size_2 = int(self.font_size_2.currentText())
        size_3 = int(self.font_size_3.currentText())
        return get_dynamic_font_size_func(size_1, size_2, size_3)


    def run_conversion(self):
        """
        讀取 Excel 檔案，依指定欄位將資料寫進去，由右寫至左;
        欄位不夠則複製原本的表格在下一頁貼上原始表格繼續處理資料。
        直到處理完畢在保存word檔案
        """
        self.btn_run = self.sender()  # 取得觸發此函式的按鈕物件
        self.btn_run.setEnabled(False)  # 轉換中禁用按鈕避免重複觸發
        self.btn_run.setText("轉換中...")  # 顯示轉換狀態文字
        cancelled = False  # 是否使用者中途取消

        try:
            if not self.excel_path:
                raise ValueError("請選擇 Excel 檔案。")
            if not self.word_path:
                raise ValueError("請選擇 Word 模板檔案。")
            if not self.output_folder:
                raise ValueError("請選擇輸出資料夾。")

            required_text = self.input_required.text().strip()
            if not required_text:
                raise ValueError("請輸入必須欄位。")

            if not all(col.strip().isalpha() and len(col.strip()) == 1 for col in required_text.split(",")):
                raise ValueError("必須欄位必須為英文單一字母，以逗號分隔。")

            required_cols = [c.strip().upper() for c in required_text.split(",") if c.strip()]
            optional_text = self.input_optional.text().strip()
            optional_cols = [c.strip().upper() for c in optional_text.split(",") if c.strip()] if optional_text else []

            df = read_data_auto(self.excel_path).fillna('')
            col_letter_map = {chr(65+i): name for i, name in enumerate(df.columns)}

            limit_text = self.limit_rows_combo.currentText()
            if limit_text != "全部":
                try:
                    limit = int(limit_text)
                    df = df.head(limit)
                except ValueError:
                    raise ValueError("筆數限制不是有效數字")

            total_rows = len(df)
            self.progress_bar.setMaximum(total_rows)

            for col in required_cols + optional_cols:
                if col not in col_letter_map:
                    raise ValueError(f"欄位 {col} 不存在於讀取的資料中。")

            doc = Document(self.word_path)
            if len(doc.tables) == 0:
                raise ValueError("找不到 Word 表格")

            table = doc.tables[0]
            placeholder_map, start_col = map_all_placeholders(table)

            all_data = []
            for idx, row in df.iterrows():
                if self.is_closing:
                    cancelled = True
                    break
                self.progress_bar.setValue(idx + 1)
                QApplication.processEvents()

                if not all(str(row[col_letter_map[c]]) for c in required_cols):
                    continue

                replacements = {c: str(row[col_letter_map[c]]) for c in required_cols + optional_cols}
                all_data.append(replacements)

            font_size_func = self.get_font_size_func_from_ui()

            batch_start = 0
            current_table = table
            while batch_start < len(all_data):
                remaining = len(all_data) - batch_start
                fill_count = min(15, remaining)
                written = fill_data_to_table_v2(
                    current_table,
                    placeholder_map,
                    all_data[batch_start:batch_start+fill_count],
                    start_col,
                    font_size_func=font_size_func
                )
                batch_start += written
                if batch_start < len(all_data):
                    current_table = duplicate_table_and_insert(doc, table)

            out_path = os.path.join(self.output_folder, "牌位文疏批次生成.docx")
            doc.save(out_path)

            if cancelled:
                self.progress_bar.setValue(0)
            else:
                QMessageBox.information(self, "成功", "轉換完成！")
                self.progress_bar.setValue(total_rows)

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"轉換失敗：{str(e)}")
            self.progress_bar.setValue(0)

        finally:
            self.btn_run.setEnabled(True)
            self.btn_run.setText("轉換")

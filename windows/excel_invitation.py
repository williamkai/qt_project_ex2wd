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
    col_letter_to_index,
    prepare_assign_map,
    replace_placeholders,
    save_doc_with_name,
)

class ExcelToInvitationWindow(BaseFuncWindow):
    def __init__(self, title="試算表轉召請文功能"):
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
        """
        讀取 Excel 檔案，依指定欄位產生 Word 文件。
        透過 col_letter_map 統一欄位字母對應欄位名稱，避免錯誤。
        """

        self.btn_run = self.sender()  # 取得觸發此函式的按鈕物件
        self.btn_run.setEnabled(False)  # 轉換中禁用按鈕避免重複觸發
        self.btn_run.setText("轉換中...")  # 顯示轉換狀態文字
        cancelled = False  # 是否使用者中途取消

        try:
            # 檢查必要參數是否都有設定
            if not self.excel_path:
                raise ValueError("請選擇 Excel 檔案。")
            if not self.word_path:
                raise ValueError("請選擇 Word 模板檔案。")
            if not self.output_folder:
                raise ValueError("請選擇輸出資料夾。")

            # 取得使用者輸入的必填欄位（以英文單一字母逗號分隔）
            required_text = self.input_required.text().strip()
            if not required_text:
                raise ValueError("請輸入必須欄位。")

            # 驗證必填欄位格式（必須為單一英文字母）
            if not all(col.strip().isalpha() and len(col.strip()) == 1 for col in required_text.split(",")):
                raise ValueError("必須欄位必須為英文單一字母，以逗號分隔。")

            required_cols = [c.strip().upper() for c in required_text.split(",") if c.strip()]

            # 取得選填欄位，格式同必填欄位，若無輸入則空列表
            optional_text = self.input_optional.text().strip()
            optional_cols = [c.strip().upper() for c in optional_text.split(",") if c.strip()] if optional_text else []

            try:
                # 自訂函式讀取 Excel，並將 NaN 用空字串代替
                df = read_data_auto(self.excel_path)
                df = df.fillna('')

                # 建立欄位字母（A,B,C...）對應實際欄位名稱的字典
                # 這樣能確保對應準確且不受欄位順序影響
                col_letter_map = {chr(65+i): name for i, name in enumerate(df.columns)}

                # 處理筆數限制，若用戶選擇非「全部」時取前 n 筆
                limit_text = self.limit_rows_combo.currentText()
                if limit_text != "全部":
                    try:
                        limit = int(limit_text)
                        df = df.head(limit)
                    except ValueError:
                        raise ValueError("筆數限制不是有效數字")
                    
                total_rows = len(df)
                self.progress_bar.setMaximum(total_rows)

                # 檢查所有必填及選填欄位字母是否存在於 col_letter_map
                for col in required_cols + optional_cols:
                    if col not in col_letter_map:
                        raise ValueError(f"欄位 {col} 不存在於讀取的資料中。")

                # 字體大小規則設定（可依照介面數值調整）
                font_size_rules = [
                    (8, int(self.font_size_1.currentText())),
                    (20, int(self.font_size_2.currentText())),
                    (9999, int(self.font_size_3.currentText())),
                ]

                # 開始逐筆處理 DataFrame 中的每一列
                for idx, row in df.iterrows():
                    # 檢查是否使用者已關閉視窗，中斷處理
                    if self.is_closing:
                        cancelled = True
                        break

                    self.progress_bar.setValue(idx + 1)
                    QApplication.processEvents()  # 更新 UI，避免卡死

                    # 檢查必填欄位是否皆有值，若有空值跳過該列
                    if not all(str(row[col_letter_map[c]]) for c in required_cols):
                        continue

                    # 使用 Word 模板建立文件物件
                    doc = Document(self.word_path)

                    # 產生取代字典：欄位字母 => 欄位值（字串）
                    replacements = {
                        c: str(row[col_letter_map[c]])
                        for c in required_cols + optional_cols
                    }

                    # 呼叫你自訂的函式替換 Word 中的對應標記，並套用字體大小規則
                    assign_map = prepare_assign_map(replacements, doc)
                    replace_placeholders(doc, assign_map, font_size_rules)

                    # 以第一個必填欄位值前 11 個字作為檔名主體
                    c_val = str(row[col_letter_map[required_cols[0]]])[:11]
                    filename = f"{c_val}_召請文.docx"
                    folder_name = os.path.join(self.output_folder, "Excel 轉 Word 召請文")

                    # 儲存文件到指定路徑
                    save_doc_with_name(doc, folder_name, filename)

                if cancelled:
                    self.progress_bar.setValue(0)
                else:
                    QMessageBox.information(self, "成功", "轉換完成！")
                    self.progress_bar.setValue(total_rows)

            except Exception as e:
                # 捕捉過程中任何錯誤，顯示錯誤訊息並重置進度條
                QMessageBox.critical(self, "錯誤", f"轉換失敗：{str(e)}")
                self.progress_bar.setValue(0)

        except ValueError as ve:
            QMessageBox.warning(self, "輸入錯誤", str(ve))
            return
        
        finally:
            # 無論成功失敗，最後都要恢復按鈕狀態
            self.btn_run.setEnabled(True)
            self.btn_run.setText("轉換")


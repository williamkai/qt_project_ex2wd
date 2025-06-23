import os
from io import BytesIO
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QLineEdit, QGraphicsTextItem, QComboBox,QSplitter,QGridLayout,QSizePolicy,QFrame,QMessageBox,QGraphicsSimpleTextItem
)
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QPixmap, QImage, QFont, QTransform, QPainter, QPen,QBrush
from pdf2image import convert_from_path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from pypdf import PdfReader, PdfWriter, PageObject
from PIL import Image
from collections import defaultdict

from .gold_ui_parts import (
    build_top_toolbar,
    build_horizontal_separator,
    build_splitter,
)
from modules.pdf_viewer import PDFViewer
from modules.label_manager import LabelManager
from core.conversion_utils import (
    read_data_auto,)
from core.pdf_exporter import PDFExporter
# 註冊中文字型
# 動態取得字體檔案的路徑（以目前這個檔案為基準）
current_dir = os.path.dirname(os.path.abspath(__file__))
font_path = os.path.join(current_dir, "..", "core", "Iansui-Regular.ttf")

# 正常化路徑並註冊字型
font_path = os.path.normpath(font_path)
pdfmetrics.registerFont(TTFont("Iansui", font_path))

# font_path = get_project_path("core", "Iansui-Regular.ttf")


class GoldPaperSealTransferWindow(QWidget):
    def __init__(self, title="試算表轉金紙封條"):
        super().__init__()
        self.setWindowTitle(title)
        self.is_closing = False
        self.setWindowTitle("PDF 可視化標籤定位工具")
        self.setMinimumSize(1500, 700)

        # PDF 狀態與圖像資訊
        self.pdf_path = None
        self.current_image = None
        self.image_width = None
        self.image_height = None
        self.pdf_pixmap_item = None
        self.excel_path = None

        # 場景與視圖
        self.scene = QGraphicsScene()
        self.graphics_view = QGraphicsView(self.scene)
        self.graphics_view.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.graphics_view.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.zoom_factor = 1.25

        self.setup_ui()
        self.pdf_viewer = PDFViewer(
                            parent=self,
                            scene=self.scene,
                            combo_h_split=self.combo_h_split,
                            combo_v_split=self.combo_v_split,
                            graphics_view=self.graphics_view,
                            zoom_factor=self.zoom_factor,
                        )

        self.label_manager = LabelManager(self.scene, self)

        # 在你的 MainWindow 初始化時
        self.scene.selectionChanged.connect(self.update_label_spinbox_position)


    def setup_ui(self):

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # ✅ 使用模組化工具列
        top_layout = build_top_toolbar(self)
        main_layout.addLayout(top_layout)

        # ✅ 插入分隔線
        separator = build_horizontal_separator()
        main_layout.addWidget(build_horizontal_separator())

        self.splitter = build_splitter(self)
        main_layout.addWidget(self.splitter, 9)

    def closeEvent(self, event):
        if not self.btn_export.isEnabled():
            self.is_closing = True
            QMessageBox.information(self, "提醒", "正在轉換中，已中止處理。")
        else:
            self.is_closing = True
        event.accept()


    def toggle_left_settings(self):
        if self.left_settings_widget.isVisible():
            self.left_settings_widget.hide()
            self.btn_toggle_settings.setText("顯示設定區")
        else:
            self.left_settings_widget.show()
            self.btn_toggle_settings.setText("隱藏設定區")


    def select_pdf_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 PDF 檔案", "", "PDF Files (*.pdf)")
        if path:
            self.combo_h_split.setCurrentIndex(0)
            self.combo_v_split.setCurrentIndex(0)

            self.reset_scene()
            self.pdf_viewer.pdf_path = path  # pdf_viewer 需要知道路徑
            self.btn_select_pdf.setText(f"pdf: {os.path.basename(path)[:14]}")
            self.btn_select_pdf.setMaximumWidth(200)

            success = self.pdf_viewer.load_pdf_preview(path)
            if not success:
                QMessageBox.warning(self, "錯誤", "PDF 預覽載入失敗")
            else:
                self.pdf_viewer.draw_grid_lines()



    def reset_scene(self):
        # 透過 pdf_viewer 和 label_manager 清除場景內容
        self.pdf_viewer.reset_scene()
        self.label_manager.labels.clear()  # 清除標籤清單
    

    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 Excel 檔案", "", "Excel 檔案 (*.xls *.xlsx)")
        if path:
            self.excel_path = path
            filename = os.path.basename(path)
            self.btn_select_exl.setText(f"Excel: {filename[:14]}")  # 限制文字長度
            self.btn_select_exl.setMaximumWidth(200)

    def load_pdf_preview(self):
        if not self.pdf_path:
            print("⚠️ PDF 路徑尚未設定，無法載入預覽")
            return

        images = convert_from_path(self.pdf_path, first_page=1, last_page=1, fmt="png")
        if images:
            image = images[0]
            self.current_image = image
            self.image_width, self.image_height = image.size

            # 顯示圖片
            qt_image = QImage(image.tobytes(), image.width, image.height, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qt_image)
            self.pdf_pixmap_item = self.scene.addPixmap(pixmap)

            # 重設視圖與縮放
            self.graphics_view.setSceneRect(self.scene.itemsBoundingRect())
            self.graphics_view.resetTransform()
            self.graphics_view.fitInView(self.scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)
            self.graphics_view.scale(self.zoom_factor, self.zoom_factor)

            # 若切割選單存在，畫格線
            if hasattr(self, "combo_h_split") and hasattr(self, "combo_v_split"):
                self.draw_grid_lines()
    #####

    def draw_grid_lines(self):
        self.pdf_viewer.draw_grid_lines()
  
    def change_zoom(self):
        zoom_text = self.zoom_combo.currentText().replace("x", "")
        zoom_factor = float(zoom_text)
        self.pdf_viewer.change_zoom(zoom_factor)


    # 新增標籤
    def add_text_label(self):
        label_id = self.combo_label_id.currentText()
        size = int(self.combo_label_size.currentText() or 20)
        self.label_manager.add_text_label(label_id, font_size=size)

    def remove_selected_label(self):
        self.label_manager.remove_selected_label()
   
    def compute_label_offset(self, index, h_count, v_count, image_width, image_height):
        offset_x, offset_y = self.label_manager.compute_label_offset(index, h_count, v_count, image_width, image_height)
        return offset_x, offset_y

    def align_selected_labels_vertically(self):
        selected_items = self.scene.selectedItems()
        if len(selected_items) < 2:
            return
        base_y = selected_items[0].pos().y()
        for item in selected_items:
            item.setPos(item.pos().x(), base_y)

    def align_selected_labels_horizontally(self):
        selected_items = self.scene.selectedItems()
        if len(selected_items) < 2:
            return
        base_x = selected_items[0].pos().x()
        for item in selected_items:
            item.setPos(base_x, item.pos().y())

    def update_label_spinbox_position(self):
        selected_items = self.scene.selectedItems()
        print(f"[DEBUG] 被選取：{len(selected_items)} 個")
        if len(selected_items) == 1:
            pos = selected_items[0].pos()
            print(f"[DEBUG] 座標更新：{pos.x()}, {pos.y()}")
            self.spin_label_x.blockSignals(True)
            self.spin_label_y.blockSignals(True)
            self.spin_label_x.setValue(pos.x())
            self.spin_label_y.setValue(pos.y())
            self.spin_label_x.blockSignals(False)
            self.spin_label_y.blockSignals(False)


    def update_selected_label_position(self):
        selected_items = self.scene.selectedItems()
        if len(selected_items) == 1:
            x = self.spin_label_x.value()
            y = self.spin_label_y.value()
            selected_items[0].setPos(x, y)

    def export_pdf(self):
        if not self.pdf_viewer.pdf_path or not self.label_manager.labels:
            QMessageBox.warning(self, "警告", "⚠️ 沒有載入 PDF 或沒有標籤")
            return
        
        if not self.excel_path:
            QMessageBox.warning(self, "警告", "⚠️ 請選擇 Excel 檔案。")
            return

        data_list,error_msg = self.get_excel_data()

        if  error_msg:
            QMessageBox.warning(self, "警告", f"⚠️ Excel 資料錯誤：{error_msg}")
            return

        label_map = defaultdict(list)
        for label_id, item in self.label_manager.labels:
            label_map[label_id].append(item)

        missing_labels = []
        for label_id in label_map.keys():
            if not any(label_id in data_row and data_row[label_id] for data_row in data_list):
                missing_labels.append(label_id)

        if missing_labels:
            QMessageBox.warning(
                self,
                "資料缺失警告",
                f"資料中找不到標籤欄位：{', '.join(missing_labels)}，請檢查 Excel 資料或標籤設定。"
            )
            return

        self.btn_export.setEnabled(False)
        self.btn_export.setText("載入中...")
        label_param_settings = self.collect_label_param_settings()
        try:
            exporter = PDFExporter(
                pdf_path=self.pdf_viewer.pdf_path,
                labels=self.label_manager.labels,
                image_width=self.pdf_viewer.image_width,
                image_height=self.pdf_viewer.image_height,
                h_count=int(self.combo_h_split.currentText()),
                v_count=int(self.combo_v_split.currentText()),
                font_path=font_path,
                data=data_list,
                compute_offset_func=self.compute_label_offset,
                label_param_settings=label_param_settings,
            )

            output_path, _ = QFileDialog.getSaveFileName(self, "儲存 PDF", "output.pdf", "PDF Files (*.pdf)")
            if output_path:
                exporter.export(output_path)

        finally:
            self.btn_export.setEnabled(True)
            self.btn_export.setText("執行轉換")
 
    def collect_label_param_settings(self) -> dict:
        """
        從 UI 中的標籤參數設定區塊，收集各標籤的自訂參數。
        回傳格式: {'A': {'font_size': 20, 'direction': '水平', 'wrap_limit': 10}, ...}
        """
        label_param_settings = {}

        for i in range(self.label_param_layout.count()):
            gb = self.label_param_layout.itemAt(i).widget()
            label_id = gb.combo_id.currentText()
            label_param_settings[label_id] = {
                "font_size": int(gb.combo_font.currentText()),
                "direction": gb.combo_dir.currentText(),
                "wrap_limit": int(gb.word_limit.currentText())
            }

        return label_param_settings
    
    def get_excel_data(self):
        try:
            # 自訂函式讀取 Excel，並將 NaN 用空字串代替
            df = read_data_auto(self.excel_path)
            df = df.fillna('')

            # 取得使用者設定
            process_mode = self.combo_process_mode.currentText()
            is_fixed = self.radio_mode_fixed.isChecked()

            # 決定資料範圍
            if is_fixed:
                limit_text = self.combo_row_limit.currentText()
                if limit_text == "全部":
                    # 「全部」但跳過第 0 列（標題）
                    df_filtered = df.iloc[:]
                else:
                    limit = int(limit_text)
                    df_filtered = df.iloc[:limit]
            else:
                start = self.spin_row_start.value() # index 從 0 開始
                end = self.spin_row_end.value()
                # ✅ 防呆檢查：start 至少從 1 開始
                if start <= 0 or end < start or end > len(df):
                    return [], f"自訂範圍不合法：從 {start} 到 {end}，但資料總筆數為 {len(df)}"


                df_filtered = df.iloc[start-1:end]

            # 建立欄位字母（A,B,C...）對應實際欄位名稱的字典
            # 這樣能確保對應準確且不受欄位順序影響
            col_letter_map = {chr(65+i): name for i, name in enumerate(df.columns)}
            result_list = []
            for _, row in df_filtered.iterrows():
                result = {k: str(row[v]) for k, v in col_letter_map.items()}
                result_list.append(result)
            
            return result_list, ""


                
            # return [
            # {"A": "地藏王", "B": "劉德華測試換行用"},
            # {"A": "觀音佛", "B": "張學友"},
            # {"A": "普賢菩薩", "B": "郭富城"},
            # {"A": "文殊菩薩", "B": "黎明"},
            # {"A": "釋迦如來", "B": "周星馳"},
            # {"A": "藥師佛", "B": "吳宗憲"},
            # {"A": "阿彌陀佛", "B": "黃子佼"},
            # ], ""

        except Exception as e:
            return [], f"讀取 Excel 發生錯誤：{e}"
        
        finally:
            print("讀取 Excel 檔案完成")

        
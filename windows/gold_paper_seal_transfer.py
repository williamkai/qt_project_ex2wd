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

from core.pdf_exporter import PDFExporter
# 註冊中文字型
pdfmetrics.registerFont(TTFont("Iansui", "/home/william/桌面/地藏王廟/qt_project_ex2wd/core/Iansui-Regular.ttf"))


class GoldPaperSealTransferWindow(QWidget):
    def __init__(self, title="試算表轉金紙封條"):
        super().__init__()
        self.setWindowTitle(title)
        self.is_closing = False
        self.setWindowTitle("PDF 可視化標籤定位工具")
        self.setMinimumSize(1000, 700)

        # PDF 狀態與圖像資訊
        self.pdf_path = None
        self.current_image = None
        self.image_width = None
        self.image_height = None
        self.pdf_pixmap_item = None
        # 場景與視圖
        self.scene = QGraphicsScene()
        self.graphics_view = QGraphicsView(self.scene)
        self.graphics_view.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.graphics_view.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.zoom_factor = 1.25  # 初始縮放因子


        # 場景上的互動物件
        self.labels = []       # (label_id, QGraphicsTextItem)
        self.grid_lines = []   # QGraphicsLineItem
        self.grid_line_width = 3  # 可搭配 UI QSpinBox 調整
        self.block_numbers = []

        self.setup_ui()


    def closeEvent(self, event):
        if not self.btn_export.isEnabled():
            self.is_closing = True
            QMessageBox.information(self, "提醒", "正在轉換中，已中止處理。")
        else:
            self.is_closing = True
        event.accept()


    def setup_ui(self):

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # 🟩 上方工具列使用 QHBoxLayout + 四個 QWidget 區塊
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)

        # ==== 區塊 1：PDF 選擇 ====
        pdf_widget = QWidget()
        pdf_layout = QHBoxLayout(pdf_widget)
        pdf_layout.setContentsMargins(0, 0, 0, 0)
        self.btn_select_pdf = QPushButton("選擇 PDF")
        self.btn_select_pdf.setFixedWidth(100)
        self.btn_select_pdf.clicked.connect(self.select_pdf_file)
        pdf_layout.addWidget(self.btn_select_pdf)
        pdf_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top_layout.addWidget(pdf_widget)

        # ==== 區塊 1-1：exl 選擇 ====
        exl_widget = QWidget()
        exl_layout = QHBoxLayout(exl_widget)
        exl_layout.setContentsMargins(0, 0, 0, 0)
        self.btn_select_exl = QPushButton("選擇 Excel")
        self.btn_select_exl.setFixedWidth(100)
        self.btn_select_exl.clicked.connect(self.select_excel_file)
        exl_layout.addWidget(self.btn_select_exl)
        exl_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top_layout.addWidget(exl_widget)


        # ==== 區塊 2：標籤管理 ====
        label_widget = QWidget()
        label_layout = QHBoxLayout(label_widget)
        label_layout.setContentsMargins(0, 0, 0, 0)
        label_layout.setSpacing(10)
        label_layout.addWidget(QLabel("標籤名稱："))

        self.combo_label_id = QComboBox()
        self.combo_label_id.addItems([chr(i) for i in range(ord('A'), ord('Z') + 1)])
        self.combo_label_id.setFixedWidth(50)
        label_layout.addWidget(self.combo_label_id)

        self.combo_label_size = QComboBox()
        self.combo_label_size .addItems([str(i) for i in range(16, 81, 2)])
        self.combo_label_size.setCurrentIndex(4)
        self.combo_label_size.setFixedWidth(50)
        label_layout.addWidget(self.combo_label_size)

        self.btn_add_label = QPushButton("加入標籤")
        self.btn_add_label.setFixedWidth(90)
        self.btn_add_label.clicked.connect(self.add_text_label)
        label_layout.addWidget(self.btn_add_label)

        self.btn_remove_label = QPushButton("移除標籤")
        self.btn_remove_label.setFixedWidth(90)
        self.btn_remove_label.clicked.connect(self.remove_selected_label)
        label_layout.addWidget(self.btn_remove_label)
        label_layout.addStretch()  # 這行讓控件靠左，右邊空白撐開

        label_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top_layout.addWidget(label_widget)

        # ==== 區塊 3：縮放與設定切換 ====
        control_widget = QWidget()
        control_layout = QHBoxLayout(control_widget)
        control_layout.setContentsMargins(0, 0, 0, 0)
        control_layout.setSpacing(5)

        control_layout.addWidget(QLabel("縮放比例："))
        self.zoom_combo = QComboBox()
        self.zoom_combo.addItems(["1.0x", "1.25x", "1.5x", "2.0x"])
        self.zoom_combo.setCurrentIndex(1)
        self.zoom_combo.currentIndexChanged.connect(self.change_zoom)
        self.zoom_combo.setFixedWidth(80)
        control_layout.addWidget(self.zoom_combo)

        self.btn_toggle_settings = QPushButton("隱藏設定區")
        self.btn_toggle_settings.setCheckable(True)
        self.btn_toggle_settings.setChecked(True)
        self.btn_toggle_settings.setFixedWidth(120)
        self.btn_toggle_settings.clicked.connect(self.toggle_left_settings)
        control_layout.addWidget(self.btn_toggle_settings)
        control_layout.addStretch() 

        control_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top_layout.addWidget(control_widget)

        # ==== 區塊 4：產生 PDF ====
        export_widget = QWidget()
        export_layout = QHBoxLayout(export_widget)
        export_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_export = QPushButton("執行轉換")
        self.btn_export.setFixedWidth(150)
        self.btn_export.clicked.connect(self.export_pdf)
        export_layout.addWidget(self.btn_export)

        export_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top_layout.addWidget(export_widget)

        # ⬇️ 加入整個 top_layout
        main_layout.addLayout(top_layout, 1)

        # ✅ 插入一條水平分隔線（用 QFrame）
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)       # 設定為水平線
        separator.setFrameShadow(QFrame.Shadow.Sunken)    # 顯示陰影
        separator.setStyleSheet("color: #ccc;; height: 2px;")
        main_layout.addWidget(separator)
        # === 下方畫面：使用 QSplitter 左右分割 ===
        self.splitter = QSplitter(Qt.Orientation.Horizontal)  # 水平分割，左設定區 / 右畫面區

        # ➤ 左側設定區（你可以之後加入更多設定控件）
        self.left_settings_widget = QWidget()
        left_layout = QVBoxLayout(self.left_settings_widget)
        left_layout.addWidget(QLabel("設定區域"))  # 示意文字
        left_layout.addStretch()  # 推到底部

        #########
        # 水平 / 垂直 分割設定
        split_title = QLabel("頁面分割設定")
        split_title.setStyleSheet("font-weight: bold;")
        left_layout.addWidget(split_title)

        # 垂直切割
        h_split_layout = QHBoxLayout()
        h_split_layout.addWidget(QLabel("垂直分割："))
        self.combo_h_split = QComboBox()
        self.combo_h_split.addItems([str(i) for i in range(1, 6)])
        self.combo_h_split.setCurrentIndex(0)
        h_split_layout.addWidget(self.combo_h_split)
        left_layout.addLayout(h_split_layout)

        # 水平切割
        v_split_layout = QHBoxLayout()
        v_split_layout.addWidget(QLabel("水平分割："))
        self.combo_v_split = QComboBox()
        self.combo_v_split.addItems([str(i) for i in range(1, 6)])
        self.combo_v_split.setCurrentIndex(0)
        v_split_layout.addWidget(self.combo_v_split)
        left_layout.addLayout(v_split_layout)

        # 畫格線按鈕
        btn_draw_grid = QPushButton("畫出格線")
        btn_draw_grid.clicked.connect(self.draw_grid_lines)
        left_layout.addWidget(btn_draw_grid)
        #########

        # ➤ 右側 PDF 預覽畫面
        self.right_pdf_widget = QWidget()
        right_layout = QVBoxLayout(self.right_pdf_widget)
        right_layout.addWidget(self.graphics_view)

        # ➤ 將兩邊 widget 加進 splitter（左右切分）
        self.splitter.addWidget(self.left_settings_widget)
        self.splitter.addWidget(self.right_pdf_widget)

        self.splitter.setSizes([50, 950])  # 左右區塊初始寬度（px）
        self.splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #aaa;  /* 分隔線顏色 */
                border: none;
                width: 2px;  /* 垂直分隔線寬度，改 height: 2px 可變上下切 */
            }
        """)

        # 將 splitter 加進主畫面（下方）
        main_layout.addWidget(self.splitter, 9)

        # 將主畫面 layout 套用到視窗
        self.setLayout(main_layout)


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

            self.reset_scene()  # ✅ 清除舊場景資料

            self.pdf_path = path
            filename = os.path.basename(path)
            self.btn_select_pdf.setText(f"pdf: {filename[:14]}")
            self.btn_select_pdf.setMaximumWidth(200)

            self.load_pdf_preview()

    def reset_scene(self):
        """清空所有場景物件與狀態變數"""
        # 清除標籤
        for _, label in self.labels:
            self.scene.removeItem(label)
        self.labels.clear()

        # 清除格線
        for line in self.grid_lines:
            self.scene.removeItem(line)
        self.grid_lines.clear()

        # 清除區塊編號
        if hasattr(self, "block_numbers"):
            for num in self.block_numbers:
                self.scene.removeItem(num)
            self.block_numbers.clear()
        else:
            self.block_numbers = []

        # 清除圖片
        if self.pdf_pixmap_item:
            self.scene.removeItem(self.pdf_pixmap_item)
            self.pdf_pixmap_item = None

        # 清除其他圖片狀態
        self.current_image = None
        self.image_width = None
        self.image_height = None

        # 清空場景
        self.scene.clear()

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
        if not self.pdf_path or not self.pdf_pixmap_item:
            QMessageBox.warning(self, "錯誤", "請先載入 PDF 檔案才能畫格線")
            return

        try:
            h_count = int(self.combo_h_split.currentText())
            v_count = int(self.combo_v_split.currentText())
        except Exception:
            QMessageBox.warning(self, "錯誤", "分割數量無效")
            return

        # 若沒設定，使用預設線寬
        line_width = getattr(self, "grid_line_width", 1)

        # 清除舊格線
        for line in self.grid_lines:
            self.scene.removeItem(line)
        self.grid_lines.clear()

        # 清除舊的區塊編號
        if hasattr(self, "block_numbers"):
            for num in self.block_numbers:
                self.scene.removeItem(num)
            self.block_numbers.clear()
        else:
            self.block_numbers = []

        width = self.image_width
        height = self.image_height

        # 垂直線（切水平區塊）
        for i in range(1, h_count):
            x = width / h_count * i
            line = self.scene.addLine(x, 0, x, height, QPen(Qt.GlobalColor.blue, line_width, Qt.PenStyle.DashLine))
            self.grid_lines.append(line)

        # 水平線（切垂直區塊）
        for j in range(1, v_count):
            y = height / v_count * j
            line = self.scene.addLine(0, y, width, y, QPen(Qt.GlobalColor.blue, line_width, Qt.PenStyle.DashLine))
            self.grid_lines.append(line)

        # 顯示順序編號（右上角為 1，橫向從右到左，縱向由上到下）
        index = 1
        for row in range(v_count):  # 垂直（上到下）
            for col in reversed(range(h_count)):  # 水平（右到左）
                block_x = width / h_count * col + width / h_count / 2
                block_y = height / v_count * row + height / v_count / 2
                text_item = QGraphicsSimpleTextItem(str(index))
                text_item.setBrush(QBrush(Qt.GlobalColor.darkGreen))
                text_item.setFont(QFont("Arial", 30))
                text_item.setPos(block_x - 10, block_y - 12)  # 微調位置
                self.scene.addItem(text_item)
                self.block_numbers.append(text_item)
                index += 1

        # 重設視圖大小與縮放
        self.graphics_view.setSceneRect(self.scene.itemsBoundingRect())
        self.graphics_view.fitInView(self.scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)
        self.graphics_view.scale(self.zoom_factor, self.zoom_factor)

    def change_zoom(self):
        if not self.pdf_pixmap_item:
            print("⚠️ 尚未載入 PDF，請先選擇檔案")
            return

        zoom_text = self.zoom_combo.currentText().replace("x", "")
        self.zoom_factor = float(zoom_text)

        self.graphics_view.resetTransform()
        self.graphics_view.fitInView(self.scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)
        self.graphics_view.scale(self.zoom_factor, self.zoom_factor)

    # 新增標籤
    def add_text_label(self):
        label_id = self.combo_label_id.currentText()
        preview_text = f"標籤:{label_id}"
        wsize = self.combo_label_size.currentText() or "20"

        label = QGraphicsTextItem(preview_text)
        label.setFont(QFont("Arial", int(wsize)))
        label.setDefaultTextColor(Qt.GlobalColor.red)
        label.setFlag(QGraphicsTextItem.GraphicsItemFlag.ItemIsMovable)
        label.setFlag(QGraphicsTextItem.GraphicsItemFlag.ItemIsSelectable)
        label.setData(0, label_id)
        label.setPos(QPointF(50, 50))

        self.scene.addItem(label)
        self.labels.append((label_id, label))

    # 移除標籤
    def remove_selected_label(self):
        selected_items = self.scene.selectedItems()
        for item in selected_items:
            if isinstance(item, QGraphicsTextItem):
                # 從 list 移除該 QGraphicsTextItem
                self.labels = [(lid, litem) for (lid, litem) in self.labels if litem != item]
                self.scene.removeItem(item)

    def compute_label_offset(self, index, h_count, v_count, image_width, image_height):
        """讓 index=0 的標籤就在原位，其他往左下擴展"""
        block_width = image_width / h_count
        block_height = image_height / v_count

        local_index = index % (h_count * v_count)
        col = local_index % h_count   # 左到右（0～h_count-1）
        row = local_index // h_count  # 上到下

        offset_x = -block_width * col  # 注意：向左偏移
        offset_y = block_height * row  # 向下偏移

        return offset_x, offset_y

    def export_pdf(self):
        if not self.pdf_path or not self.labels:
            QMessageBox.warning(
                self,
                "警告",
                "⚠️ 沒有載入 PDF 或沒有標籤"
            )
            print("⚠️ 沒有載入 PDF 或沒有標籤")
            return
        # 先取得資料

        data_list = self.get_excel_data()  # 或從實際來源讀取

        # label_map 是 (label_id -> list of QGraphicsTextItem)
        label_map = defaultdict(list)
        for label_id, item in self.labels:
            label_map[label_id].append(item)

        # 檢查標籤是否都有對應資料欄位
        missing_labels = []
        for label_id in label_map.keys():
            # 判斷資料中是否至少有一筆有此欄位且非空
            if not any(label_id in data_row and data_row[label_id] for data_row in data_list):
                missing_labels.append(label_id)

        if missing_labels:
            QMessageBox.warning(
                self,
                "資料缺失警告",
                f"資料中找不到標籤欄位：{', '.join(missing_labels)}，請檢查 Excel 資料或標籤設定。"
            )
            return
        
        self.btn_export = self.btn_export.sender()
        self.btn_export.setEnabled(False)
        self.btn_export.setText("載入中...")
        try:
            # 產生 PDFExporter 並執行
            exporter = PDFExporter(
                pdf_path=self.pdf_path,
                labels=self.labels,
                image_width=self.image_width,
                image_height=self.image_height,
                h_count=int(self.combo_h_split.currentText()),
                v_count=int(self.combo_v_split.currentText()),
                font_path="/home/william/桌面/地藏王廟/qt_project_ex2wd/core/Iansui-Regular.ttf",
                data=self.get_excel_data(),  # 或你要測試的 test_data
                compute_offset_func=self.compute_label_offset,
                direction_map={"A": "vertical", "B": "rtl"}  # 這個依你的資料而定
            )

            output_path, _ = QFileDialog.getSaveFileName(self, "儲存 PDF", "output.pdf", "PDF Files (*.pdf)")
            if output_path:
                exporter.export(output_path)

        except ValueError as ve:
            QMessageBox.warning(self, "輸入錯誤", str(ve))
            return
        
        finally:
            # 無論成功失敗，最後都要恢復按鈕狀態
            self.btn_export.setEnabled(True)
            self.btn_export.setText("執行轉換")

    def get_excel_data(self):
        return [
            {"A": "地藏王", "B": "劉德華"},
            {"A": "觀音佛", "B": "張學友"},
            {"A": "普賢菩薩", "B": "郭富城"},
            {"A": "文殊菩薩", "B": "黎明"},
            {"A": "釋迦如來", "B": "周星馳"},
            {"A": "藥師佛", "B": "吳宗憲"},
            {"A": "阿彌陀佛", "B": "黃子佼"},
        ]
    
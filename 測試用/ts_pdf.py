import os
from io import BytesIO
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QLineEdit, QGraphicsTextItem, QComboBox,QSplitter,QGridLayout,QSizePolicy,QFrame
)
from PyQt6.QtCore import Qt, QPointF
from PyQt6.QtGui import QPixmap, QImage, QFont, QTransform, QPainter
from pdf2image import convert_from_path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from pypdf import PdfReader, PdfWriter, PageObject
from PIL import Image
# 註冊中文字型
pdfmetrics.registerFont(TTFont("Iansui", "/home/william/桌面/地藏王廟/qt_project_ex2wd/core/Iansui-Regular.ttf"))

class PDFEditorWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 可視化標籤定位工具")
        self.setMinimumSize(1000, 700)

        self.pdf_path = None
        self.current_image = None
        self.scene = QGraphicsScene()
        self.graphics_view = QGraphicsView(self.scene)
        self.graphics_view.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.graphics_view.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)

        self.zoom_factor = 1.25  # 初始縮放因子
        self.image_width = None
        self.image_height = None
        self.pdf_pixmap_item = None
        # self.labels = {}  # 儲存 label_id: QGraphicsTextItem
        self.labels = [] 
        self.setup_ui()

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

        self.btn_export = QPushButton("執行並產生 PDF")
        self.btn_export.clicked.connect(self.export_pdf)
        self.btn_export.setFixedWidth(150)
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
            self.pdf_path = path
            filename = os.path.basename(path)
            self.btn_select_pdf.setText(f"pdf: {filename[:14]}")  # 限制文字長度
            self.btn_select_pdf.setMaximumWidth(200)
            self.load_pdf_preview()

    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 Excel 檔案", "", "Excel 檔案 (*.xls *.xlsx)")
        if path:
            self.excel_path = path
            filename = os.path.basename(path)
            self.btn_select_exl.setText(f"Excel: {filename[:14]}")  # 限制文字長度
            self.btn_select_exl.setMaximumWidth(200)

    def load_pdf_preview(self):
        images = convert_from_path(self.pdf_path, first_page=1, last_page=1, fmt="png")
        if images:
            image = images[0]
            self.current_image = image
            self.image_width, self.image_height = image.size
            qt_image = QImage(image.tobytes(), image.width, image.height, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qt_image)
            self.scene.clear()
            self.labels.clear()
            self.pdf_pixmap_item = self.scene.addPixmap(pixmap)
            self.graphics_view.setSceneRect(self.scene.itemsBoundingRect())
            self.graphics_view.resetTransform()
            self.graphics_view.fitInView(self.scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)
            self.graphics_view.scale(self.zoom_factor, self.zoom_factor)

    def change_zoom(self):
        zoom_text = self.zoom_combo.currentText().replace("x", "")
        self.zoom_factor = float(zoom_text)
        self.load_pdf_preview()

    # 新增標籤
    def add_text_label(self):
        label_id = self.combo_label_id.currentText()
        preview_text =  f"標籤:{label_id}" #self.input_text.text() or
        wsize = self.combo_label_size.currentText() or "20"  # 預設保底
        label = QGraphicsTextItem(preview_text)
        label.setFont(QFont("Arial", int(wsize)))
        label.setFlag(QGraphicsTextItem.GraphicsItemFlag.ItemIsMovable)
        label.setFlag(QGraphicsTextItem.GraphicsItemFlag.ItemIsSelectable)
        label.setDefaultTextColor(Qt.GlobalColor.red)
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
    def export_pdf(self):
        if not self.pdf_path or not self.labels:
            return

        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=A4)
        page_width, page_height = A4

        for label_id, item in self.labels:
            pos = item.pos()
            x_ratio = page_width / self.image_width
            y_ratio = page_height / self.image_height
            x_mm = pos.x() * x_ratio
            y_mm = page_height - (pos.y() * y_ratio)

            c.setFont("Iansui", 14)
            c.drawString(x_mm, y_mm, f"{{{label_id}}}")

        c.save()
        packet.seek(0)
        overlay_pdf = PdfReader(packet)

        template = PdfReader(self.pdf_path)
        base_page = template.pages[0]
        new_page = PageObject.create_blank_page(
            width=base_page.mediabox.width,
            height=base_page.mediabox.height,
        )
        new_page.merge_page(base_page)
        new_page.merge_page(overlay_pdf.pages[0])

        writer = PdfWriter()
        writer.add_page(new_page)

        output_path, _ = QFileDialog.getSaveFileName(self, "儲存 PDF", "output.pdf", "PDF Files (*.pdf)")
        if output_path:
            with open(output_path, "wb") as f:
                writer.write(f)
            print(f"✅ 成功寫入 PDF：{output_path}")

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    window = PDFEditorWindow()
    window.show()
    sys.exit(app.exec())

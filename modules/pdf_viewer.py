from PyQt6.QtWidgets import QGraphicsScene, QMessageBox,QGraphicsSimpleTextItem
from PyQt6.QtGui import QPen, QBrush, QFont, QImage,QPixmap
from PyQt6.QtCore import Qt
from pdf2image import convert_from_path

class PDFViewer:
    def __init__(self,parent, scene, combo_h_split, combo_v_split, graphics_view, zoom_factor=1.25):
        self.parent = parent
        self.scene = scene
        self.combo_h_split = combo_h_split
        self.combo_v_split = combo_v_split
        self.graphics_view = graphics_view
        self.zoom_factor = zoom_factor
        self.pdf_path = None

        self.grid_lines = []
        self.block_numbers = []
        self.pdf_pixmap_item = None
        self.current_image = None
        self.image_width = None
        self.image_height = None

    def reset_scene(self):
        for line in self.grid_lines:
            self.scene.removeItem(line)
        self.grid_lines.clear()

        for num in self.block_numbers:
            self.scene.removeItem(num)
        self.block_numbers.clear()

        if self.pdf_pixmap_item:
            self.scene.removeItem(self.pdf_pixmap_item)
            self.pdf_pixmap_item = None

        self.current_image = None
        self.image_width = None
        self.image_height = None
        self.scene.clear()

    def load_pdf_preview(self, pdf_path):
        if not pdf_path:
            print("⚠️ PDF 路徑尚未設定，無法載入預覽")
            return False

        images = convert_from_path(pdf_path, first_page=1, last_page=1, fmt="png")
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
            return True
        return False

    def draw_grid_lines(self):
        if not self.pdf_path or not self.pdf_pixmap_item:
            QMessageBox.warning(self.parent, "錯誤", "請先載入 PDF 檔案才能畫格線")
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

    def change_zoom(self, zoom_factor):
        if not self.pdf_pixmap_item:
            print("⚠️ 尚未載入 PDF，請先選擇檔案")
            return

        self.zoom_factor = zoom_factor
        self.graphics_view.resetTransform()
        self.graphics_view.fitInView(self.scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)
        self.graphics_view.scale(self.zoom_factor, self.zoom_factor)
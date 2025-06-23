from PyQt6.QtWidgets import QGraphicsTextItem
from PyQt6.QtGui import QFont
from PyQt6.QtCore import QPointF, Qt
from .movable_labelItem import MovableLabelItem
class LabelManager:
    def __init__(self, scene, main_window):
        self.scene = scene
        self.labels = []
        self.main_window = main_window

    def add_text_label(self, label_id, font_size=20):
        preview_text = f"標籤:{label_id}"
        # label = QGraphicsTextItem(preview_text)
        label = MovableLabelItem(preview_text)
        label.set_main_window(self.main_window)  # 傳入主視窗
        label.setFont(QFont("Arial", font_size))
        label.setDefaultTextColor(Qt.GlobalColor.red)
        # label.setFlag(QGraphicsTextItem.GraphicsItemFlag.ItemIsMovable)
        # label.setFlag(QGraphicsTextItem.GraphicsItemFlag.ItemIsSelectable)
        label.setData(0, label_id)
        label.setPos(QPointF(50, 50))
        self.scene.addItem(label)
        self.labels.append((label_id, label))

    def remove_selected_label(self):
        selected_items = self.scene.selectedItems()
        for item in selected_items:
            if isinstance(item, QGraphicsTextItem):
                self.labels = [(lid, litem) for (lid, litem) in self.labels if litem != item]
                self.scene.removeItem(item)

    def compute_label_offset(self, index, h_count, v_count, image_width, image_height):
        block_width = image_width / h_count
        block_height = image_height / v_count

        local_index = index % (h_count * v_count)
        col = local_index % h_count
        row = local_index // h_count

        offset_x = -block_width * col
        offset_y = block_height * row

        return offset_x, offset_y

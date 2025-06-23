from PyQt6.QtWidgets import QGraphicsTextItem
from PyQt6.QtCore import QPointF, QEvent, Qt
from PyQt6.QtGui import QFont

class MovableLabelItem(QGraphicsTextItem):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFlags(
            QGraphicsTextItem.GraphicsItemFlag.ItemIsMovable |
            QGraphicsTextItem.GraphicsItemFlag.ItemIsSelectable |
            QGraphicsTextItem.GraphicsItemFlag.ItemSendsGeometryChanges  # ✅ 必加
        )
        self.setFont(QFont("Arial", 20))
        self.setDefaultTextColor(Qt.GlobalColor.red)
        self._main_window = None

    def itemChange(self, change, value):
        if change == QGraphicsTextItem.GraphicsItemChange.ItemPositionHasChanged:
            if self.isSelected() and self._main_window:
                print("[DEBUG] 位置變動觸發！")
                self._main_window.update_label_spinbox_position()
        return super().itemChange(change, value)

    def set_main_window(self, main_window):
        self._main_window = main_window

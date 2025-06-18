from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QPushButton, QApplication,)
from windows.excel_invitation import ExcelToInvitationWindow
from windows.invitation_transfer import InvitationTransferWindow
from windows.receipt_transfer import ReceiptTransferWindow
from windows.other_transfer import OtherTransferWindow

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("地藏王廟小工具")
        self.setMinimumSize(400, 300)
        self.layout = QVBoxLayout(self)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.func_windows = {}

        self.create_buttons()

    def create_buttons(self):
        buttons = [
            ("試算表轉召請文", self.open_excel_to_invitation_window),
            ("試算表轉拔薦文疏-本", self.open_invitation_transfer_window),
            ("試算表轉收據", self.open_receipt_transfer_window),
            ("其他轉換功能", self.open_other_transfer_window),
        ]

        for text, slot in buttons:
            btn = QPushButton(text)
            btn.setFixedSize(200, 40)
            btn.clicked.connect(slot)
            self.layout.addWidget(btn)

    def open_func_window(self, window_class, title):
        if title in self.func_windows:
            win = self.func_windows[title]
            if win.isVisible():
                win.raise_()
                win.activateWindow()
                return
        win = window_class(title)
        win.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        win.destroyed.connect(lambda: self.func_windows.pop(title, None))
        self.func_windows[title] = win
        win.show()

    def open_excel_to_invitation_window(self):
        self.open_func_window(ExcelToInvitationWindow, "試算表轉召請文")

    def open_invitation_transfer_window(self):
        self.open_func_window(InvitationTransferWindow, "試算表轉拔薦文疏-本")

    def open_receipt_transfer_window(self):
        self.open_func_window(ReceiptTransferWindow, "試算表轉收據")

    def open_other_transfer_window(self):
        self.open_func_window(OtherTransferWindow, "其他轉換功能")

    def closeEvent(self, event):
        for win in list(self.func_windows.values()):
            win.close()
        event.accept()


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


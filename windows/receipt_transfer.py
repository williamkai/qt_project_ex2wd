import os,re
import pandas as pd
from docx import Document
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFontMetrics
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QApplication, QLabel, QFileDialog,
    QSizePolicy, QLineEdit, QHBoxLayout, QFormLayout, QMessageBox,QProgressBar
)
from docx.shared import Pt
from docx.oxml.ns import qn
from core.base_windows import BaseFuncWindow



class ReceiptTransferWindow(BaseFuncWindow):
    def __init__(self, title="收據轉換功能"):
        super().__init__(title)

    def setup_ui(self):
        super().setup_ui()
        label_func = QLabel("收據轉換功能（還沒寫）")
        self.layout().addWidget(label_func)

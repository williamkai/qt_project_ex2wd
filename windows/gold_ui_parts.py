# windows/gold_ui_parts.py

from PyQt6.QtWidgets import (
    QWidget, QLabel, QPushButton, QComboBox, QHBoxLayout, QVBoxLayout,
    QFileDialog, QSizePolicy, QFrame, QSplitter, QLineEdit, QMessageBox
)
from PyQt6.QtCore import Qt


def build_top_toolbar(window):
    """
    建立視窗上方的工具列區域，包含：
    - PDF / Excel 選擇
    - 標籤設定與加入
    - 縮放設定
    - 匯出按鈕
    """
    top_layout = QHBoxLayout()
    top_layout.setSpacing(10)

    # ==== PDF 選擇 ====
    pdf_widget = QWidget()
    pdf_layout = QHBoxLayout(pdf_widget)
    pdf_layout.setContentsMargins(0, 0, 0, 0)
    window.btn_select_pdf = QPushButton("選擇 PDF")
    window.btn_select_pdf.setFixedWidth(100)
    window.btn_select_pdf.clicked.connect(window.select_pdf_file)
    pdf_layout.addWidget(window.btn_select_pdf)
    pdf_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    top_layout.addWidget(pdf_widget)

    # ==== Excel 選擇 ====
    exl_widget = QWidget()
    exl_layout = QHBoxLayout(exl_widget)
    exl_layout.setContentsMargins(0, 0, 0, 0)
    window.btn_select_exl = QPushButton("選擇 Excel")
    window.btn_select_exl.setFixedWidth(100)
    window.btn_select_exl.clicked.connect(window.select_excel_file)
    exl_layout.addWidget(window.btn_select_exl)
    exl_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    top_layout.addWidget(exl_widget)

    # ==== 標籤管理 ====
    label_widget = QWidget()
    label_layout = QHBoxLayout(label_widget)
    label_layout.setContentsMargins(0, 0, 0, 0)
    label_layout.setSpacing(10)

    label_layout.addWidget(QLabel("標籤名稱："))

    window.combo_label_id = QComboBox()
    window.combo_label_id.addItems([chr(i) for i in range(ord('A'), ord('Z') + 1)])
    window.combo_label_id.setFixedWidth(50)
    label_layout.addWidget(window.combo_label_id)

    window.combo_label_size = QComboBox()
    window.combo_label_size.addItems([str(i) for i in range(16, 81, 2)])
    window.combo_label_size.setCurrentIndex(4)
    window.combo_label_size.setFixedWidth(50)
    label_layout.addWidget(window.combo_label_size)

    window.btn_add_label = QPushButton("加入標籤")
    window.btn_add_label.setFixedWidth(90)
    window.btn_add_label.clicked.connect(window.add_text_label)
    label_layout.addWidget(window.btn_add_label)

    window.btn_remove_label = QPushButton("移除標籤")
    window.btn_remove_label.setFixedWidth(90)
    window.btn_remove_label.clicked.connect(window.remove_selected_label)
    label_layout.addWidget(window.btn_remove_label)

    label_layout.addStretch()
    label_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    top_layout.addWidget(label_widget)

    # ==== 縮放與設定切換 ====
    control_widget = QWidget()
    control_layout = QHBoxLayout(control_widget)
    control_layout.setContentsMargins(0, 0, 0, 0)
    control_layout.setSpacing(5)

    control_layout.addWidget(QLabel("縮放比例："))
    window.zoom_combo = QComboBox()
    window.zoom_combo.addItems(["1.0x", "1.25x", "1.5x", "2.0x"])
    window.zoom_combo.setCurrentIndex(1)
    window.zoom_combo.currentIndexChanged.connect(window.change_zoom)
    window.zoom_combo.setFixedWidth(80)
    control_layout.addWidget(window.zoom_combo)

    window.btn_toggle_settings = QPushButton("隱藏設定區")
    window.btn_toggle_settings.setCheckable(True)
    window.btn_toggle_settings.setChecked(True)
    window.btn_toggle_settings.setFixedWidth(120)
    window.btn_toggle_settings.clicked.connect(window.toggle_left_settings)
    control_layout.addWidget(window.btn_toggle_settings)
    control_layout.addStretch()
    control_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    top_layout.addWidget(control_widget)

    # ==== 產生 PDF 按鈕 ====
    export_widget = QWidget()
    export_layout = QHBoxLayout(export_widget)
    export_layout.setContentsMargins(0, 0, 0, 0)
    window.btn_export = QPushButton("執行轉換")
    window.btn_export.setFixedWidth(150)
    window.btn_export.clicked.connect(window.export_pdf)
    export_layout.addWidget(window.btn_export)
    export_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    top_layout.addWidget(export_widget)

    return top_layout


def build_horizontal_separator():
    """回傳一條水平分隔線（QFrame）"""
    separator = QFrame()
    separator.setFrameShape(QFrame.Shape.HLine)
    separator.setFrameShadow(QFrame.Shadow.Sunken)
    separator.setStyleSheet("color: #ccc; height: 2px;")
    return separator


def build_left_settings(window):
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.addWidget(QLabel("設定區域"))

    title = QLabel("頁面分割設定")
    title.setStyleSheet("font-weight: bold;")
    layout.addWidget(title)

    # 垂直切割
    h_split_layout = QHBoxLayout()
    h_split_layout.addWidget(QLabel("垂直分割："))
    window.combo_h_split = QComboBox()
    window.combo_h_split.addItems([str(i) for i in range(1, 6)])
    window.combo_h_split.setCurrentIndex(0)
    h_split_layout.addWidget(window.combo_h_split)
    layout.addLayout(h_split_layout)

    # 水平切割
    v_split_layout = QHBoxLayout()
    v_split_layout.addWidget(QLabel("水平分割："))
    window.combo_v_split = QComboBox()
    window.combo_v_split.addItems([str(i) for i in range(1, 6)])
    window.combo_v_split.setCurrentIndex(0)
    v_split_layout.addWidget(window.combo_v_split)
    layout.addLayout(v_split_layout)

    # 畫格線按鈕
    btn_draw = QPushButton("畫出格線")
    btn_draw.clicked.connect(window.draw_grid_lines)
    layout.addWidget(btn_draw)

    layout.addStretch()
    return widget


def build_pdf_preview(window):
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.addWidget(window.graphics_view)
    return widget


def build_splitter(window):
    splitter = QSplitter(Qt.Orientation.Horizontal)
    window.left_settings_widget = build_left_settings(window)
    window.right_pdf_widget = build_pdf_preview(window)

    splitter.addWidget(window.left_settings_widget)
    splitter.addWidget(window.right_pdf_widget)

    splitter.setSizes([50, 950])
    splitter.setStyleSheet("""
        QSplitter::handle {
            background-color: #aaa;
            border: none;
            width: 2px;
        }
    """)
    return splitter
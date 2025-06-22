# windows/gold_ui_parts.py

from PyQt6.QtWidgets import (
    QWidget, QLabel, QPushButton, QComboBox, QHBoxLayout, QVBoxLayout,
    QFileDialog, QSizePolicy, QFrame, QSplitter, QLineEdit, QMessageBox,QScrollArea
,QSpinBox,QGroupBox)
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
    # 外層 Widget，回傳的就是這個
    outer_widget = QWidget()
    outer_layout = QVBoxLayout(outer_widget)

    # 加入標題
    outer_layout.addWidget(QLabel("設定區域"))

    # 建立可捲動內容區域
    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True)
    scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

    # 實際可捲動的內容放在這
    scroll_content = QWidget()
    scroll_layout = QVBoxLayout(scroll_content)
    scroll_layout.setSpacing(15)

    # === 加入分割設定區塊 ===
    scroll_layout.addWidget(build_split_settings(window))

    # ✅ 加入標籤參數設定區塊
    scroll_layout.addWidget(build_label_param_settings(window))

    # 🧩 未來可以加：標籤參數區塊、資料參數設定、保存設定等等
    # scroll_layout.addWidget(build_label_param_settings(window))
    # scroll_layout.addWidget(build_data_param_settings(window))

    scroll_layout.addStretch()
    scroll_content.setLayout(scroll_layout)
    scroll_area.setWidget(scroll_content)

    outer_layout.addWidget(scroll_area)
    return outer_widget

def build_split_settings(window):
    widget = QWidget()
    layout = QVBoxLayout(widget)

    title = QLabel("📄 頁面分割設定")
    title.setStyleSheet("font-weight: bold;")
    layout.addWidget(title)

    # 垂直切割
    h_split_layout = QHBoxLayout()
    label = QLabel("垂直分割：")
    label.setFixedWidth(80) 
    h_split_layout.addWidget(label)
    window.combo_h_split = QComboBox()
    window.combo_h_split.setFixedWidth(50)
    window.combo_h_split.addItems([str(i) for i in range(1, 6)])
    window.combo_h_split.setCurrentIndex(0)
    h_split_layout.addWidget(window.combo_h_split)
    h_split_layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
    
    layout.addLayout(h_split_layout)

    # 水平切割
    v_split_layout = QHBoxLayout()
    label = QLabel("水平分割：")
    label.setFixedWidth(80) 
    v_split_layout.addWidget(label)
    window.combo_v_split = QComboBox()
    window.combo_v_split.setFixedWidth(50)
    window.combo_v_split.addItems([str(i) for i in range(1, 6)])
    window.combo_v_split.setCurrentIndex(0)
    v_split_layout.addWidget(window.combo_v_split)
    v_split_layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
    layout.addLayout(v_split_layout)
    
    # 畫格線按鈕
    btn_draw = QPushButton("畫出格線")
    btn_draw.setFixedWidth(100)
    btn_draw.clicked.connect(window.draw_grid_lines)
    layout.addWidget(btn_draw)

    return widget


def build_pdf_preview(window):
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.addWidget(window.graphics_view)
    return widget


def build_splitter(window):
    splitter = QSplitter(Qt.Orientation.Horizontal)
    window.left_settings_widget = build_left_settings(window)
    window.left_settings_widget.setMinimumWidth(200) 
    window.right_pdf_widget = build_pdf_preview(window)

    splitter.addWidget(window.left_settings_widget)
    splitter.addWidget(window.right_pdf_widget)

    splitter.setSizes([260, 740])
    splitter.setStyleSheet("""
        QSplitter::handle {
            background-color: #aaa;
            border: none;
            width: 2px;
        }
    """)
    return splitter


def build_label_param_settings(window):
    container = QWidget()
    container_layout = QVBoxLayout(container)
    container_layout.setSpacing(10)

    # 🔖 第一行：標題 + ➕按鈕
    row1_widget = QWidget()
    label_param_settings_row1 = QHBoxLayout(row1_widget)  # 改成掛在 QWidget 上
    title = QLabel("🔖標籤參數設定")
    title.setStyleSheet("font-weight: bold;")
    label_param_settings_row1.addWidget(title)

    # ➕ 新增按鈕
    btn_add = QPushButton("➕")
    btn_add.setFixedWidth(50)
    btn_add.clicked.connect(lambda: add_label_param_row(window))
    label_param_settings_row1.addWidget(btn_add)
    label_param_settings_row1.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

    # ✅ 現在可以加這個 row1_widget 了
    container_layout.addWidget(row1_widget)

    # 動態新增參數列的容器
    window.label_param_container = QWidget()
    window.label_param_layout = QVBoxLayout(window.label_param_container)
    window.label_param_layout.setSpacing(5)
    container_layout.addWidget(window.label_param_container)



    return container


def add_label_param_row(window):
    # 用 QGroupBox 包起來（有標題的邊框容器）
    groupbox = QGroupBox()
    groupbox.setStyleSheet("QGroupBox { border: 1px solid gray; margin-top: 5px; padding: 5px; }")
    groupbox.setFixedWidth(320)
    layout = QVBoxLayout(groupbox)
    layout.setContentsMargins(10, 10, 10, 10)
    layout.setSpacing(8)

    # 第一行：標籤 ID、字體大小
    row1 = QHBoxLayout()
    combo_id = QComboBox()
    combo_id.addItems([chr(i) for i in range(ord('A'), ord('Z') + 1)])
    combo_id.setFixedWidth(50)
    row1.addWidget(QLabel("標籤："))
    row1.addWidget(combo_id)

    combo_font = QComboBox()
    combo_font.addItems([str(i) for i in range(16, 81, 2)])
    combo_font.setCurrentText("20")
    combo_font.setFixedWidth(60)
    row1.addWidget(QLabel("字體大小："))
    row1.addWidget(combo_font)
    row1.addStretch()
    layout.addLayout(row1)

    # 第二行：方向、換行限制
    row2 = QHBoxLayout()
    combo_dir = QComboBox()
    combo_dir.addItems(["水平", "垂直"])
    combo_dir.setFixedWidth(60)
    row2.addWidget(QLabel("方向："))
    row2.addWidget(combo_dir)

    word_limit = QComboBox()
    word_limit.addItems([str(i) for i in range(4, 20, 1)])
    word_limit.setCurrentText("10")
    word_limit.setFixedWidth(60)
    row2.addWidget(QLabel("換行限制："))
    row2.addWidget(word_limit)
    row2.addStretch()
    layout.addLayout(row2)

    # 第三行：取消按鈕（放右側）
    row3 = QHBoxLayout()
    row3.addStretch()
    btn_remove = QPushButton("取消")
    btn_remove.setFixedWidth(80)
    row3.addWidget(btn_remove)
    row3.addStretch()
    layout.addLayout(row3)

    # 取消按鈕綁定：移除這個 groupbox
    def on_remove():
        window.label_param_layout.removeWidget(groupbox)
        groupbox.deleteLater()
    btn_remove.clicked.connect(on_remove)

    # 加入到容器
    window.label_param_layout.addWidget(groupbox)
    # 🔽 加在這裡
    groupbox.combo_id = combo_id
    groupbox.combo_font = combo_font
    groupbox.combo_dir = combo_dir
    groupbox.word_limit = word_limit
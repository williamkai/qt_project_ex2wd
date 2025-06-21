import os
from io import BytesIO
from collections import defaultdict
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter, PageObject
from PyQt6.QtCore import QPointF

class PDFExporter:
    def __init__(self, pdf_path, labels, image_width, image_height,
                 h_count, v_count, font_path, data, compute_offset_func,
                 direction_map=None):
        self.pdf_path = pdf_path
        self.labels = labels
        self.image_width = image_width
        self.image_height = image_height
        self.h_count = h_count
        self.v_count = v_count
        self.font_path = font_path
        self.data = data
        self.compute_offset_func = compute_offset_func
        self.direction_map = direction_map or {}

        self.font_name = "Iansui"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont(self.font_name, font_path))

    def draw_text(self, c, text, x_pdf, y_pdf, font_size, direction):
        if direction == "vertical":
            for idx, ch in enumerate(text):
                line_y_pdf = y_pdf - idx * font_size
                c.drawString(x_pdf, line_y_pdf, ch)
        elif direction == "rtl":  # 右到左
            width = pdfmetrics.stringWidth(text, self.font_name, font_size)
            c.drawRightString(x_pdf + width, y_pdf, text)
        else:  # 默認水平左到右
            c.drawString(x_pdf, y_pdf, text)

    def export(self, output_path):
        if not self.pdf_path or not self.labels:
            print("⚠️ 沒有載入 PDF 或沒有標籤")
            return

        blocks_per_page = self.h_count * self.v_count
        label_map = defaultdict(list)
        for label_id, item in self.labels:
            label_map[label_id].append(item)

        template = PdfReader(self.pdf_path)
        base_page = template.pages[0]
        pdf_width = float(base_page.mediabox.width)
        pdf_height = float(base_page.mediabox.height)
        x_ratio = pdf_width / self.image_width
        y_ratio = pdf_height / self.image_height

        writer = PdfWriter()
        total_pages = (len(self.data) + blocks_per_page - 1) // blocks_per_page

        for page_num in range(total_pages):
            packet = BytesIO()
            c = canvas.Canvas(packet, pagesize=(pdf_width, pdf_height))
            print(f"\n🧾 第 {page_num+1} 頁：")

            for i in range(blocks_per_page):
                data_index = page_num * blocks_per_page + i
                if data_index >= len(self.data):
                    break

                data_row = self.data[data_index]
                offset_x, offset_y = self.compute_offset_func(
                    i, self.h_count, self.v_count, self.image_width, self.image_height
                )

                for label_id, item_list in label_map.items():
                    label_text = data_row.get(label_id, "")
                    if not label_text:
                        continue

                    for item in item_list:
                        orig_pos = item.pos()
                        new_pos = orig_pos + QPointF(offset_x, offset_y)
                        text_height = item.boundingRect().height()
                        font_size = item.font().pointSize()

                        x_pdf = new_pos.x() * x_ratio
                        y_pdf = pdf_height - ((new_pos.y() + text_height) * y_ratio)

                        direction = self.direction_map.get(label_id, "vertical")
                        c.setFont(self.font_name, font_size)
                        self.draw_text(c, label_text, x_pdf, y_pdf, font_size, direction)

            print(f"[Debug] 寫入 canvas，label 數量: {len(label_map)}")

            c.save()
            packet.seek(0)
            overlay_pdf = PdfReader(packet)
            if not overlay_pdf.pages:
                print("⚠️ PDF Overlay 沒有頁面，跳過合併")
                continue  # 或是 return，避免崩潰

            new_page = PageObject.create_blank_page(width=pdf_width, height=pdf_height)
            new_page.merge_page(base_page)
            new_page.merge_page(overlay_pdf.pages[0])
            writer.add_page(new_page)

        with open(output_path, "wb") as f:
            writer.write(f)
        print(f"✅ 成功寫入 PDF：{output_path}")

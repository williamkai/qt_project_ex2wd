import os,re
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
                 label_param_settings=None):
        self.pdf_path = pdf_path
        self.labels = labels
        self.image_width = image_width
        self.image_height = image_height
        self.h_count = h_count
        self.v_count = v_count
        self.font_path = font_path
        self.data = data
        self.compute_offset_func = compute_offset_func
        self.label_param_settings = label_param_settings or {}

        self.font_name = "Iansui"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont(self.font_name, font_path))

    def draw_text(self, canvas, text, x, y, font_size, direction, wrap_limit):
        # ✅ 先設定字型
        canvas.setFont(self.font_name, font_size)
        # 🔸 自訂關鍵詞（可放在類別屬性 or 傳入參數）
        keywords = ["金紙", "營業所", "店"]

        # 🔸 斷行處理：先關鍵詞，再字數
        lines = self.split_text_by_keywords(text, wrap_limit, keywords)
        total_lines = len(lines)
        # lines = [text[i:i+wrap_limit] for i in range(0, len(text), wrap_limit)]
        # total_lines = len(lines)

        if direction == "垂直":
            for line_idx, line in enumerate(lines):
                offset_x = -line_idx * font_size
                # offset_x = (total_lines - 1 - line_idx) * font_size
                # offset_x = line_idx * font_size  # ✅ 每一行往右移
                for char_idx, char in enumerate(line):
                    canvas.drawString(
                        x + offset_x, 
                        y - char_idx * font_size,   # 垂直堆疊
                        char
                    )
        elif direction == "水平":
            for line_idx, line in enumerate(lines):
                canvas.drawString(
                    x,
                    y - line_idx * font_size,       # 每行往下
                    line
                )
        else:
            # fallback: 不處理換行
            canvas.drawString(x, y, text)
    def split_text_by_keywords(self,text: str, wrap_limit: int, keywords: list[str]) -> list[str]:
        segments = []

        # 🔸 用正則先處理每一種關鍵詞邏輯
        remaining = text

        # 🔶 金紙：先處理一次性關鍵詞（金紙）
        for keyword in keywords:
            if keyword == "金紙":
                match = re.search(f"(.*?){keyword}(.*)", remaining)
                if match:
                    before = match.group(1)
                    after = keyword + match.group(2)
                    if before.strip():
                        segments.append(before)
                    segments.append(after)
                    remaining = ""
                    break  # ✅ 金紙只處理一次，結束
        if remaining:
            segments.append(remaining)

        # 🔶 再處理可多次出現的（營業所、店）
        new_segments = []
        for seg in segments:
            # 對每個 segment 再拆：處理多次關鍵詞分割（保留關鍵詞）
            parts = [seg]
            for kw in ["營業所", "店"]:
                temp = []
                for part in parts:
                    # 保留關鍵詞在句尾：使用 lookahead
                    sub_parts = re.split(f"(?<={kw})", part)
                    temp.extend([s for s in sub_parts if s])  # 去除空段
                parts = temp
            new_segments.extend(parts)

        # 🔶 wrap_limit 分段
        lines = []
        for seg in new_segments:
            seg = seg.strip()
            for i in range(0, len(seg), wrap_limit):
                lines.append(seg[i:i+wrap_limit])

        return lines

        # for keyword in keywords:
        #     pattern = re.escape(keyword)

        #     if keyword == "金紙":
        #         # 🔸 金紙 -> 前面一段 + 金紙開頭的後段
        #         match = re.search(f"(.*?){pattern}.*", remaining)
        #         if match:
        #             prefix = match.group(1)
        #             suffix = remaining[len(prefix):]
        #             if prefix:
        #                 segments.append(prefix)
        #             segments.append(suffix)
        #             remaining = ""  # 處理完後結束
        #             break

        #     else:
        #         # 🔸 營業所 / 店 -> 在關鍵詞**尾巴**切開
        #         match = re.search(f"(.*?{pattern})(.*)", remaining)
        #         if match:
        #             first = match.group(1)
        #             second = match.group(2)
        #             segments.append(first)
        #             remaining = second

        # if remaining:
        #     segments.append(remaining)

        # # 🔸 第二層斷行：每一段再根據 wrap_limit 拆成小段
        # lines = []
        # for seg in segments:
        #     for i in range(0, len(seg), wrap_limit):
        #         lines.append(seg[i:i+wrap_limit])

        # return lines

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

                        params = self.label_param_settings.get(label_id, {})
                        font_size = params.get("font_size") or 22
                        direction = params.get("direction", "水平")  # 預設垂直
                        wrap_limit = params.get("wrap_limit", 10)
                        self.draw_text(c, label_text, x_pdf, y_pdf, font_size, direction, wrap_limit)

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

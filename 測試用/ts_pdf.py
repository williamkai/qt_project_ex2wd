import copy
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from io import BytesIO
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter, PageObject
pdfmetrics.registerFont(TTFont("Iansui", "/home/william/桌面/地藏王廟/qt_project_ex2wd/core/Iansui-Regular.ttf"))

TEMPLATE_PATH = "/home/william/桌面/地藏王廟/qt_project_ex2wd/測試用/地藏王1.pdf"
OUTPUT_PATH = "/home/william/桌面/地藏王廟/qt_project_ex2wd/測試用/output_overlay.pdf"

def generate_overlay_page(data_for_page):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)

    base_x = 100 * mm
    start_y = 270 * mm
    y_step = 25 * mm
    font_size = 14

    for idx, item in enumerate(data_for_page):
        y = start_y - idx * y_step
        name = item["name"]
        wish = item["wish"]

        c.setFont("Iansui", font_size)
        c.drawString(base_x, y, f"祈福者：{name}")
        c.drawString(base_x, y - 8 * mm, f"祝願：{wish}")

    c.save()
    packet.seek(0)
    return PdfReader(packet)


def generate_filled_pdf(data, template_path, output_path):
    template = PdfReader(template_path)
    writer = PdfWriter()

    data_per_page = 8
    total_pages = (len(data) + data_per_page - 1) // data_per_page

    for i in range(total_pages):
        page_data = data[i * data_per_page : (i + 1) * data_per_page]
        print(f"Page {i+1} data:", page_data)

        overlay_pdf = generate_overlay_page(page_data)
        base_page = template.pages[0]

        # 建立新的空白頁面，大小同模板頁
        new_page = PageObject.create_blank_page(
            width=base_page.mediabox.width,
            height=base_page.mediabox.height,
        )
        # 把模板頁合併進去 new_page
        new_page.merge_page(base_page)
        # 再把文字 overlay 頁合併進去
        new_page.merge_page(overlay_pdf.pages[0])

        writer.add_page(new_page)

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"✅ 已產生 PDF：{output_path}")


# 測試資料
sample_data = [{"name": f"祈福者 {i+1}", "wish": f"願望 {i+1}"} for i in range(20)]

if __name__ == "__main__":
    generate_filled_pdf(sample_data, TEMPLATE_PATH, OUTPUT_PATH)

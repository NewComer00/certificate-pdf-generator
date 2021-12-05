import io
import csv
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.fonts import addMapping
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def write_text(canvas_obj, text, text_box, font_name, font_size, font_color):
    # text_box: (left_boundary, right_boundary, height_base)
    # font_color: (R, G ,B) between 0~1

    # auto ajust font size when text is too long
    text_len = pdfmetrics.stringWidth(text, font_name, font_size)
    while text_len > text_box[1] - text_box[0]:
        font_size -= 0.5
        text_len = pdfmetrics.stringWidth(text, font_name, font_size)
    canvas_obj.setFont(font_name, size=font_size)
    canvas_obj.setFillColorRGB(*font_color)
    canvas_obj.drawString(0.5 * (text_box[0] + text_box[1] - text_len), text_box[2], text)


# register new font
FONT_PATH = "font/"
FONT_FILE = FONT_PATH + "chinese.msyh.ttf"
pdfmetrics.registerFont(TTFont('MicrosoftYaHei', FONT_FILE))

PROJECT_PATH = "example/"
DATA_FILE = PROJECT_PATH + "data/data.csv"
TEMPLATE_FILE = PROJECT_PATH + "template/template.pdf"
OUTPUT_FILE = PROJECT_PATH + "output/%s.pdf"

with open(DATA_FILE, newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:

        pagesize = (425.95 * mm, 285.97 * mm)
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=pagesize)

        text_box = (0.76 * pagesize[0], 0.95 * pagesize[0], 0.885 * pagesize[1])
        text = row["编号"]
        font_name = 'MicrosoftYaHei'
        font_size = 18
        font_color = (0.25, 0.25, 0.25)
        write_text(can, text, text_box, font_name, font_size, font_color)

        text_box = (0.37 * pagesize[0], 0.74 * pagesize[0], 0.375 * pagesize[1])
        text = row["作品"]
        font_name = 'MicrosoftYaHei'
        font_size = 28
        font_color = (0.4, 0.4, 0.4)
        write_text(can, text, text_box, font_name, font_size, font_color)

        text_box = (0.37 * pagesize[0], 0.74 * pagesize[0], 0.285 * pagesize[1])
        text = row["队员"]
        font_name = 'MicrosoftYaHei'
        font_size = 28
        font_color = (0.4, 0.4, 0.4)
        write_text(can, text, text_box, font_name, font_size, font_color)

        can.save()

        #move to the beginning of the StringIO buffer
        packet.seek(0)

        # create a new PDF with Reportlab
        new_pdf = PdfFileReader(packet)
        # read your existing PDF
        existing_pdf = PdfFileReader(open(TEMPLATE_FILE, "rb"))
        output = PdfFileWriter()
        # add the "watermark" (which is the new pdf) on the existing page
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)
        # finally, write "output" to a real file
        file_name = OUTPUT_FILE % row["编号"]
        outputStream = open(file_name, "wb")
        output.write(outputStream)
        outputStream.close()

        print("File \"%s\" is created." % file_name)

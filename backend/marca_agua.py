from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
import io

def agregar_marca_agua(pdf_entrada, pdf_salida, imagen_path):

    reader = PdfReader(pdf_entrada)
    writer = PdfWriter()

    for page in reader.pages:

        width = float(page.mediabox.width)
        height = float(page.mediabox.height)

        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))

        # 🔥 TRANSPARENCIA REAL (CLAVE)
        c.saveState()
        c.setFillAlpha(0.5)  # 👈 aquí está la magia

        x = 80
        y = 180

        c.drawImage(
            imagen_path,
            x,
            y,
            width=400,
            height=150,
            mask='auto'
        )

        c.restoreState()
        c.save()

        packet.seek(0)

        watermark = PdfReader(packet)
        marca = watermark.pages[0]

        # 🔥 fondo correcto
        marca.merge_page(page) 
        writer.add_page(marca)

    with open(pdf_salida, "wb") as f:
        writer.write(f)
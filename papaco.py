import openpyxl
import qrcode
import uuid
import tempfile
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from PIL import Image, ImageDraw, ImageFont


def create_qr_code(data):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    img = img.convert("RGBA")
    return img


def generate_pdf(url: str, password: str):
    random_id = uuid.uuid4()

    url_qr_img = create_qr_code(url)
    password_qr_img = create_qr_code(password)

    pdf_filename = f"/Users/gardusig/Desktop/qrcode/{random_id}.pdf"
    c = canvas.Canvas(pdf_filename, pagesize=landscape(A4))
    width, height = landscape(A4)
    max_qr_width = max_qr_height = (height - 100) / 2  # Leave some margins
    text_height = 30
    total_height = max_qr_height + text_height + 20  # 20 for padding

    with tempfile.NamedTemporaryFile(suffix=".png") as url_qr_file, tempfile.NamedTemporaryFile(suffix=".png") as password_qr_file:
        url_qr_img.save(url_qr_file.name)
        password_qr_img.save(password_qr_file.name)

        # Calculate positions
        section_width = width / 2
        url_x = section_width / 2 - max_qr_width / 2
        password_x = section_width + section_width / 2 - max_qr_width / 2
        qr_y = height / 2 - max_qr_height / 2

        # Measure text width to center the titles
        url_text = "URL"
        password_text = "Password"
        url_text_width = c.stringWidth(url_text, "Helvetica", 12)
        password_text_width = c.stringWidth(password_text, "Helvetica", 12)

        # Draw URL text and QR code
        c.drawString(url_x + max_qr_width / 2 - url_text_width /
                     2, qr_y + max_qr_height + text_height, url_text)
        c.drawImage(url_qr_file.name, url_x, qr_y,
                    width=max_qr_width, height=max_qr_height)

        # Draw Password text and QR code
        c.drawString(password_x + max_qr_width / 2 - password_text_width /
                     2, qr_y + max_qr_height + text_height, password_text)
        c.drawImage(password_qr_file.name, password_x, qr_y,
                    width=max_qr_width, height=max_qr_height)

    c.save()
    print(f"PDF saved as {pdf_filename}")


def main():
    workbook = openpyxl.load_workbook('input.xlsx')
    for row in workbook.active.iter_rows(min_row=2, values_only=True):
        generate_pdf(row[4], row[3])


if __name__ == "__main__":
    main()

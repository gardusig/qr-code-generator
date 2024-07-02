import openpyxl
import qrcode
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


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


def draw_border_and_corners(c, width, height):
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(3)
    c.rect(10, 10, width - 20, height - 20)

    c.setLineWidth(3)
    corner_size = 30
    c.line(10, height - 10, 10 + corner_size, height - 10)
    c.line(10, height - 10, 10, height - 10 - corner_size)

    c.line(width - 10, height - 10, width - 10 - corner_size, height - 10)
    c.line(width - 10, height - 10, width - 10, height - 10 - corner_size)

    c.line(10, 10, 10 + corner_size, 10)
    c.line(10, 10, 10, 10 + corner_size)

    c.line(width - 10, 10, width - 10 - corner_size, 10)
    c.line(width - 10, 10, width - 10, 10 + corner_size)


def draw_qr_code_and_text(c, qr_img, x, y, max_qr_width, max_qr_height, text, vertical_align='top'):
    text_height = 40
    font_size = 16
    with tempfile.NamedTemporaryFile(suffix=".png") as qr_file:
        qr_img.save(qr_file.name)
        qr_text_width = c.stringWidth(text, "Helvetica", font_size)

        # Calculate text position to center it
        text_x = x + (max_qr_width - qr_text_width) / 2

        if vertical_align == 'top':
            text_y = y + max_qr_height + 10
            qr_y = y
        else:
            text_y = y - text_height - 10
            qr_y = y - max_qr_height + 90  # Adjusted value for proper centering

        # Draw text and QR code
        c.setFont("Helvetica", font_size)
        c.drawString(text_x, text_y, text)
        c.drawImage(qr_file.name, x, qr_y,
                    width=max_qr_width, height=max_qr_height)


def add_page_to_pdf(c, url, password):
    url_qr_img = create_qr_code(url)
    password_qr_img = create_qr_code(password)

    width, height = A4
    # Increase size for larger QR code and text
    max_qr_width = max_qr_height = (height - 100) / 2

    draw_border_and_corners(c, width, height)

    # Calculate positions
    section_height = height / 2
    qr_x = width / 2 - max_qr_width / 2
    url_y = section_height + (section_height / 2) - (max_qr_height / 2) - 30
    password_y = section_height / 2 + \
        (section_height / 2) - max_qr_height / 2 - \
        224

    draw_qr_code_and_text(c, url_qr_img, qr_x, url_y, max_qr_width,
                          max_qr_height, "URL", vertical_align='top')
    draw_qr_code_and_text(c, password_qr_img, qr_x, password_y,
                          max_qr_width, max_qr_height, "Senha", vertical_align='top')

    # Draw the middle line
    c.setLineWidth(3)
    c.line(10, height / 2, width - 10, height / 2)

    c.showPage()


def main():
    pdf_filename = "/Users/gardusig/Desktop/qrcode/combined.pdf"
    c = canvas.Canvas(pdf_filename, pagesize=A4)
    workbook = openpyxl.load_workbook('input.xlsx')
    for row in workbook.active.iter_rows(min_row=2, values_only=True):
        add_page_to_pdf(c, row[4], row[3])
    c.save()
    print(f"PDF saved as {pdf_filename}")


if __name__ == "__main__":
    main()

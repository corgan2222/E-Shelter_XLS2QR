from reportlab.graphics import renderPDF
from reportlab.pdfgen import canvas
import qrcode
import qrcode.image.svg
from svglib.svglib import svg2rlg
import os

point = 1
inch = 72

def make_qr_code_drawing(data, size, filename):
    qr = qrcode.QRCode(
        version=2,  # QR code version a.k.a size, None == automatic
        error_correction=qrcode.constants.ERROR_CORRECT_H,  # lots of error correction
        box_size=size,  # size of each 'pixel' of the QR code
        border=4  # minimum size according to spec
        )
    qr.add_data(data)
    qrcode_svg = qr.make_image(image_factory=qrcode.image.svg.SvgPathFillImage)
    svg_file = filename
    qrcode_svg.save(svg_file)  # store as an SVG file    
    return svg2rlg(svg_file)

def make_pdf_file(svg_filepath, pdf_filepath, qrdata, floor, room, label, gw, internalID, id):

    qrcode_rl = make_qr_code_drawing(qrdata, 8, svg_filepath)
    c = canvas.Canvas(pdf_filepath, pagesize=(3 * inch, 1.5 * inch))    
    c.setStrokeColorRGB(0,0,0)
    c.setFillColorRGB(0,0,0)
    c.setFont("Helvetica", 7 * point)
    c.drawString( 1.55 * inch, 1.3 * inch, "Floor: " + floor )
    c.drawString( 1.55 * inch, 1.2 * inch, "Room: " + room )
    c.drawString( 1.55 * inch, 1.1 * inch, label )
    c.setFont("Helvetica", 5 * point)
    c.drawString( 1.55 * inch, 0.7 * inch, "Gateway:" )
    c.drawString( 1.55 * inch, 0.6 * inch, gw )
    c.drawString( 1.55 * inch, 0.5 * inch, "ID:" )
    c.drawString( 1.55 * inch, 0.4 * inch, internalID)
    c.setFont("Helvetica", 14 * point)
    c.drawString( 1.55 * inch, 0.1 * inch, id )   
    renderPDF.draw(qrcode_rl, c, 0, 0)
    c.showPage()
    c.save()

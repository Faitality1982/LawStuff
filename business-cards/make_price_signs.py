#!/usr/bin/env python3
"""Generate High Caliber estate sale price signs matching the existing
sign package (navy header/footer bands cropped from the original PDF,
Archivo type, navy/gold/cream palette)."""

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

IN = 72.0
W, H = letter  # 612 x 792

NAVY = (23/255, 41/255, 63/255)
GOLD = (180/255, 146/255, 70/255)
CREAM = (239/255, 235/255, 224/255)
GRAY = (70/255, 70/255, 70/255)

HEADER_H = 425/300 * IN   # 1.4167"
FOOTER_H = 0.5 * IN

pdfmetrics.registerFont(TTFont("ArchivoBlack", "fonts/ArchivoBlack-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Archivo", "fonts/Archivo-Regular.ttf"))
pdfmetrics.registerFont(TTFont("ArchivoMed", "fonts/Archivo-Medium.ttf"))
pdfmetrics.registerFont(TTFont("ArchivoSemi", "fonts/Archivo-SemiBold.ttf"))


def fitted_size(text, font, target_size, max_width):
    w = pdfmetrics.stringWidth(text, font, target_size)
    return target_size if w <= max_width else target_size * max_width / w


def draw_bands(c):
    c.drawImage("header300.png", 0, H - HEADER_H, width=W, height=HEADER_H)
    c.drawImage("footer300.png", 0, 0, width=W, height=FOOTER_H)


def headline(c, lines, y_top, size=80, max_w=6.6*IN, leading=1.08, color=NAVY):
    """Draw centered Archivo Black headline lines; returns y below block."""
    sizes = [fitted_size(t, "ArchivoBlack", size, max_w) for t in lines]
    s = min(sizes)
    y = y_top
    c.setFillColorRGB(*color)
    for t in lines:
        y -= s
        c.setFont("ArchivoBlack", s)
        c.drawCentredString(W/2, y, t)
        y -= s * (leading - 1)
    return y


def divider(c, y, width=1.5*IN, thickness=3):
    c.setStrokeColorRGB(*GOLD)
    c.setLineWidth(thickness)
    c.line(W/2 - width/2, y, W/2 + width/2, y)


def subtext(c, lines, y_top, size=15.5, leading=1.55, color=GRAY):
    c.setFillColorRGB(*color)
    c.setFont("Archivo", size)
    y = y_top
    for t in lines:
        c.drawCentredString(W/2, y, t)
        y -= size * leading
    return y


def price_row(c, y, item, price, note=""):
    """One cream row: item name left, price right. y = row top."""
    row_w, row_h = 5.7*IN, 0.66*IN
    x = W/2 - row_w/2
    c.setFillColorRGB(*CREAM)
    c.roundRect(x, y - row_h, row_w, row_h, 5, stroke=0, fill=1)
    c.setFillColorRGB(*NAVY)
    c.setFont("ArchivoSemi", 20)
    ty = y - row_h/2 - 7
    c.drawString(x + 0.32*IN, ty, item)
    c.setFont("ArchivoBlack", 27)
    c.drawRightString(x + row_w - 0.32*IN, y - row_h/2 - 9.5, price)
    if note:
        c.setFont("ArchivoMed", 11)
        c.setFillColorRGB(*GRAY)
        c.drawRightString(x + row_w - 0.32*IN, y - row_h + 6, note)
    return y - row_h


def statement_sign(c, head_lines, price, sub_lines):
    draw_bands(c)
    body_top = H - HEADER_H
    y = headline(c, head_lines, body_top - 1.35*IN, size=74)
    # giant price
    c.setFillColorRGB(*NAVY)
    c.setFont("ArchivoBlack", 150)
    y -= 150 * 1.12
    c.drawCentredString(W/2, y, price)
    # small gold EACH tag, letterspaced
    c.setFillColorRGB(*GOLD)
    c.setFont("ArchivoSemi", 17)
    word = "E A C H"
    c.drawCentredString(W/2, y - 0.42*IN, word)
    divider(c, y - 0.85*IN)
    subtext(c, sub_lines, y - 1.35*IN)
    c.showPage()


def table_sign(c, head_lines, rows, sub_lines):
    draw_bands(c)
    body_top = H - HEADER_H
    y = headline(c, head_lines, body_top - 0.72*IN, size=54)
    divider(c, y - 0.32*IN, width=1.9*IN)
    y -= 0.85*IN
    for item, price in rows:
        y = price_row(c, y, item, price) - 0.17*IN
    subtext(c, sub_lines, y - 0.35*IN, size=13.5)
    c.showPage()


c = canvas.Canvas("HighCaliber-price-signs.pdf", pagesize=letter)
c.setTitle("High Caliber Estate Sales — Price Signs")

# 1. Garage — hand tools
statement_sign(
    c,
    ["ALL HAND TOOLS"],
    "$2",
    ["Wrenches, sockets, screwdrivers, pliers, hammers — your pick.",
     "Power tools and specialty tools are priced as marked."],
)

# 2. Electronics room — media
table_sign(
    c,
    ["MOVIES & MUSIC"],
    [("VHS tapes", "$1"),
     ("Cassette tapes", "$1"),
     ("CDs", "$2"),
     ("Vinyl records", "$3")],
    ["Prices are per item.",
     "Box sets and collectible records are priced as marked."],
)

# 3. Bedroom closets — clothing
table_sign(
    c,
    ["ALL CLOTHING"],
    [("Shirts & tops", "$2"),
     ("Pants, jeans & skirts", "$3"),
     ("Dresses", "$4"),
     ("Shoes & boots (pair)", "$4"),
     ("Coats & jackets", "$5"),
     ("Hats, belts & scarves", "$1")],
    ["Prices are per item.",
     "Designer and specialty pieces are priced as marked."],
)

# 4. Costume jewelry
statement_sign(
    c,
    ["UNMARKED", "COSTUME JEWELRY"],
    "$2",
    ["Necklaces, bracelets, earrings, pins — any unmarked piece.",
     "Fine jewelry and marked pieces are priced as tagged."],
)

c.save()
print("wrote HighCaliber-price-signs.pdf")

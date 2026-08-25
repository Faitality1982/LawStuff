#!/usr/bin/env python3
"""Impose business-card PDFs onto Avery 8371 sheets.

Avery 8371: US Letter (8.5" x 11"), 10 cards per sheet laid out
2 columns x 5 rows, each card 3.5" x 2", no gutters between cards,
0.75" side margins and 0.5" top/bottom margins.

Each input PDF's first page is placed 10-up on its own output sheet.
Print-shop files with a standard bleed (e.g. OvernightPrints
3.75" x 2.25" = 1/8" per side) are detected and center-cropped to the
3.5" x 2" trim size; anything else is scaled to fit. Portrait-oriented
cards are rotated to landscape.

Usage:
    python avery8371.py card1.pdf [card2.pdf ...] [-o output.pdf]
    python avery8371.py card.pdf --guides   # add light cut guides
"""

import argparse
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import RectangleObject

PT_PER_IN = 72.0

SHEET_W = 8.5 * PT_PER_IN
SHEET_H = 11.0 * PT_PER_IN
CARD_W = 3.5 * PT_PER_IN
CARD_H = 2.0 * PT_PER_IN
COLS, ROWS = 2, 5
MARGIN_X = 0.75 * PT_PER_IN   # (8.5 - 2*3.5) / 2
MARGIN_Y = 0.5 * PT_PER_IN    # (11 - 5*2) / 2

# Standard full-bleed allowances (per side, inches), matched with tolerance.
KNOWN_BLEEDS = (0.125, 0.1, 0.0625)
SIZE_TOL = 3.0  # points


def _prepare_card(page):
    """Crop/scale the card page and return a Transformation mapping it onto
    a 3.5x2 box anchored at the origin.

    pypdf clips merged content to the source page's cropbox, so shrinking
    the cropbox is a true crop — bleed cannot spill onto adjacent cards.
    """
    crop = page.cropbox
    w, h = float(crop.width), float(crop.height)
    rotated = h > w
    cw, ch = (h, w) if rotated else (w, h)  # dimensions as laid out

    detected_bleed = None
    for bleed in KNOWN_BLEEDS:
        b = bleed * PT_PER_IN
        if abs(cw - (CARD_W + 2 * b)) <= SIZE_TOL and abs(ch - (CARD_H + 2 * b)) <= SIZE_TOL:
            detected_bleed = b
            break

    if detected_bleed:
        page.cropbox = RectangleObject((
            float(crop.left) + detected_bleed,
            float(crop.bottom) + detected_bleed,
            float(crop.right) - detected_bleed,
            float(crop.top) - detected_bleed,
        ))
        crop = page.cropbox
        w, h = float(crop.width), float(crop.height)
        cw, ch = (h, w) if rotated else (w, h)

    t = Transformation().translate(-float(crop.left), -float(crop.bottom))
    if rotated:
        # (x, y) -> (-y, x); shift right so content sits in [0, h] x [0, w]
        t = t.rotate(90).translate(h, 0)

    scale = min(CARD_W / cw, CARD_H / ch)
    if abs(scale - 1.0) > 1e-6:
        t = t.scale(scale, scale)
    # center inside the 3.5x2 slot
    t = t.translate((CARD_W - cw * scale) / 2, (CARD_H - ch * scale) / 2)

    kind = ("bleed cropped" if detected_bleed else
            "exact size" if abs(scale - 1.0) <= 1e-6 else f"scaled {scale:.3f}x")
    return t, kind


def make_sheet(writer, card_page):
    sheet = writer.add_blank_page(width=SHEET_W, height=SHEET_H)
    t, kind = _prepare_card(card_page)
    for row in range(ROWS):
        for col in range(COLS):
            x = MARGIN_X + col * CARD_W
            y = SHEET_H - MARGIN_Y - (row + 1) * CARD_H
            sheet.merge_transformed_page(card_page, t.translate(x, y), expand=False)
    return sheet, kind


def add_guides(sheet):
    """Draw faint hairline cut guides in the margins (outside card area)."""
    from io import BytesIO
    from reportlab.pdfgen import canvas

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(SHEET_W, SHEET_H))
    c.setLineWidth(0.25)
    c.setStrokeColorRGB(0.6, 0.6, 0.6)
    tick = 0.25 * PT_PER_IN
    for col in range(COLS + 1):
        x = MARGIN_X + col * CARD_W
        c.line(x, SHEET_H - MARGIN_Y + tick / 2, x, SHEET_H - MARGIN_Y + tick)
        c.line(x, MARGIN_Y - tick, x, MARGIN_Y - tick / 2)
    for row in range(ROWS + 1):
        y = MARGIN_Y + row * CARD_H
        c.line(MARGIN_X - tick, y, MARGIN_X - tick / 2, y)
        c.line(SHEET_W - MARGIN_X + tick / 2, y, SHEET_W - MARGIN_X + tick, y)
    c.save()
    buf.seek(0)
    overlay = PdfReader(buf).pages[0]
    sheet.merge_page(overlay)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Tile business cards onto Avery 8371 sheets (10-up).")
    ap.add_argument("cards", nargs="+", help="Business card PDF(s); first page of each is used")
    ap.add_argument("-o", "--output", default="avery8371-sheets.pdf", help="Output PDF path")
    ap.add_argument("--guides", action="store_true", help="Add faint cut guides in the margins")
    args = ap.parse_args(argv)

    writer = PdfWriter()
    for path in args.cards:
        p = Path(path)
        if not p.exists():
            sys.exit(f"error: file not found: {p}")
        reader = PdfReader(str(p))
        card = reader.pages[0]
        sheet, kind = make_sheet(writer, card)
        if args.guides:
            add_guides(sheet)
        print(f"  {p.name}: 10-up sheet added ({kind})")

    with open(args.output, "wb") as f:
        writer.write(f)
    print(f"wrote {args.output} ({len(args.cards)} sheet(s))")


if __name__ == "__main__":
    main()

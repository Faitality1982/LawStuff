"""Turn scanned Big Book PDFs into ordered page images for the spread deck.

Each scan is one physical page including the punched margin. In a spiral book
the punch holes sit on the binding edge, so hole position tells us which side
of the spread a page belongs on:

    holes on the RIGHT edge  -> verso  (left-hand page)
    holes on the LEFT edge   -> recto  (right-hand page)

Writes page JPEGs to pages/ and a manifest to pages.json.
"""
import io, json, os, re, sys
import pymupdf
from PIL import Image

DPI = 130
OUT = "pages"

def hole_side(im):
    """Return 'left', 'right', or None by looking for dark blobs in the
    outer 7% of each edge. Punch holes scan as near-black circles."""
    g = im.convert("L").resize((im.width // 4, im.height // 4))
    w, h = g.size
    strip = max(4, int(w * 0.07))
    px = g.load()
    def darkness(x0, x1):
        dark = 0
        total = 0
        for y in range(int(h * 0.08), int(h * 0.92)):
            for x in range(x0, x1):
                total += 1
                if px[x, y] < 90:
                    dark += 1
        return dark / max(total, 1)
    left = darkness(0, strip)
    right = darkness(w - strip, w)
    if max(left, right) < 0.02:
        return None, left, right
    return ("left" if left > right else "right"), left, right

def main(pdfs):
    os.makedirs(OUT, exist_ok=True)
    manifest = []
    seq = 0
    for pdf in pdfs:
        doc = pymupdf.open(pdf)
        for i, page in enumerate(doc):
            pm = page.get_pixmap(dpi=DPI)
            im = Image.open(io.BytesIO(pm.tobytes("png"))).convert("RGB")
            side, dl, dr = hole_side(im)
            name = "p%04d.jpg" % seq
            im.save(os.path.join(OUT, name), quality=82, optimize=True)
            # OCR-free page-number guess from the PDF text layer, if any
            txt = page.get_text().strip()
            m = re.match(r"^\s*(\d{1,3})\b", txt) or re.search(r"\b(\d{1,3})\s*$", txt[:200])
            manifest.append({
                "seq": seq, "file": name, "source": os.path.basename(pdf),
                "srcPage": i, "w": im.width, "h": im.height,
                "holes": side, "darkLeft": round(dl, 4), "darkRight": round(dr, 4),
                "pageNoGuess": m.group(1) if m else None,
                "side": "recto" if side == "left" else ("verso" if side == "right" else None),
            })
            seq += 1
        doc.close()
    json.dump(manifest, open("pages.json", "w"), indent=1)
    have = sum(1 for p in manifest if p["holes"])
    print("rendered %d pages; hole side detected on %d" % (len(manifest), have))

if __name__ == "__main__":
    main(sys.argv[1:])

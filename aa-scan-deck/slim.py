"""Render a lighter set of page images for the combined deck.

The section decks stay at full quality; this exists so the whole book fits in
one file small enough to email and download.
"""
import io, json, os, sys
import pymupdf
from PIL import Image

DPI = int(sys.argv[1]) if len(sys.argv) > 1 else 108
Q   = int(sys.argv[2]) if len(sys.argv) > 2 else 72
OUT = "pages-slim"

SRC = [
 "scans/5215354a-08182026_ALCOHOLIC_S_ANONYMOUS.pdf",
 "scans/43630d84-08182026_THERE_IS__SOLUTION_D_Njj__t7_i5uri_Cr__ft_VG_Lt_6TLi.pdf",
 "scans/231419ad-08182026___INTO_ACTION_73_invariably_they_got_drunk._Having_per.pdf",
 "scans/6c79fd5d-08182026_TO_EMPLOYERS_141_normal_will_do_incredible_things._Aft.pdf",
]

os.makedirs(OUT, exist_ok=True)
seq = 0
for pdf in SRC:
    doc = pymupdf.open(pdf)
    for page in doc:
        pm = page.get_pixmap(dpi=DPI)
        im = Image.open(io.BytesIO(pm.tobytes("png"))).convert("RGB")
        im.save(os.path.join(OUT, "p%04d.jpg" % seq), quality=Q, optimize=True, progressive=True)
        seq += 1
    doc.close()
print("wrote %d pages at %d dpi q%d" % (seq, DPI, Q))

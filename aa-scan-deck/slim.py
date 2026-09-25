"""Render a lighter set of page images for the combined deck.

The section decks stay at full quality; this exists so the whole book fits in
one file small enough to email and download.

Source is `pages/` rather than the scan PDFs. The PDFs are gone and `scans/`
is gitignored, so the only surviving copy of the pages is the JPEGs
`recover.py` pulls back out of the committed section decks. Those renders are
SRC_DPI, so a target dpi is just a rescale of them.

    python3 slim.py 98 66      # 98 dpi, JPEG quality 66
"""
import os, sys
from PIL import Image

SRC_DPI = 130           # what prep.py rendered pages/ at
SRC = os.environ.get("SLIM_SRC", "pages")
OUT = os.environ.get("SLIM_OUT", "pages-slim")

DPI = int(sys.argv[1]) if len(sys.argv) > 1 else 108
Q   = int(sys.argv[2]) if len(sys.argv) > 2 else 72

if not os.path.isdir(SRC):
    raise SystemExit("no %s/ - run `python3 recover.py` first" % SRC)

scale = min(1.0, DPI / SRC_DPI)
os.makedirs(OUT, exist_ok=True)

names = sorted(n for n in os.listdir(SRC) if n.lower().endswith(".jpg"))
for name in names:
    im = Image.open(os.path.join(SRC, name)).convert("RGB")
    if scale < 1.0:
        im = im.resize(
            (max(1, round(im.width * scale)), max(1, round(im.height * scale))),
            Image.LANCZOS,
        )
    im.save(os.path.join(OUT, name), quality=Q, optimize=True, progressive=True)

print("wrote %d pages at %d dpi q%d (from %s at %d dpi)"
      % (len(names), DPI, Q, SRC, SRC_DPI))

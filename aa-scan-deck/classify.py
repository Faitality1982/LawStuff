"""Re-classify rendered pages as verso/recto from the punch-hole edge.

In a spiral book the holes are on the binding edge:
    holes on the RIGHT  -> verso (left-hand page)
    holes on the LEFT   -> recto (right-hand page)

Signal: down the outer 3.5% of each edge, the row-mean brightness of a punched
edge swings hard (hole, paper, hole, paper) while a plain margin is flat.
"""
import json
import numpy as np
from PIL import Image

STRIP = 0.035
FLOOR = 5.0     # absolute std below this is a plain margin
RATIO = 1.6     # winner must beat loser by this much

def edge_std(path):
    g = np.asarray(Image.open(path).convert("L").resize((400, 600)), dtype=float)
    h, w = g.shape
    s = max(3, int(w * STRIP))
    band = slice(int(h * 0.05), int(h * 0.95))
    return (g[band, 0:s].mean(axis=1).std(), g[band, w - s:w].mean(axis=1).std())

ROMAN = [(1000,"m"),(900,"cm"),(500,"d"),(400,"cd"),(100,"c"),(90,"xc"),
          (50,"l"),(40,"xl"),(10,"x"),(9,"ix"),(5,"v"),(4,"iv"),(1,"i")]

def roman(n):
    out = ""
    for v, s in ROMAN:
        while n >= v:
            out += s
            n -= v
    return out

def label(seq):
    """Printed folio for each scan, anchored on four verified points:
    seq 2 = v (Contents), seq 3 = vii (Preface), seq 24 = xxviii, seq 25 = 1.
    Front matter runs contiguously from seq 3; arabic runs contiguously from
    seq 25 to the end. Title and copyright pages carry no printed number."""
    if seq <= 1:
        return ""
    if seq == 2:
        return "v"
    if seq <= 24:
        return roman(seq + 4)      # seq 3 -> vii
    return str(seq - 24)           # seq 25 -> 1


def main():
    pages = json.load(open("pages.json"))
    for p in pages:
        l, r = edge_std("pages/" + p["file"])
        hi, lo = max(l, r), min(l, r)
        side = None
        if hi > FLOOR and hi > RATIO * lo:
            side = "left" if l > r else "right"
        elif hi > FLOOR:
            # Handwriting down the outer margin can make both edges busy.
            # Fall back to the stronger edge and flag it for review.
            side = "left" if l > r else "right"
            p["weak"] = True
        p["holes"] = side
        p["edgeL"], p["edgeR"] = round(l, 2), round(r, 2)
        p["side"] = {"left": "recto", "right": "verso"}.get(side)
        p["label"] = label(p["seq"])
    json.dump(pages, open("pages.json", "w"), indent=1)
    n = sum(1 for p in pages if p["side"])
    print("classified %d/%d" % (n, len(pages)))
    print("unresolved:", [p["seq"] for p in pages if not p["side"]])

if __name__ == "__main__":
    main()

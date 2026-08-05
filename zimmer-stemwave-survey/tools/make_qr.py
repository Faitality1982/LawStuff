#!/usr/bin/env python3
"""
Generate the QR codes for the Stem Wave survey.

    pip install segno
    python3 tools/make_qr.py --base https://zimmerstemwave.netlify.app

Produces, per placement, an SVG for print and a PNG for screen/email.

Why the settings are what they are:
  * error="H"  -- ~30% of the code is recoverable. Survives a coffee ring, a
                  thumbprint, and a logo dropped in the middle. The counter
                  card will get handled all day.
  * border=4   -- the quiet zone. Four modules is the spec minimum. Most
                  botched QR prints are a designer cropping this off.
  * black on white, no gradients or brand colours. Scanners want contrast and
    front-office lighting is bad.
"""

import argparse
import pathlib
import sys

try:
    import segno
except ImportError:
    sys.exit("segno is not installed.  pip install segno")


# label -> (src param, what it's for, printed QR width in inches)
PLACEMENTS = {
    "counter":  ("counter", "Front desk / check-in and check-out (primary)", 1.5),
    "room1":    ("room1",   "Treatment room 1", 2.5),
    "room2":    ("room2",   "Treatment room 2", 2.5),
    "card":     ("card",    "Business-card handout", 1.0),
    "poster":   ("poster",  "Waiting room poster", 4.0),
    "email":    ("email",   "Email to patient list", 0.0),
    "sms":      ("sms",     "Text message", 0.0),
}


def verify(paths: list[tuple[str, pathlib.Path, str]]) -> bool:
    """Read every generated code back and confirm it carries the right URL.

    Uses zxing-cpp, the decoder family Android and most scanning apps are built
    on. Optional -- skipped with a note if it isn't installed.

    A note on OpenCV: its QR detector fails outright on some perfectly valid
    symbols (it could not even *locate* the room1 code, at any resolution,
    while zxing read it fine and it survived blur/rotation/downscale testing
    exactly as well as its siblings). Do not use cv2 as the arbiter here.
    """
    try:
        import cv2
        import zxingcpp
    except ImportError:
        print("  (skipping read-back check: pip install zxing-cpp opencv-python-headless)\n")
        return True

    ok = True
    for name, png, expected in paths:
        res = zxingcpp.read_barcodes(cv2.imread(str(png)))
        got = res[0].text if res else ""
        if got != expected:
            ok = False
            print(f"  VERIFY FAILED  {name}: expected {expected!r}, decoded {got!r}")
    print("  Read-back check: all codes decode to the correct URL.\n" if ok else "")
    return ok


def build(base: str, outdir: pathlib.Path, only: list[str] | None) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    base = base.rstrip("/")

    rows = []
    made: list[tuple[str, pathlib.Path, str]] = []
    for name, (src, purpose, inches) in PLACEMENTS.items():
        if only and name not in only:
            continue

        url = f"{base}/?src={src}"
        qr = segno.make(url, error="H")

        svg = outdir / f"qr-{name}.svg"
        png = outdir / f"qr-{name}.png"

        # scale is in SVG user units; the printed size is set in the layout,
        # so this just needs to be crisp and reasonably sized.
        qr.save(svg, scale=10, border=4, dark="#000000", light="#ffffff")
        qr.save(png, scale=20, border=4, dark="#000000", light="#ffffff")

        rows.append((name, url, purpose, inches, qr.version, qr.symbol_size(scale=1)[0]))
        made.append((name, png, url))

    w = max(max(len(r[0]) for r in rows) + 2, len("PLACEMENT") + 2)
    print(f"\nWrote {len(rows) * 2} files to {outdir}/\n")
    verify(made)
    print(f"{'PLACEMENT'.ljust(w)}{'PRINT AT'.ljust(11)}URL")
    print("-" * 78)
    for name, url, purpose, inches, ver, mods in rows:
        size = f'{inches}"'.ljust(11) if inches else "screen".ljust(11)
        print(f"{name.ljust(w)}{size}{url}")
        print(f"{''.ljust(w)}{''.ljust(11)}{purpose}  (v{ver}, {mods} modules)")
    print()
    print("Sizing rule: QR width  ~=  scan distance / 10.")
    print("TEST EVERY PRINTED PIECE with a real phone before the print run —")
    print("iOS and Android both, and one older handset if you can find one.")
    print("Print the URL in readable text under every code; a meaningful share")
    print("of this age group will type it rather than scan.\n")


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate survey QR codes.")
    ap.add_argument(
        "--base",
        default="https://zimmerstemwave.netlify.app",
        help="Base URL of the deployed survey (default: %(default)s)",
    )
    ap.add_argument(
        "--out",
        default="qr",
        type=pathlib.Path,
        help="Output directory (default: %(default)s)",
    )
    ap.add_argument(
        "--only",
        nargs="*",
        choices=sorted(PLACEMENTS),
        help="Generate only these placements (default: all)",
    )
    args = ap.parse_args()

    if not args.base.startswith(("http://", "https://")):
        sys.exit("--base must start with http:// or https://")

    build(args.base, args.out, args.only)


if __name__ == "__main__":
    main()

# Avery 8371 Business Card Imposition

Tiles business-card PDFs (e.g. OvernightPrints proofs) onto **Avery 8371**
sheets: US Letter, 10 cards per sheet, 2 columns × 5 rows, each card
3.5" × 2", edge-to-edge with 0.75" side and 0.5" top/bottom margins.

## Setup

```
pip install pypdf reportlab
```

## Usage

```
python avery8371.py "HighCaliber-business-card-Casey-OvernightPrints.pdf" ^
                    "HighCaliber-business-card-Will-OvernightPrints.pdf" ^
                    -o HighCaliber-Avery8371.pdf
```

(Use `\` instead of `^` for line continuation on Mac/Linux, or put it all on
one line.)

Each input PDF gets its own sheet with 10 copies of the card. Options:

- `-o output.pdf` — output file name (default `avery8371-sheets.pdf`)
- `--guides` — add faint cut-guide ticks in the margins (useful when
  printing on plain cardstock instead of pre-perforated Avery stock)

## What it handles automatically

- **Bleed cropping** — print-shop files at standard full-bleed sizes
  (3.75" × 2.25", 3.7" × 2.2", 3.625" × 2.125") are center-cropped to the
  3.5" × 2" trim size, so nothing spills onto neighboring cards.
- **Portrait cards** — rotated to landscape.
- **Other sizes** — scaled to fit and centered in the 3.5" × 2" slot.

## Printing

Print at **100% / Actual size** — do not use "Fit to page", which shrinks
the layout and misaligns it with the Avery perforations.

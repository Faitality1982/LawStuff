# Big Book Study — scanned page spreads

Displays the group's annotated Big Book two pages at a time, as an open spread,
for the Wednesday Big Book Study at 903 Court Street, Port Huron (A.A. District 23).

## The decks

| File | Covers | Spreads |
|---|---|---|
| `big-book-1-front-matter.pptx` | Title page through page 16 | 22 |
| `big-book-2-there-is-a-solution.pptx` | Pages 17–72 | 29 |
| `big-book-3-into-action.pptx` | Pages 73–140 | 35 |
| `big-book-4-to-employers.pptx` | Pages 141 to the end | 27 |

Split by section because a single deck runs 72 MB. Each one opens with the title
slide and the A.A.W.S. copyright notice, then the spreads. Every spread carries
the copyright line as a footer.

216 scanned pages, all accounted for.

## Layout and margins

Slide is 13.33 x 7.50 in. The pages are 0.675 as wide as they are tall, so height
is the binding constraint and the spread comes out **8.75 in wide**, leaving:

| Where | Free |
|---|---|
| Each side | **2.29 in** |
| Above | 0.35 in |
| Below | 0.65 in (footer uses 0.30) |

That side margin cannot be recovered by scaling — squeezing the top and bottom to
almost nothing only takes it from 2.29 to 2.02 in. So it is permanent space, and
the printed folio now sits in it, on the outer edge of each page the way a book
sets it. Roughly 1.6 in per side is still clear if more belongs there later.

## Page numbers

Derived, not OCR'd — the scans have no text layer. Anchored on four points read
off the pages themselves: seq 2 is `v` (Contents), seq 3 is `vii` (Preface),
seq 24 is `xxviii`, seq 25 is `1` (Chapter 1). Front matter runs contiguously in
roman numerals from seq 3; arabic runs contiguously from seq 25 to `191` at the
end. Title and copyright pages carry no printed number and get none.

Every derived label agrees with its detected side — odd numbers land on rectos,
even on versos, all 216 of them.

## Rebuilding

```bash
npm install pptxgenjs
pip install pymupdf pillow numpy
python3 prep.py scans/*.pdf     # PDFs -> pages/*.jpg + pages.json
python3 classify.py             # verso/recto from the punch-hole edge
node build.js                   # four section decks
COMBINED=1 node build.js        # one 72 MB deck instead
```

`pages/` and `scans/` are gitignored — only the generators are tracked.

## How pages are ordered

Each scan is one physical page including the punched margin. In a spiral book the
holes sit on the binding edge, so hole position says which side of the spread a
page belongs on:

- holes on the **right** → verso (left-hand page, even page number)
- holes on the **left** → recto (right-hand page, odd page number)

`classify.py` reads the outer 3.5% of each edge and measures how much the row
brightness swings down the strip. A punched edge alternates hole/paper and swings
hard; a plain margin is flat. All 216 pages classified; the sequence alternates
verso/recto almost perfectly, breaking only where an unscanned blank verso leaves
two rectos back to back — which is what the book actually does.

Two pages (seq 44, 45) needed the fallback rule because handwritten margin notes
run down the *outer* edge and made both sides busy. Both were checked by eye
against their printed page numbers (20 and 21) and are correct. They are flagged
`"weak": true` in `pages.json`.

A verso followed by a recto becomes a spread. Anything else stands alone — the
title page, and rectos whose blank verso was not scanned.

## Permission

A.A.W.S. granted permission by email (Drew Deetz, Intellectual Property
Administrator, General Service Office) to screen-share A.A. literature during an
A.A. meeting, provided their copyright notice is displayed. It is on slide 2 of
every deck and in the footer of every spread.

Intended for use inside the group's own class. Not for distribution.

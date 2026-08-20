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

Slide is 13.33 x 7.50 in. The pages run the full height — a 7.38 in band, 0.06 in
clear top and bottom — because in a room the size of the page is the whole point.

The pages are 0.675 as wide as they are tall, so height is the binding constraint
and the spread comes out **9.96 in wide**, leaving **1.69 in on each side**. That
side margin is permanent; it cannot be traded away for a taller page.

It carries two things: the printed folio, set on the outer edge of each page the
way a book sets it, and the copyright notice, turned on its side and running up
the far left edge. Nothing sits along the bottom any more, which is what let the
pages grow.

## Page numbers

Derived, not OCR'd — the scans have no text layer. Anchored on four points read
off the pages themselves: seq 2 is `v` (Contents), seq 3 is `vii` (Preface),
seq 24 is `xxviii`, seq 25 is `1` (Chapter 1). Front matter runs contiguously in
roman numerals from seq 3; arabic runs contiguously from seq 25 to `191` at the
end. Title and copyright pages carry no printed number and get none.

Every derived label agrees with its detected side — odd numbers land on rectos,
even on versos, all 216 of them. That check is the proof of completeness: a
missing page would flip the parity of everything after it and show up at once.

### Audit: all 216 pages verified

Every folio was read off the scans directly (corner crops, four contact sheets).
The run is contiguous in three stretches, broken only by two blank versos that
were never scanned:

| Stretch | Pages | Note |
|---|---|---|
| seq 2 | `v` | Contents. **`vi` is its blank back — not scanned.** |
| seq 3–24 | `vii`–`xxviii` | contiguous |
| seq 25–199 | `1`–`175` | contiguous |
| seq 200–215 | `177`–`192` | **`176` is blank — not scanned.** seq 200 is the APPENDICES divider. |

Nothing is missing from the book itself. The two absent numbers are blank pages.

Note the scan runs well past page 164 — 164 ends "A Vision for You," then Doctor
Bob's Nightmare (165–175) and the Appendices follow, all present.

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

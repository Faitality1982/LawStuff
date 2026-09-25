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

Each section deck also carries the photo slides that fall inside its pages, so
they run 35, 36, 39 and 30 slides respectively.

Plus `aa-big-book-spreads.pptx` — the whole book in one file, 131 slides
(110 spreads, the title and copyright slides, and the 19 photo slides).
It is built from a lighter set of page images (`slim.py`, 98 dpi) so it lands at
24.1 MB, under both Gmail's 25 MB attachment cap and the 30 MiB chat limit. At
projection size it is indistinguishable from the section decks; those keep the
full 130 dpi renders.

Split by section because a full-quality combined deck runs 48 MB. Each one opens with the title
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

The scan PDFs are gone, so a rebuild starts from `recover.py`, which pulls the
216 page images back out of the committed section decks. `prep.py` and
`classify.py` are kept for the record but need `scans/`, which no longer exists.

```bash
npm install pptxgenjs
pip install pymupdf pillow numpy python-pptx
python3 recover.py              # committed decks -> pages/*.jpg
python3 fetch_photos.py         # photo-slide images -> photos/ + photos.json
node build.js                   # four section decks
python3 slim.py 98 66           # lighter page images for the combined deck
PAGES_DIR=pages-slim COMBINED=1 node build.js   # the whole book, under 25 MB
```

`PHOTOS_ONLY=1 node build.js` writes `photo-slides-proof.pptx` — the 19 photo
slides on their own. Checking a frame, an image size or a credit line should not
cost a rebuild of 216 page images.

`slim.py` rescales `pages/` rather than re-rendering the PDFs, for the same
reason: `pages/` is the only surviving copy of the scans.

`pages/`, `pages-slim/` and `scans/` are gitignored. `photos/` is tracked — the
images are small and the deck cannot be rebuilt without them.

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

## Photo slides

Chris Zimmer's placement list (his notebook, two pages) is in `PHOTOS` in
`build.js`. Each entry inserts a slide immediately after the spread carrying that
page, with a caption and a `Credit:` line.

Nineteen of them: Rowland Hazard (xi), Oxford Group (xii), A.A. Number Three
(xiii), Clarence Snyder (xvii), Silkworth (xxii), Towns Hospital (xxiii),
alcohol-metabolism diagram (xxv), Bill W. and the Thetcher tombstone (1),
Leonard Strong (7), Ebby Thatcher (9), Fellowship diagram (17), Carl Jung (26),
William James (28), handout sheet (63), inventory handouts (64), Eleventh Step
inventory (86), Hank Parkhurst (136), Dr. Bob (165).

**All fourteen photographs are in.** The five remaining frames are the two
diagrams and the three handouts — original artwork, to be drawn rather than
sourced.

Zimmer's emails are links to web pages rather than attached images, so the
photographs were taken from the pages he sent (and, where a better-licensed copy
existed, from Wikimedia Commons or the Library of Congress instead).

### Where each photograph came from

`fetch_photos.py` downloads them and writes `photos.json`, which `build.js` reads
to print the credit under each frame. The manifest is the provenance record: an
image in `photos/` that is not in the manifest never reaches a slide. Exact
source titles, never a search — Commons holds three other Rowland Hazards, and
the Library of Congress item captioned "[William James, half-length portrait]"
is a Hogarth painting of an 18th-century namesake, not the psychologist.

| Slug | Source | Licence |
|---|---|---|
| `carl-jung` | Wikimedia Commons (ETH-Bibliothek) | Public Domain Mark |
| `william-james` | Wikimedia Commons (National Portrait Gallery) | Public domain |
| `thetcher-tombstone` | Wikimedia Commons | CC BY-SA 3.0 |
| `oxford-group` | Library of Congress, Harris & Ewing | No known restrictions |
| `towns-hospital` | Flickr, Eden/Janine/Jim | CC BY 2.0 |
| `bill-w` | aamidsurrey.org.uk | not cleared |
| `dr-bob` | aamidsurrey.org.uk | not cleared |
| `silkworth` | aamidsurrey.org.uk | not cleared |
| `ebby-thatcher` | aamidsurrey.org.uk | not cleared |
| `rowland-hazard` | aamidsurrey.org.uk | not cleared |
| `clarence-snyder` | aamidsurrey.org.uk | not cleared |
| `aa-number-three` | aamidsurrey.org.uk | not cleared |
| `hank-parkhurst` | aamidsurrey.org.uk | not cleared |
| `leonard-strong` | aalkies.wordpress.com | not cleared |

The nine marked *not cleared* are early A.A. archival photographs. The sites
hosting them are not the rights holders and grant no licence; the rights almost
certainly sit with A.A. archives. `photos.json` records that on every one rather
than implying a permission nobody gave. **This is a separate question from the
A.A.W.S. literature permission below, which covers the scanned pages and not
these photographs.** They are in the deck because they are the pictures Zimmer
picked, for display inside the group's own class.

`oxford-group` is a portrait of Frank Buchman, the group's founder, and the slide
says so — there is no properly-sourced photograph of the group itself.
`towns-hospital` is the building as it stands today, and the slide says that too.

### Two corrections

- The Leonard Strong slide read "Dr. Leonard Strong, M.D." He was an osteopath,
  not a physician; it now reads "Dr. Leonard V. Strong, Jr."
- Earlier passes corrected three spellings against the record: Dotson (not
  Dodson), Snyder (not Synder), Carl Jung (not Karl).

### Sizing

The archival portraits are small — several are under 250 px on the long edge,
which is simply how they survive. `build.js` never draws one larger than
`MIN_DPI` (72) would justify, so a 173 px portrait lands at 2.4 in rather than
being stretched across the frame. A small sharp photograph reads from the back of
the room; a big soft one does not.

## Recovering the page images

`pages/` and `pages-slim/` are gitignored, so a fresh container has the decks but
not the JPEGs. `recover.py` pulls every page image back out of the committed
section decks in slide order and rewrites `pages/`. No need for the source PDFs.

## Permission

A.A.W.S. granted permission by email (Drew Deetz, Intellectual Property
Administrator, General Service Office) to screen-share A.A. literature during an
A.A. meeting, provided their copyright notice is displayed. It is on slide 2 of
every deck and in the footer of every spread.

Intended for use inside the group's own class. Not for distribution.

That permission covers the scanned pages. It does not cover the nine early A.A.
photographs marked *not cleared* above, whose rights sit elsewhere — see the
photo-slide section.

---

## Next step (picking this up in a new session)

The deck is complete and correct, and all fourteen photographs are in. The only
outstanding work is the five artwork frames.

1. `pip install python-pptx pillow numpy pymupdf && npm install pptxgenjs`
2. `python3 recover.py` — rebuilds `pages/` from the committed section decks
3. `python3 fetch_photos.py` — rebuilds `photos/` (only needed on a fresh clone
   if `photos/` is somehow missing; it is tracked)
4. `node build.js` for the four section decks;
   `python3 slim.py 98 66 && PAGES_DIR=pages-slim COMBINED=1 node build.js`
   for the single combined file (keep it under 25 MB so it can be emailed).

Use `PHOTOS_ONLY=1 node build.js` while working on a photo or artwork frame —
it writes `photo-slides-proof.pptx`, the 19 slides on their own, in a second.

### Still to be made (no sourcing required)

Two diagrams and three handouts from Zimmer's list are original artwork, not
photographs: alcohol metabolism (xxv), the Fellowship (17), handout sheet (63),
inventory handouts (64), Eleventh Step inventory (86). Frames are in place;
content to be specified. Zimmer's notebook gives no detail beyond the titles, so
what each one should say still has to come from him or from the group.

A drawn frame needs no change to `build.js`: save the artwork as
`photos/<slug>.jpg`, add an entry to `photos.json` with its `file`, `px` and a
`credit`, and it renders like any photograph. The slugs are
`alcohol-metabolism`, `fellowship`, `handout-sheet`, `inventory-handouts` and
`eleventh-step-inventory`.

### Size budget

The combined deck sits at 24.1 MB against the 25 MB mail cap. The page images
are the bulk of that; `slim.py 98 66` is what holds it down. The photographs are
about 1 MB of the total — `MAXPX` in `fetch_photos.py` caps them at 1000 px,
which is already more than a 1080p projector can show at the size they are drawn.
If the artwork frames push it over, drop `slim.py` to 95 dpi before touching the
photographs.

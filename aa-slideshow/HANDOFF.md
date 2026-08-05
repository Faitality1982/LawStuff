# AA Big Book Study Slideshow — Handoff

Everything a fresh session needs to pick this up. Read this first, then `SPEC.md`
for the slide-by-slide record.

---

## What this is

A PowerPoint deck for an A.A. Big Book study — the Wednesday 7:00 PM meeting,
District 23 (Port Huron, MI). Built page by page from photographs of a
hand-annotated spiral-bound Big Book (4th edition). Will photographs a page,
describes the highlighting and margin notes, and the deck grows by a few slides.

**Current state: 49 slides**, covering the front matter through Bill's Story
page 11. The work is ongoing — more pages are coming.

## Where everything lives

Repository `Faitality1982/LawStuff`, branch `claude/aa-slideshow-planning-7hxes4`,
directory `aa-slideshow/`:

| File | What it is |
|---|---|
| `aa-big-book-study.pptx` | The deliverable. 49 slides, 16:9, fully editable. |
| `build.js` | pptxgenjs generator. **The source of truth — edit this, never the .pptx.** |
| `render.py` | Reads a .pptx back and emits HTML for visual QA. |
| `preview.png` | Full-deck preview strip, one slide per 750px. |
| `SPEC.md` | Slide-by-slide record: sources, decisions, open questions. |
| `HANDOFF.md` | This file. |

## Rebuilding

```bash
cd aa-slideshow
npm install pptxgenjs                       # once
pip install python-pptx defusedxml Pillow   # once
node build.js                               # writes the .pptx
python3 /root/.claude/skills/pptx/scripts/office/validate.py aa-big-book-study.pptx
```

Visual QA, since **LibreOffice is broken in this sandbox** (`soffice` fails to
load any source file at all, not just ours — don't waste time on it):

```bash
python3 render.py aa-big-book-study.pptx preview.html
# headless Chrome's viewport is ~90px SHORTER than --window-size.
# Ask for extra height and crop, or the bottom of every slide is silently lost.
/opt/pw-browsers/chromium-1194/chrome-linux/chrome --headless --no-sandbox \
  --disable-gpu --hide-scrollbars --window-size=1345,$((750*NSLIDES+120)) \
  --screenshot=full.png "file://$PWD/preview.html"
```

`render.py` is approximate by design — it does not draw shape outlines, and it
flattens double underlines to single. When a mark matters, verify it in the
slide XML instead:

```bash
python3 -c "import zipfile;print(zipfile.ZipFile('aa-big-book-study.pptx').read('ppt/slides/slide19.xml').decode())" | grep -o 'u=\"dbl\"'
```

Previews get large. The full strip has been rejected by the file-send endpoint
above ~1.5 MB — crop to just the new slides and save as JPEG at ~0.6 scale.

---

## The permission situation — read before adding anything

A.A.W.S. granted permission by email (Drew Deetz, Intellectual Property
Administrator, General Service Office). The grant is **narrow**: no objection to
screen-sharing A.A. literature found on aa.org or in authorized e-books during
an A.A. meeting, **provided their copyright disclaimer is displayed.**

The disclaimer runs in full on slide 1 and in short form as a footer on every
content slide. `addFooter()` in `build.js` — call it on every new slide.

It covers **literature**. It does not cover photographs. Slide 10 therefore ships
with an empty dashed frame where the hospital-bed picture of Bill D. goes; Will
places that image himself.

---

## Design system

Do not invent new visual vocabulary. The deck has a grammar and it is consistent
across 49 slides:

**Colors** — navy `1B2A41` (text, margin notes), green `2C5F2D` / tint `EAF1E6`,
blue `065A82` / tint `E2EDF4`, magenta `9E1B60` / tint `F9E3EE`, yellow tint
`FFF7D9`, muted `6B7A8C`. Fonts: Cambria headings, Calibri body.

**The rules that matter:**

1. **Card tint = the highlighter color on the page.** Green highlight → green
   card. Blue → blue. Pink → magenta. Yellow → yellow.
2. **Pink highlighter always means a "must."** It renders magenta, and the must
   itself gets pulled out into a solid magenta bar labeled `A "MUST"`.
3. **A solid bar takes the color of the highlighter it came from** — magenta for
   a must, blue for the IMPORTANT paragraph on slide 30 — **and navy when the bar
   carries Will's margin note instead of book text.**
4. **Light backgrounds.** The source page is marked "handouts"; the deck has to
   print. Dark is reserved for section openers (slides 1 and 32).
5. Underlines and double underlines in the book are reproduced as underlines and
   double underlines (`underline: { style: "dbl" }`). Circled words are set bold
   in the section color.

**Helpers in `build.js`** — use these rather than hand-placing shapes:

- `header(s, title, sub, box)` — title, subtitle, and the page-reference corner box
- `addFooter(s)` — the required copyright footer. **Every slide.**
- `cite(s, y, tail)` — the italic source line
- `marginNote(s, y, h, text)` — full-width navy bar for a margin note
- `markedCard(s, {y,h,tint,dark,runs,note,size})` — tinted card with the margin
  note set small in the right gutter. The workhorse for Bill's Story, where
  nearly every highlight is tagged with a "level."
- `idCard(s, {x,y,w,h,label,name,sub})` — compact navy card resolving a circled
  word the book leaves unnamed

**Speaker notes carry the teaching.** Every slide has `addNotes()` with what to
say — the history, the connections between passages, where to slow down. This is
half the value of the deck. Keep writing them.

---

## Working method

1. Will sends photographs of a page (or several) with a spoken description of the
   highlighting and margin notes.
2. Read the images. **Transcribe carefully** — quote the book exactly. When an
   annotation is unreadable at full-page scale, crop and enlarge it with PIL
   before committing to a reading; several have been resolved that way.
3. Build the slides, validate, render, look at the result.
4. Update `SPEC.md`, commit, push, send the .pptx.
5. Report what was built and flag anything uncertain.

**Verify facts before they go on a slide.** Two dates from the margin notes were
wrong and were corrected on the slide with the change recorded in the speaker
notes and `SPEC.md` (Rush 1776→1784, Trotter 1782→1804). Someone in that room has
a phone. Never silently alter Will's notes — change it, say so, and offer the
revert.

**Never reorder a quotation.** Caught once while bolding a phrase mid-sentence.
Quote exactly, bold in place.

---

## Open items

- **Slide 31 was built from dictation, not a photograph.** Verify the wording of
  the closing line of The Doctor's Opinion and add the page number.
- **The gloss on "pot"** (slide 33) reads **PAIL** when enlarged, which fits the
  epitaph. The final letter is not perfectly clear — confirm against the book.
- **"Show Sheet"** — a star at the top of p. xxv is a presenter cue. It sits in
  slide 27's speaker notes. Nobody knows yet which sheet it means; if it is a
  handout that belongs in the deck, it needs its own slide.
- **"Newcomers"** (slide 13) is attached to the "not a religious organization"
  paragraph. The note has no leader line and could belong to the paragraph above.
- **Offered, not built:** a recap slide listing the "levels" in order — quitting
  projects → egotistic → carousing → insanity → sponging off family → blinded to
  life → will power → "was I crazy?" It would make Will's organizing frame visible
  at a glance. Waiting on his call.
- **aadistrict23.org returns HTTP 403 to automated fetches.** Not a proxy problem;
  the origin blocks bots. Worth raising with Will separately — a meeting list that
  crawlers cannot read is invisible to search and to AI answers at 2 a.m.

---

## Voice

Will's notes are the spine of this deck, not decoration. He is teaching from a
book he knows well, and the margin notes are what he plans to say out loud. Build
the slides around them, keep the book's text exact, and keep the speaker notes in
his register — direct, unsentimental, no inspirational filler.

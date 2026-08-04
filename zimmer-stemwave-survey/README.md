# Zimmer Chiropractic — Stem Wave interest survey

Anonymous patient survey measuring whether there's demand for a Stem Wave
(SoftWave/shockwave) therapy offering at **$49 discovery visit + $2,760
package = $2,809**, non-covered, paid up front.

Reached by QR code at the front desk. Nothing links to it.

- **Deploy steps:** [DEPLOY.md](DEPLOY.md)
- **Methodology and reasoning:** the plan doc this was built from

---

## What it's for

Four business questions, in order:

1. **Is there a treatable population?** — pain location and chronicity
2. **Is the need unmet?** — what they've tried, satisfaction, surgery status
3. **Is there interest?** — neutral concept description, then appeal
4. **Will they pay, and how much?** — Van Westendorp, then price reveal, then
   purchase probability

---

## Three things that must not be "optimised away"

These aren't style choices. Breaking any one of them makes the data worthless
or creates a compliance problem.

### 1. Price questions come before the price reveal

The Van Westendorp screen (`vw`) is ordered before the reveal screen (`reveal`)
in `questions.js`. Show someone $2,809 first and every dollar answer they give
clusters around $2,809 — that's anchoring, and it turns the pricing block into
an expensive way to learn nothing.

If you reorder anything in `questions.js`, keep `vw` before `reveal`.

### 2. The survey is anonymous; leads are separate

Neither Netlify nor Cloudflare will sign a HIPAA BAA below Enterprise. So no
identifiable health information may land in this datastore. The design:

- `responses` — anonymous. No name, contact, DOB, or chart number. Broad pain
  checkboxes, not diagnoses. Anonymous health data isn't PHI.
- `leads` — the optional callback request. Separate endpoint, separate table,
  **no key linking it to a response**, and a day-only date so the two can't be
  rejoined by timestamp.

`submit.js` actively rejects any answer key named `name`, `email`, `phone`,
`contact`, `dob`, or `address`, so a future edit to `questions.js` can't quietly
break this.

Do not add a foreign key. Do not "improve" `leads.created_day` into a timestamp.

### 3. Intent is measured on an 11-point probability scale

Stated purchase intent overstates real behaviour, badly and consistently. `e1`
and `e2` use a 0–10 Juster scale, and `analyze.py` reports top-box (9–10)
separately from the mean. **Only 9s and 10s are real intent.** Verbal scales
("very likely / somewhat likely") are measurably worse and shouldn't replace it.

---

## Layout

```
public/
  index.html      shell
  styles.css      accessibility rules are load-bearing — see the file header
  config.js       branding, pricing, feature flags     <- edit here
  questions.js    the question bank, as data           <- and here
  survey.js       the engine (shouldn't need touching)
functions/api/
  submit.js       POST anonymous responses
  lead.js         POST optional contact, separate table
  export.js       GET  CSV/JSON, gated by EXPORT_KEY secret
tools/
  make_qr.py      QR codes per placement, ECC-H
  analyze.py      Van Westendorp + Juster + segments
  make_fixture.py synthetic data for testing the analyser
schema.sql        D1 tables
```

Vanilla HTML/CSS/JS. No framework, no build step, no dependencies. One page,
~20 screens of show/hide — React here would be a build pipeline in exchange for
nothing, and this needs to still run untouched in a year.

---

## Editing the survey

**Change wording, add or remove a question, reorder:** `questions.js` only.
It's data — `survey.js` renders whatever's there.

**Change branding, prices, or turn features on and off:** `config.js`.

Screen types available: `info`, `single`, `multi` (with optional `max` and
`exclusive` options), `scale`, `currencyGroup`, `text`. Conditional display via
`showIf(answers)`.

If you add a branching question, add its assumed answer to
`config.progressAssumes` — otherwise the progress counter grows mid-survey when
the branch opens, which reads as "this got longer" and costs completions.

---

## Accessibility

Respondents skew older and are, by definition, in pain and standing at a
counter. `styles.css` enforces 18px+ body text, 48px+ tap targets, real form
inputs, and **buttons rather than a drag slider** for the 0–10 scales — sliders
are miserable with a tremor, a splint, or cold hands.

---

## Testing

Local dev with a real D1 database:

```bash
npm install
npm run db:init:local
npm run dev
```

Exercise the analyser without waiting for real responses:

```bash
python3 tools/make_fixture.py --n 140 --out fixture.csv
python3 tools/analyze.py fixture.csv
```

`make_fixture.py` output is **invented test data**. Never mix it with a real
export and never show its numbers to Dr. Zimmer.

The survey flow was verified end-to-end in Chromium at a 390×844 phone
viewport: both branch paths, resume, back-navigation, branch reversal, the
multi-select cap and exclusive options, the Van Westendorp ordering guard, the
payload shape, and that a submitted survey can't be resumed into.

---

## Status

Built and tested locally. **Not yet deployed** — the build environment had no
Cloudflare credentials. See [DEPLOY.md](DEPLOY.md).

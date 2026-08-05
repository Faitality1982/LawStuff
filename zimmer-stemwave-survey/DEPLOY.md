# Deploy runbook — Netlify

Everything here runs from a machine logged into Netlify. The build session had
no Netlify credentials, so **none of this has been executed** — treat every step
as unverified until you've run it.

About 20 minutes end to end.

---

## What you're deploying

| | |
|---|---|
| Static site | `public/` — the survey, plain HTML/CSS/JS, no build step |
| Functions | `netlify/functions/` — `submit` and `export`, Netlify Functions v2 |
| Database | **Netlify Blobs** — one JSON record per response, built in, free |
| URL | `https://zimmerstemwave.netlify.app` (the QR codes encode this) |

There is no separate database service to sign up for. Blobs is part of
Netlify and is enabled automatically the first time a function writes to it.

---

## 1. Deploy

Two ways. Pick one.

### A. From the Netlify UI, no terminal

A `netlify.toml` at the **root of the LawStuff repo** points Netlify at this
subdirectory, so the import needs no configuration:

1. app.netlify.com → **Add new project** → **Import an existing project**
2. GitHub → **LawStuff**
3. Branch: `claude/new-session-9smud0`
4. Leave every build setting alone — the root `netlify.toml` fills them in
5. **Deploy**

Then rename the site to `zimmerstemwave` under Project configuration → Change
site name, so it matches the URL the QR codes encode.

Every push to that branch redeploys automatically.

### B. From the CLI

```bash
cd zimmer-stemwave-survey
npm install
npx netlify login
npx netlify sites:create --name zimmerstemwave
npx netlify deploy --prod
```

Use `sites:create`, not `netlify init` — `init` wires up git-based CI builds,
which duplicates option A and means debugging a build that doesn't need to
exist.

If the name `zimmerstemwave` is taken, pick another and regenerate the QR
codes (step 4). Ten-second job — just don't print anything first.

---

## 2. Set the export key

This is the password that protects the response data. Without it the export
endpoint refuses to serve anything.

```bash
openssl rand -base64 32          # generate one
npx netlify env:set EXPORT_KEY 'paste-the-value-here'
npm run deploy                   # env vars apply to new deploys
```

Keep a copy in your password manager. There is no way to recover it from
Netlify's UI later — you'd just set a new one.

Check it works:

```bash
curl -sI "https://zimmerstemwave.netlify.app/api/export?key=WRONG" | head -1
# expect: HTTP/2 404   (a wrong key looks like a missing page, on purpose)

curl -sI "https://zimmerstemwave.netlify.app/api/export?key=YOUR_KEY" | head -1
# expect: HTTP/2 200
```

---

## 3. Walk the survey once, on a phone

Open <https://zimmerstemwave.netlify.app/?src=counter> **on an actual phone**,
not a desktop window. Confirm:

- [ ] Answering "No" to the first question skips the pain and pricing blocks
- [ ] The four price questions appear **before** the $2,809 reveal
- [ ] Entering the price boxes out of order warns you once, then lets you pass
- [ ] Choosing "Monthly payments" reveals the follow-up amount question
- [ ] The counter never counts *upward* (it may shrink from 21 to 9)
- [ ] Nothing anywhere asks for your name
- [ ] After submitting, closing and reopening gives a **fresh** survey

Then confirm the response actually landed:

```bash
EXPORT_KEY='your-key' npm run export
```

It writes `responses-YYYY-MM-DD.csv` and prints the record count. If that's 1,
the whole pipeline works.

**If the count is 0 but the survey said thank-you**, the write failed silently —
check `npx netlify logs:function submit`.

---

## 4. Regenerate the QR codes (only if the site name changed)

```bash
pip install segno zxing-cpp opencv-python-headless
python3 tools/make_qr.py --base https://YOUR-SITE.netlify.app
```

It reads every code back and confirms it decodes to the right URL before
printing the summary. Also update the printed URL in
`print/counter-card.html`.

**Then test the print physically.** Print the card at its real size, tape it
where it's going to live, and scan it with an iPhone, an Android, and the
oldest phone you can borrow — from where a patient actually stands, in the
light that's actually there. Front-desk lighting is usually bad and often has
window glare. Matte stock, no lamination: gloss is the most common reason a
card that scanned on your desk fails at the counter.

---

## 5. Before it goes live

- [ ] **Dr. Zimmer has signed off on the concept copy.** The description in
      `questions.js` (screen `concept`) is deliberately conservative — mechanism
      and logistics, no outcome claims. If he wants regenerative or stem-cell
      language added, that's his call to make knowingly: it's the kind of claim
      that draws FTC and state-board attention in chiropractic marketing.
- [ ] Confirm the pricing in `config.js` matches what he's considering.
- [ ] Confirm a monthly payment plan is genuinely on the table
      (`config.enableMonthlyPlanQuestion`). If he'd never offer financing, turn
      it off — but that question is the highest-value one in the survey, so
      turn it off only if the answer really is never.
- [ ] Brief the front desk (script below).

---

## 6. Front desk script

Print this and stick it by the terminal. The ask is the whole ballgame — a
receptionist who sounds unsure gets a 10% scan rate, one who sounds like it
matters gets 40%.

> "Before you go — Dr. Zimmer's thinking about bringing in a new treatment for
> chronic pain, and he wants to know what patients think before he commits.
> There's a QR code right here, it's about three minutes and it's completely
> anonymous. Would you mind?"

Points that matter, in order:

1. **"Dr. Zimmer wants to know what you think"** — it's a favour to someone
   they know, not a marketing funnel.
2. **"Completely anonymous"** — say it every time, it's the main objection.
   It's also literally true: there is no name field anywhere.
3. **"About three minutes"** — and it's true, so it stays true.
4. **"You can finish it in the car"** — answers save on their phone. This
   rescues everyone who says "I'm in a rush."

Do **not** describe the therapy off the cuff — the survey does that in
controlled wording, and an improvised description in the lobby is exactly the
claims risk the copy was written to avoid.

---

## 7. Getting the data out

```bash
EXPORT_KEY='your-key' npm run export
python3 tools/analyze.py responses-2026-09-15.csv
```

The analyser prints the price curves, the four Van Westendorp intersections,
where $2,760 falls relative to them, Juster purchase intent, and the key
segment (tried 3+ things, still dissatisfied). It writes
`price-sensitivity.png` if matplotlib is installed.

**Don't analyse before ~50 responses.** Van Westendorp curves need the sample;
below 30 the script says so and you should believe it.

Raw JSON instead of CSV:

```bash
curl -sS "https://zimmerstemwave.netlify.app/api/export?key=YOUR_KEY&format=json" | less
```

---

## 8. Optional: a branded URL

`zimmerstemwave.netlify.app` is fine and works immediately. If you'd rather the
card read `stemwave.logicloomllc.com`, that domain is on Cloudflare and it's one
record:

> Cloudflare DNS → logicloomllc.com → Add record
> Type `CNAME`, Name `stemwave`, Target `zimmerstemwave.netlify.app`,
> Proxy status **DNS only** (grey cloud — Netlify needs to terminate TLS itself)

Then in Netlify: Domain management → Add a domain → `stemwave.logicloomllc.com`.

Regenerate the QR codes afterwards (step 4). Don't print anything until the
final URL is decided.

---

## 9. Shutting it down

It's a temporary survey. When it's done:

1. Export the responses and check the CSV opens. **Netlify Blobs is the only
   copy.**
2. Pull the printed QR codes off the counter.
3. Delete the site in the Netlify UI (Site configuration → Danger zone), which
   deletes the blob store with it.

Step 1 first, and open the file before you delete anything.

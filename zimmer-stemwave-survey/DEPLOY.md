# Deploy runbook

Everything here has to run from a machine with your Cloudflare login. The build
session had no Cloudflare credentials, so none of this has been executed yet —
treat every step as unverified until you've run it.

Budget about 40 minutes for the first deploy, most of it waiting on DNS.

---

## 0. Where this lives

The target is **`logicloomllc.com`**, which is already fully on Cloudflare:

```
NS   alice.ns.cloudflare.com, mark.ns.cloudflare.com
A    logicloomllc.com -> 172.67.199.21, 104.21.21.135   (Cloudflare proxy)
MX   route1/2/3.mx.cloudflare.net                        (Email Routing)
TXT  v=spf1 include:_spf.mx.cloudflare.net ~all
```

That makes this easy. The zone is in the same Cloudflare account as the Pages
project, so **Pages creates the DNS record itself** — there is no external DNS
panel to touch, and mail is Cloudflare Email Routing rather than an outside
provider, so nothing here can endanger it.

> Note for anyone reading old notes: `logicloom.com` (no `llc`) is a *different*
> domain, sitting on Microsoft nameservers. It is not the target and should not
> be touched.

**Plan: `stemwave.logicloomllc.com`.**

Path routing (`logicloomllc.com/stemwave`) is technically possible now that the
zone is on Cloudflare, but it needs a Worker route sitting in front of the live
root site. That's more moving parts on something already working, for a survey
with a six-week life. The subdomain is zero-touch and deletes cleanly.

---

## 1. Create the D1 database

```bash
cd zimmer-stemwave-survey
npm install
npx wrangler login

npx wrangler d1 create stemwave-survey
```

Copy the `database_id` it prints into `wrangler.toml`, replacing
`REPLACE_WITH_DATABASE_ID`. Then create the tables:

```bash
npm run db:init:remote
```

Verify:

```bash
npx wrangler d1 execute stemwave-survey --remote \
  --command "SELECT name FROM sqlite_master WHERE type='table'"
```

You want `responses` and `leads`.

---

## 2. Run it locally first

```bash
npm run db:init:local
npm run dev
```

Open <http://localhost:8788/?src=counter>. **Use your phone's viewport**, not a
desktop window — this is a phone-only survey in practice.

Walk the whole thing once and confirm:

- [ ] Answering "No" to the first question skips the pain and pricing blocks
- [ ] The four price questions appear **before** the $2,809 reveal
- [ ] Entering the price boxes out of order warns you once, then lets you pass
- [ ] Choosing "Monthly payments" reveals the follow-up amount question
- [ ] The counter never counts *upward* (it may shrink from 21 to 9)
- [ ] After submitting, closing and reopening gives a **fresh** survey

Then check the row landed:

```bash
npx wrangler d1 execute stemwave-survey --local \
  --command "SELECT id, src, completed, vw_valid, path FROM responses"
```

---

## 3. Deploy to Pages

```bash
npx wrangler pages project create zimmer-stemwave-survey \
  --production-branch main
npm run deploy
```

Bind the database in the dashboard — **this is the step people forget**, and
without it every submission returns a 500:

> Workers & Pages → zimmer-stemwave-survey → Settings → Bindings →
> Add → D1 database → variable name `DB` → database `stemwave-survey`

Set it for **both** Production and Preview. Then set the export secret:

```bash
npx wrangler pages secret put EXPORT_KEY
# paste a long random string; keep a copy in your password manager
```

Generate one with:

```bash
openssl rand -base64 32
```

Redeploy after adding bindings (`npm run deploy`) — bindings only attach to
deployments created after them.

---

## 4. Custom domain

Because `logicloomllc.com` is in the same Cloudflare account, this is one
screen and no manual DNS:

> Workers & Pages → zimmer-stemwave-survey → Custom domains →
> Set up a custom domain → `stemwave.logicloomllc.com` → Activate domain

Cloudflare recognises the zone, **adds the CNAME for you**, and issues the
certificate automatically. Usually live in a couple of minutes.

Confirm:

```bash
dig +short stemwave.logicloomllc.com
curl -sSI https://stemwave.logicloomllc.com | head -3
```

Two things to leave alone while you're in the DNS tab: the root `A` records and
anything under `route*.mx.cloudflare.net` / `_spf.mx.cloudflare.net`. Those are
the live site and Email Routing. Adding a subdomain doesn't affect either — just
don't edit them by accident.

**Fallback:** ship `zimmer-stemwave-survey.pages.dev` as-is. It works
immediately. A branded domain reads as more legitimate on a printed card, so
prefer the subdomain, but don't let it block the pilot.

---

## 5. Generate the QR codes

Only after the URL is final and loading:

```bash
pip install segno
python3 tools/make_qr.py --base https://stemwave.logicloomllc.com
```

Writes SVG (print) and PNG (screen) per placement into `qr/`.

**Then test them physically.** Print the counter card at its real size, tape it
where it's going to live, and scan it with:

- [ ] an iPhone
- [ ] an Android phone
- [ ] the oldest phone you can borrow
- [ ] from where a patient actually stands, in the light that's actually there

Front-desk lighting is usually bad and often has glare from a window. A code
that scans on your desk can fail on a glossy laminated card under a downlight.
Matte stock, no lamination.

---

## 6. Before it goes live

- [ ] **Dr. Zimmer has signed off on the concept copy.** The description in
      `questions.js` (screen `concept`) is deliberately conservative — mechanism
      and logistics, no outcome claims. If he wants regenerative or stem-cell
      language added, that's his call to make knowingly: it's the kind of claim
      that draws FTC and state-board attention in chiropractic marketing.
- [ ] Confirm the pricing in `config.js` still matches what he's considering.
- [ ] Decide whether the contact-capture screen stays
      (`config.enableLeadCapture`). If the front desk would rather handle
      follow-up by hand, set it to `false`.
- [ ] Confirm a monthly payment plan is genuinely on the table
      (`config.enableMonthlyPlanQuestion`). If he'd never offer financing, turn
      it off — but that question is the highest-value one in the survey, so
      turn it off only if the answer really is never.
- [ ] Brief the front desk (script below).

---

## 7. Front desk script

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
3. **"About three minutes"** — and it's true, so it stays true.
4. **"You can finish it in the car"** — answers save on their phone. This
   rescues everyone who says "I'm in a rush."

Do **not** say it's about a specific new machine, and don't describe the
therapy — the survey does that in controlled wording, and an off-the-cuff
description in the lobby is exactly the claims risk we designed around.

---

## 8. Getting the data out

```bash
curl -sS "https://stemwave.logicloomllc.com/api/export?key=YOUR_EXPORT_KEY" \
  -o responses.csv

curl -sS "https://stemwave.logicloomllc.com/api/export?key=YOUR_EXPORT_KEY&table=leads" \
  -o leads.csv
```

Wrong or missing key returns a plain 404, so the endpoint doesn't advertise
itself. Then:

```bash
python3 tools/analyze.py responses.csv
```

Prints the price curves, the intersections, where $2,760 falls, Juster intent,
and the key segment. Writes `price-sensitivity.png` if matplotlib is installed.

**Don't analyse before ~50 responses.** Van Westendorp curves need the sample;
below 30 the script says so and you should believe it.

---

## 9. Shutting it down

It's a temporary survey. When it's done:

1. Export both tables and save them somewhere durable.
2. Pull the printed QR codes off the counter.
3. Delete the Pages project and the D1 database:
   ```bash
   npx wrangler d1 delete stemwave-survey
   ```
4. Remove the `stemwave` CNAME from Microsoft DNS.

Step 1 first. The leads table has real people's phone numbers in it and it is
the only place they exist.

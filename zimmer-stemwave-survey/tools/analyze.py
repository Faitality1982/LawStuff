#!/usr/bin/env python3
"""
Analyse the survey export.

    python3 tools/analyze.py stemwave-responses-2026-08-20.csv

Produces the numbers that actually drive the decision:
  * Van Westendorp price curves -> PMC, PME, IPP, OPP, and where $2,760 falls
  * Juster purchase probability -> mean probability and top-box
  * The segment crosstab that matters: tried several things, still unhappy

Standard library only. If matplotlib happens to be installed it also writes
a price-curve chart; if not, everything else still runs.
"""

import argparse
import csv
import pathlib
import sys
from collections import Counter

PACKAGE_PRICE = 2760

# Minimum usable time on survey. Anything faster is straight-lining, not reading.
MIN_DURATION_S = 45

# Van Westendorp needs a real sample before the curves mean anything.
VW_MIN_N = 30
VW_GOOD_N = 50


# --------------------------------------------------------------------- helpers
def num(v):
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").replace("$", ""))
    except ValueError:
        return None


def pct(n, d):
    return 0.0 if not d else 100.0 * n / d


def bar(fraction, width=28):
    filled = int(round(fraction * width))
    return "#" * filled + "." * (width - filled)


def rule(title):
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


# ------------------------------------------------------------- van westendorp
def curves(rows, grid):
    """Return the four cumulative curves over `grid`.

    too_cheap / bargain are DESCENDING (share who consider it at least that
    cheap); expensive / too_expensive are ASCENDING.
    """
    n = len(rows)
    out = {"too_cheap": [], "bargain": [], "expensive": [], "too_expensive": []}
    for p in grid:
        out["too_cheap"].append(sum(1 for r in rows if r["cheap"] >= p) / n)
        out["bargain"].append(sum(1 for r in rows if r["bargain"] >= p) / n)
        out["expensive"].append(sum(1 for r in rows if r["expensive"] <= p) / n)
        out["too_expensive"].append(sum(1 for r in rows if r["tooexp"] <= p) / n)
    return out


def intersect(grid, a, b):
    """First crossing of two curves, linearly interpolated."""
    for i in range(1, len(grid)):
        d0, d1 = a[i - 1] - b[i - 1], a[i] - b[i]
        if d0 == 0:
            return grid[i - 1]
        if d0 * d1 < 0:
            t = d0 / (d0 - d1)
            return grid[i - 1] + t * (grid[i] - grid[i - 1])
    return None


def van_westendorp(rows):
    vw = []
    for r in rows:
        vals = {
            "cheap": num(r.get("vw_cheap")),
            "bargain": num(r.get("vw_bargain")),
            "expensive": num(r.get("vw_expensive")),
            "tooexp": num(r.get("vw_tooexp")),
        }
        if any(v is None for v in vals.values()):
            continue
        if r.get("vw_valid") == "0":
            continue
        vw.append(vals)

    rule("VAN WESTENDORP PRICE SENSITIVITY")
    print(f"Usable responses: {len(vw)}")

    if len(vw) < VW_MIN_N:
        print(f"\n  Below {VW_MIN_N} usable responses. Not enough to draw the curves.")
        print("  Report these as directional only — do not set a price off this.")
        if not vw:
            return None
    elif len(vw) < VW_GOOD_N:
        print(f"  Between {VW_MIN_N} and {VW_GOOD_N} — readable, but thin. Treat as provisional.")

    lo = min(r["cheap"] for r in vw)
    hi = max(r["tooexp"] for r in vw)
    step = max(25.0, (hi - lo) / 400.0)
    grid = []
    p = lo
    while p <= hi:
        grid.append(p)
        p += step

    c = curves(vw, grid)
    pts = {
        "PMC": intersect(grid, c["too_cheap"], c["expensive"]),
        "PME": intersect(grid, c["too_expensive"], c["bargain"]),
        "IPP": intersect(grid, c["bargain"], c["expensive"]),
        "OPP": intersect(grid, c["too_cheap"], c["too_expensive"]),
    }

    labels = {
        "PMC": "Point of Marginal Cheapness  (floor)",
        "OPP": "Optimal Price Point          (resistance minimised)",
        "IPP": "Indifference Price Point     (typical / market norm)",
        "PME": "Point of Marginal Expensiveness (ceiling)",
    }
    print()
    for k in ("PMC", "OPP", "IPP", "PME"):
        v = pts[k]
        print(f"  {labels[k]:<52} {'$' + format(round(v), ',') if v else 'not found'}")

    if pts["PMC"] and pts["PME"]:
        print(f"\n  Acceptable price range: ${round(pts['PMC']):,} to ${round(pts['PME']):,}")

    rule(f"WHERE ${PACKAGE_PRICE:,} FALLS")
    pme, pmc = pts["PME"], pts["PMC"]
    if pme and PACKAGE_PRICE > pme:
        over = 100 * (PACKAGE_PRICE - pme) / pme
        print(f"  ABOVE the ceiling by {over:.0f}%.")
        print("  The price itself is the barrier. Lowering it or breaking it into")
        print("  smaller committed units is the lever — messaging will not fix this.")
    elif pmc and PACKAGE_PRICE < pmc:
        print("  BELOW the floor. Patients would suspect it doesn't work.")
        print("  There is room to raise the price.")
    elif pme and pmc:
        posn = 100 * (PACKAGE_PRICE - pmc) / (pme - pmc)
        print(f"  INSIDE the acceptable range, {posn:.0f}% of the way from floor to ceiling.")
        print("  Price is not the blocker. If intent is still low, the problem is")
        print("  explanation or financing — a completely different fix.")

    tooexp_at = sum(1 for r in vw if r["tooexp"] <= PACKAGE_PRICE) / len(vw)
    print(f"\n  {pct(sum(1 for r in vw if r['tooexp'] <= PACKAGE_PRICE), len(vw)):.0f}% "
          f"say ${PACKAGE_PRICE:,} is already past what they'd consider.")

    try:
        # Clip the x-axis so the interesting region isn't squashed by a
        # handful of respondents who typed a very large "too expensive".
        tooexps = sorted(r["tooexp"] for r in vw)
        p90 = tooexps[int(0.90 * (len(tooexps) - 1))]
        xmax = max(p90, PACKAGE_PRICE * 1.35)
        plot(grid, c, pts, xmax)
    except Exception:
        pass

    return pts


def plot(grid, c, pts, xmax):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(9, 5.5))
    series = [
        ("too_cheap", "Too cheap", "#7a9bbd"),
        ("bargain", "Good value", "#2f7d3a"),
        ("expensive", "Getting expensive", "#c9812c"),
        ("too_expensive", "Too expensive", "#a3352a"),
    ]
    for key, label, colour in series:
        ax.plot(grid, [v * 100 for v in c[key]], label=label, color=colour, lw=2)

    # Stagger the labels vertically — PMC/OPP and IPP/PME routinely land within
    # a few dollars of each other and would otherwise print on top of one another.
    for row, (k, v) in enumerate(sorted(
            ((k, v) for k, v in pts.items() if v), key=lambda kv: kv[1])):
        ax.axvline(v, color="#9aa7b2", ls=":", lw=1)
        ax.text(v, 101 + (row % 2) * 4.5, f"{k}\n${round(v):,}",
                ha="center", va="bottom", fontsize=7.5, color="#4a5a6a",
                linespacing=1.15)

    ax.axvline(PACKAGE_PRICE, color="#16202b", ls="--", lw=1.6)
    ax.text(PACKAGE_PRICE, 52, f"  proposed\n  ${PACKAGE_PRICE:,}",
            fontsize=9.5, color="#16202b", fontweight="bold")

    ax.set_xlabel("Price for the full course (USD)")
    ax.set_ylabel("Share of respondents (%)")
    ax.set_title("Stem Wave therapy — price sensitivity", pad=26)
    ax.set_xlim(0, xmax)
    ax.set_ylim(0, 118)
    ax.grid(alpha=0.15)
    ax.legend(loc="center right", fontsize=9, frameon=False)
    fig.tight_layout()
    fig.savefig("price-sensitivity.png", dpi=160)
    print("\n  Chart written to price-sensitivity.png")


# -------------------------------------------------------------------- juster
def juster(rows, field, label):
    vals = [num(r.get(field)) for r in rows]
    vals = [v for v in vals if v is not None]
    if not vals:
        return
    n = len(vals)
    top = sum(1 for v in vals if v >= 9)
    soft = sum(1 for v in vals if 7 <= v <= 8)
    mean_p = sum(vals) / n / 10.0

    print(f"\n  {label}   (n = {n})")
    print(f"    Mean probability      {mean_p * 100:5.1f}%   {bar(mean_p)}")
    print(f"    Top box (9-10)        {pct(top, n):5.1f}%   {bar(top / n)}   <- real intent")
    print(f"    Soft (7-8)            {pct(soft, n):5.1f}%")


# ------------------------------------------------------------------- crosstab
def counts(rows, field, title, order=None):
    c = Counter(r.get(field, "") for r in rows if r.get(field))
    if not c:
        return
    total = sum(c.values())
    print(f"\n  {title}")
    keys = order if order else [k for k, _ in c.most_common()]
    for k in keys:
        if k not in c:
            continue
        print(f"    {k:<26} {c[k]:>4}  {pct(c[k], total):5.1f}%  {bar(c[k] / total, 20)}")


def multi_counts(rows, field, title):
    c = Counter()
    n = 0
    for r in rows:
        v = r.get(field, "")
        if not v:
            continue
        n += 1
        for part in v.split("|"):
            if part:
                c[part] += 1
    if not n:
        return
    print(f"\n  {title}   (n = {n}, multi-select)")
    for k, v in c.most_common():
        print(f"    {k:<26} {v:>4}  {pct(v, n):5.1f}%  {bar(v / n, 20)}")


# ----------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Analyse the Stem Wave survey export.")
    ap.add_argument("csvfile", type=pathlib.Path)
    ap.add_argument("--keep-fast", action="store_true",
                    help="Keep responses faster than %d seconds" % MIN_DURATION_S)
    args = ap.parse_args()

    if not args.csvfile.exists():
        sys.exit(f"No such file: {args.csvfile}")

    with args.csvfile.open(newline="", encoding="utf-8-sig") as fh:
        raw = list(csv.DictReader(fh))

    rule("DATA HYGIENE")
    print(f"  Rows in file                  {len(raw)}")

    rows = [r for r in raw if r.get("completed") == "1"]
    print(f"  Completed                     {len(rows)}  ({pct(len(rows), len(raw)):.0f}%)")

    if not args.keep_fast:
        before = len(rows)
        rows = [r for r in rows if (num(r.get("duration_s")) or 999) >= MIN_DURATION_S]
        if before - len(rows):
            print(f"  Dropped as straight-lining    {before - len(rows)}  "
                  f"(under {MIN_DURATION_S}s)")

    flagged = sum(1 for r in rows if r.get("vw_valid") == "0")
    if flagged:
        print(f"  Non-monotonic price answers   {flagged}  (excluded from curves only)")

    print(f"  Analysed                      {len(rows)}")

    if not rows:
        sys.exit("\nNothing to analyse yet.")

    counts(rows, "src", "Responses by placement")

    rule("WHO ANSWERED")
    counts(rows, "a1", "Current pain",
           ["most_days", "on_off", "past_year", "none"])
    counts(rows, "a3", "How long",
           ["lt1m", "1_6m", "6_12m", "1_3y", "gt3y"])
    multi_counts(rows, "a2", "Where")
    counts(rows, "f3", "Patient status", ["current", "former", "no"])
    counts(rows, "f4", "Age", ["18_34", "35_49", "50_64", "65p", "decline"])

    rule("UNMET NEED")
    multi_counts(rows, "b1", "Already tried")
    counts(rows, "b2", "Satisfaction with results",
           ["very_sat", "somewhat_sat", "neutral", "somewhat_dis", "very_dis"])
    counts(rows, "b3", "Surgery discussed",
           ["had_it", "recommended", "mentioned", "no", "unsure"])

    rule("CONCEPT REACTION")
    counts(rows, "c2", "Appeal",
           ["very", "somewhat", "neutral", "not_very", "not_at_all"])
    multi_counts(rows, "c4", "Biggest hesitation")

    van_westendorp(rows)

    rule("PURCHASE INTENT (JUSTER)")
    print("  Stated intent overstates real behaviour. Top box (9-10) is the number")
    print("  to plan against; treat 7-8 as soft and 0-6 as no.")
    juster(rows, "e1", "Book the $49 Discovery Visit")
    juster(rows, "e2", "Buy the full course after a good Discovery Visit")

    rule("PAYMENT STRUCTURE")
    counts(rows, "e3", "Preferred structure",
           ["full", "two", "per_visit", "monthly", "none"])
    monthly = [num(r.get("e4_monthly")) for r in rows if r.get("e4_monthly")]
    monthly = [m for m in monthly if m]
    if monthly:
        monthly.sort()
        med = monthly[len(monthly) // 2]
        print(f"\n  Acceptable monthly payment (n = {len(monthly)})")
        print(f"    median ${med:,.0f}   mean ${sum(monthly) / len(monthly):,.0f}   "
              f"range ${monthly[0]:,.0f}-${monthly[-1]:,.0f}")
        print(f"    12 months at the median = ${med * 12:,.0f}")
    counts(rows, "e5", "HSA / FSA would matter",
           ["significant", "somewhat", "no_diff", "no_account"])

    # The segment that actually converts on a four-figure cash package.
    rule("KEY SEGMENT: tried 3+ things, still dissatisfied")
    seg = [
        r for r in rows
        if len([x for x in (r.get("b1") or "").split("|") if x and x != "nothing"]) >= 3
        and r.get("b2") in ("somewhat_dis", "very_dis")
    ]
    print(f"  Size: {len(seg)} of {len(rows)}  ({pct(len(seg), len(rows)):.0f}%)")
    if seg:
        juster(seg, "e1", "Discovery Visit — this segment")
        juster(seg, "e2", "Full course — this segment")
        print("\n  Compare against the all-respondent numbers above. If this segment")
        print("  is materially higher, that is the group to market to, and the")
        print("  average understates the real opportunity.")
    else:
        print("  Nobody yet. Worth re-checking once the sample grows.")

    open_text = [r.get("f2", "").strip() for r in rows if r.get("f2", "").strip()]
    if open_text:
        rule(f"OPEN COMMENTS ({len(open_text)})")
        for t in open_text:
            print(f"  - {t}")

    print()


if __name__ == "__main__":
    main()

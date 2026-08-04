#!/usr/bin/env python3
"""
Generate a synthetic export CSV so analyze.py can be exercised before any real
responses exist.

    python3 tools/make_fixture.py --n 120 --out fixture.csv
    python3 tools/analyze.py fixture.csv

This is TEST DATA. It is invented. Never mix it with a real export, and never
show its output to Dr. Zimmer as a finding.
"""

import argparse
import csv
import random
import uuid
from datetime import datetime, timedelta, timezone

COLUMNS = [
    "id", "created_at", "src", "duration_ms", "duration_s", "completed",
    "vw_valid", "path", "a1", "a2", "a3", "a4", "b1", "b2", "b3",
    "c1", "c2", "c4", "vw_cheap", "vw_bargain", "vw_expensive", "vw_tooexp",
    "e1", "e2", "e3", "e4_monthly", "e5", "f3", "f4", "f2",
    "ua_mobile", "screen_w",
]

SITES = ["low_back", "neck", "shoulder", "knee", "hip", "foot", "elbow",
         "hand_wrist", "neuropathy", "other"]
TRIED = ["chiro", "pt", "massage", "injections", "rx", "otc", "surgery",
         "needling", "laser"]
HESITATION = ["cost", "insurance", "efficacy", "time", "pain", "understanding", "evidence"]

COMMENTS = [
    "I'd want to know how many people it actually helps.",
    "Cost is the only thing stopping me.",
    "Would do it tomorrow if insurance covered any of it.",
    "My knee has been bad for years, willing to try anything.",
    "Need to talk to my wife about the money first.",
]


def one(rng, when):
    a1 = rng.choices(
        ["most_days", "on_off", "past_year", "none"],
        weights=[34, 38, 16, 12],
    )[0]
    has_pain = a1 != "none"

    row = {c: "" for c in COLUMNS}
    row["id"] = str(uuid.uuid4())
    row["created_at"] = when.isoformat()
    row["src"] = rng.choices(
        ["counter", "room1", "room2", "email", "card"],
        weights=[58, 12, 10, 15, 5],
    )[0]
    row["completed"] = "1"
    row["vw_valid"] = "1"
    row["path"] = "full" if has_pain else "short"
    row["a1"] = a1
    row["ua_mobile"] = "1"
    row["screen_w"] = str(rng.choice([390, 393, 412, 428]))
    row["c1"] = rng.choices(["had_elsewhere", "heard", "new"], weights=[6, 30, 64])[0]
    row["c2"] = rng.choices(
        ["very", "somewhat", "neutral", "not_very", "not_at_all"],
        weights=[20, 34, 24, 14, 8],
    )[0]
    row["c4"] = "|".join(rng.sample(HESITATION, rng.choice([1, 2])))
    row["f3"] = rng.choices(["current", "former", "no"], weights=[78, 14, 8])[0]
    row["f4"] = rng.choices(["18_34", "35_49", "50_64", "65p", "decline"],
                            weights=[12, 27, 34, 22, 5])[0]

    if not has_pain:
        secs = rng.randint(50, 130)
        row["duration_ms"] = str(secs * 1000)
        row["duration_s"] = str(secs)
        return row

    n_sites = rng.choices([1, 2, 3], weights=[58, 32, 10])[0]
    row["a2"] = "|".join(rng.sample(SITES, n_sites))
    row["a3"] = rng.choices(["lt1m", "1_6m", "6_12m", "1_3y", "gt3y"],
                            weights=[6, 18, 17, 27, 32])[0]
    row["a4"] = str(rng.randint(2, 9))

    n_tried = rng.choices([0, 1, 2, 3, 4, 5], weights=[6, 17, 24, 25, 18, 10])[0]
    row["b1"] = "|".join(rng.sample(TRIED, n_tried)) if n_tried else "nothing"
    if n_tried:
        row["b2"] = rng.choices(
            ["very_sat", "somewhat_sat", "neutral", "somewhat_dis", "very_dis"],
            # more things tried without relief -> more dissatisfaction
            weights=[10, 24, 20, 28 + 4 * n_tried, 12 + 3 * n_tried],
        )[0]
    row["b3"] = rng.choices(["had_it", "recommended", "mentioned", "no", "unsure"],
                            weights=[8, 15, 21, 46, 10])[0]

    # Van Westendorp. Anchored on a latent willingness-to-pay that rises with
    # chronicity and dissatisfaction, so the segment crosstab has real signal.
    base = rng.gauss(1150, 480)
    if row["a3"] in ("1_3y", "gt3y"):
        base *= 1.25
    if row["b2"] in ("somewhat_dis", "very_dis"):
        base *= 1.30
    if row["b3"] == "recommended":
        base *= 1.35
    base = max(180, base)

    cheap = max(40, base * rng.uniform(0.16, 0.34))
    bargain = base * rng.uniform(0.60, 0.92)
    expensive = base * rng.uniform(1.15, 1.65)
    tooexp = expensive * rng.uniform(1.25, 2.30)

    vals = [cheap, bargain, expensive, tooexp]
    # ~7% of real respondents answer these out of order; the app flags them.
    if rng.random() < 0.07:
        rng.shuffle(vals)
        row["vw_valid"] = "0"

    for key, v in zip(["vw_cheap", "vw_bargain", "vw_expensive", "vw_tooexp"], vals):
        row[key] = str(int(round(v / 5.0) * 5))

    afford = min(1.0, tooexp / 2809.0)
    appeal = {"very": 1.0, "somewhat": 0.72, "neutral": 0.42,
              "not_very": 0.18, "not_at_all": 0.04}[row["c2"]]
    p_disc = max(0.0, min(1.0, 0.55 * appeal + 0.45 * afford)) * rng.uniform(0.7, 1.15)
    row["e1"] = str(max(0, min(10, int(round(p_disc * 10)))))
    row["e2"] = str(max(0, min(10, int(round(p_disc * 10 * rng.uniform(0.45, 0.85))))))

    row["e3"] = rng.choices(
        ["full", "two", "per_visit", "monthly", "none"],
        weights=[9, 13, 22, 38, 18],
    )[0]
    if row["e3"] == "monthly":
        row["e4_monthly"] = str(int(round(rng.gauss(215, 85) / 5) * 5))
    row["e5"] = rng.choices(["significant", "somewhat", "no_diff", "no_account"],
                            weights=[26, 24, 20, 30])[0]

    if rng.random() < 0.11:
        row["f2"] = rng.choice(COMMENTS)

    secs = int(max(20, rng.gauss(178, 62)))
    row["duration_ms"] = str(secs * 1000)
    row["duration_s"] = str(secs)
    return row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=120)
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--out", default="fixture.csv")
    args = ap.parse_args()

    rng = random.Random(args.seed)
    start = datetime.now(timezone.utc) - timedelta(days=21)

    rows = []
    for i in range(args.n):
        when = start + timedelta(minutes=rng.randint(0, 21 * 24 * 60))
        r = one(rng, when)
        # a few abandoned partials, as in real life
        if rng.random() < 0.09:
            r["completed"] = "0"
        rows.append(r)

    with open(args.out, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(rows)

    print(f"Wrote {len(rows)} synthetic rows to {args.out}")
    print("TEST DATA ONLY — do not present this as findings.")


if __name__ == "__main__":
    main()

"""Fetch the photo-slide images and record where each one came from.

Every image on a photo slide has to carry a credit, so this does both jobs at
once: it downloads the file and it writes `photos.json`, which `build.js` reads
to print the `Credit:` line. The manifest is the provenance record for the
deck - an image in `photos/` that is not in the manifest never reaches a slide.

Two kinds of source, and the difference matters:

  commons  A file on Wikimedia Commons, named by its exact title. Licence and
           author come back from the API, so the credit is machine-read and
           cannot drift. Exact titles, never a search: Commons holds three
           other Rowland Hazards and two other Carl Jungs, and picking off a
           search result is how the wrong man ends up on the screen.

  cited    An early-A.A. archival photograph on a page Chris Zimmer sent.
           These are not licensed for reuse by the site hosting them and the
           rights almost certainly sit with A.A. archives. They are here
           because they are the pictures he picked, identified by the caption
           printed beside them. `rights` records that plainly rather than
           implying a licence nobody granted.

    python3 fetch_photos.py            # everything
    python3 fetch_photos.py carl-jung  # one slide
"""
import io, json, os, re, subprocess, sys, time, urllib.parse

UA  = "AA-BigBook-StudyDeck/1.0 (A.A. District 23 in-group class use)"
API = "https://commons.wikimedia.org/w/api.php"
OUT = "photos"
MANIFEST = "photos.json"
# A photo is drawn at most 4.5 in tall (see PHOTO_BAND in build.js). A 1080p
# projector over a 13.33 in slide is about 144 px/in, so ~650 px is already
# pixel-for-pixel; 1000 leaves room for a sharper room without carrying weight
# the screen can never show. The combined deck has to stay under 25 MB, and
# these are the only images in it that are not book pages.
MAXPX = 1000

ARCHIVAL = ("Early A.A. archival photograph. Rights not cleared - believed to "
            "rest with A.A. archives. In-group display only; not for reuse.")

MIDSURREY = "https://aamidsurrey.org.uk/wp-content/uploads/"

PICKS = {
  # --- Wikimedia Commons: public domain or Creative Commons -----------------
  # Commons throttles this container's egress IP hard - a bare 429, for
  # minutes at a stretch. curl() backs off and gets there; a failed slug just
  # needs the script run again for that slug.
  "carl-jung": {
    "commons": "File:ETH-BIB-Jung, Carl Gustav (1875-1961)-Portrait-Portr 14163 (cropped).tif"},
  "thetcher-tombstone": {
    "commons": "File:Thomas Thetcher Gravestone 2014-03-05 20-12 (retouched).jpg"},

  # william-james is the one frame with no working source at all. The Library
  # of Congress "[William James, half-length portrait]" (item 2016816573) is a
  # trap: it is a Hogarth painting of an 18th-century namesake, not the
  # psychologist. This is the right file, and it needs a run from a machine
  # Commons will serve.
  "william-james": {
    "commons": "File:William James by Alice M. Boughton, c. 1907, platinum print, from the National Portrait Gallery - NPG-NPG 87 37James-000001 (cropped).jpg"},

  # Fallbacks, if Commons ever stops serving these two. Both are uncleared and
  # smaller, so they are written down rather than wired up:
  #   thetcher-tombstone  royalhampshireregiment.org/.../2016/10/Thomas-Thetcher.jpg  (400x600)
  #   carl-jung           aamidsurrey.org.uk/.../2017/11/DrCarlJung.jpg               (225x309)

  # --- Library of Congress: no known restrictions --------------------------
  "oxford-group": {
    "url": "https://tile.loc.gov/storage-services/service/pnp/hec/21600/21634v.jpg",
    "page": "https://www.loc.gov/item/2016862679/",
    "source": "library-of-congress",
    "caption": "Rev. Frank N. D. Buchman, founder of the Oxford Group, ca. 1940",
    "author": "Harris & Ewing",
    "license": "No known restrictions",
    "rights": ("Library of Congress, Harris & Ewing photograph collection. "
               "No known restrictions on publication."),
    "credit": ("Rev. Frank N. D. Buchman, ca. 1940. Harris & Ewing, Library of "
               "Congress Prints and Photographs Division (LC item 2016862679). "
               "No known restrictions on publication.")},

  # --- Flickr, Creative Commons --------------------------------------------
  # The building as it stands today, not a period photograph - the caption on
  # the slide says so, so nobody reads it as 1934.
  "towns-hospital": {
    "url": "https://live.staticflickr.com/65535/53714606916_fc7b0c99ce_b.jpg",
    "page": "https://www.flickr.com/photos/edenpictures/53714606916",
    "source": "flickr-cc",
    "caption": "293 Central Park West, the former Charles B. Towns Hospital, in 2024",
    "author": "Eden, Janine and Jim",
    "license": "CC BY 2.0",
    "rights": "Creative Commons Attribution 2.0 Generic. Reusable with credit.",
    "credit": ("293 Central Park West, the former Charles B. Towns Hospital, "
               "photographed in 2024. Eden, Janine and Jim, via Flickr, CC BY 2.0.")},

  # --- Zimmer's own sources: A.A. archival photographs ----------------------
  "bill-w": {
    "url": MIDSURREY + "2017/10/BillW.jpg",
    "page": "https://aamidsurrey.org.uk/history-of-aa/bill-w/",
    "caption": "William Griffith Wilson, 1895-1971"},
  "dr-bob": {
    "url": MIDSURREY + "2018/02/Dr-Robert.jpg",
    "page": "https://aamidsurrey.org.uk/dr-robert-holbrook-smith/",
    "caption": "Robert Holbrook Smith, M.D., 1879-1950"},
  "silkworth": {
    "url": MIDSURREY + "2017/10/DrSilkworth.jpg",
    "page": "https://aamidsurrey.org.uk/history-of-aa/bill-w/",
    "caption": "William Duncan Silkworth, M.D., 1873-1951"},
  "ebby-thatcher": {
    "url": MIDSURREY + "2017/11/Ebby.jpg",
    "page": "https://aamidsurrey.org.uk/aa-history/edwin-thacher-ebby/",
    "caption": "Edwin Throckmorton Thacher, 1896-1966"},
  "leonard-strong": {
    "url": "https://aalkies.wordpress.com/wp-content/uploads/2013/08/dr-strong.jpg",
    "page": "https://aalkies.wordpress.com/aa-history-in-pictures/",
    "caption": "Dr. Leonard V. Strong, Jr., 1899-1989"},
  "rowland-hazard": {
    "url": MIDSURREY + "2017/11/Rowland-Hazard.jpg",
    "page": "https://aamidsurrey.org.uk/aa-history/rowland-hazard/",
    "caption": "Rowland Hazard III, 1881-1945"},
  "clarence-snyder": {
    "url": MIDSURREY + "2018/02/clarynce.jpg",
    "page": "https://aamidsurrey.org.uk/aa-history/key-people/",
    "caption": "Clarence H. Snyder, 1902-1984"},
  "aa-number-three": {
    "url": MIDSURREY + "2018/02/Bil-Dotson.jpg",
    "page": "https://aamidsurrey.org.uk/aa-history/key-people/",
    "caption": "Bill D. (Dotson), A.A. Number Three, 1879-1954"},
  "hank-parkhurst": {
    "url": MIDSURREY + "2018/02/Henry-Hank.jpg",
    "page": "https://aamidsurrey.org.uk/aa-history/key-people/",
    "caption": "Henry G. Parkhurst, 1895-1954"},
}


def curl(url, binary=False, tries=5):
    """Commons throttles cloud IPs hard, so back off rather than hammer."""
    delay = 10
    for _ in range(tries):
        r = subprocess.run(["curl", "-sSL", "--fail", "-A", UA, url],
                           capture_output=True, timeout=180)
        if r.returncode == 0:
            time.sleep(3)
            return r.stdout if binary else r.stdout.decode("utf-8", "replace")
        err = r.stderr.decode("utf-8", "replace")
        if "429" in err or "503" in err:
            time.sleep(delay); delay *= 2; continue
        raise RuntimeError("fetch failed (%s): %s" % (url, err[:200]))
    raise RuntimeError("rate limited after %d tries: %s" % (tries, url))


def strip_html(s):
    s = re.sub(r"<[^>]+>", " ", s or "")
    for a, b in (("&amp;", "&"), ("&nbsp;", " "), ("&quot;", '"'), ("&#039;", "'")):
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s).strip()


def commons_info(titles):
    out = {}
    for i in range(0, len(titles), 20):
        q = {"action": "query", "format": "json", "prop": "imageinfo",
             "iiprop": "url|size|extmetadata|mime",
             "titles": "|".join(titles[i:i + 20])}
        d = json.loads(curl(API + "?" + urllib.parse.urlencode(q)))
        for p in d["query"]["pages"].values():
            if "imageinfo" not in p:
                raise RuntimeError("no such Commons file: " + p.get("title", "?"))
            ii = p["imageinfo"][0]
            em = ii.get("extmetadata", {})
            g = lambda k: strip_html((em.get(k, {}) or {}).get("value", ""))
            author = g("Artist") or "Unknown"
            if author.lower().startswith(("unknown", "unbekannt")):
                author = "Unknown photographer"
            out[p["title"]] = {
                "url": ii["url"].split("?")[0],
                "author": author,
                "date": re.sub(r"\s*date QS:.*$", "", g("DateTimeOriginal")).strip(),
                "license": g("LicenseShortName") or "see source",
                "license_url": (em.get("LicenseUrl", {}) or {}).get("value", ""),
                "desc": g("ImageDescription")[:300],
                "page": "https://commons.wikimedia.org/wiki/"
                        + urllib.parse.quote(p["title"].replace(" ", "_")),
            }
    return out


def save(raw, slug):
    from PIL import Image
    im = Image.open(io.BytesIO(raw))
    if max(im.size) > MAXPX:
        f = MAXPX / max(im.size)
        im = im.resize((round(im.width * f), round(im.height * f)), Image.LANCZOS)
    name = slug + ".jpg"
    im.convert("RGB").save(os.path.join(OUT, name),
                           quality=88, optimize=True, progressive=True)
    return name, im.width, im.height


def main():
    # A bare run fetches what this container can actually reach. The Commons
    # entries are upgrades for a machine Commons will serve, so they are opt-in
    # by name rather than part of the default sweep.
    wanted = sys.argv[1:] or sorted(k for k, v in PICKS.items() if not v.get("upgrade"))
    missing = [s for s in wanted if s not in PICKS]
    if missing:
        raise SystemExit("no PICKS entry for: " + ", ".join(missing))

    os.makedirs(OUT, exist_ok=True)
    manifest = json.load(open(MANIFEST)) if os.path.exists(MANIFEST) else {}

    cw = [s for s in wanted if "commons" in PICKS[s]]
    info, failed = {}, []
    if cw:
        try:
            info = commons_info([PICKS[s]["commons"] for s in cw])
        except RuntimeError as e:
            print("  ! Commons unavailable: %s" % e)

    for slug in wanted:
        pick = PICKS[slug]
        try:
            if "commons" in pick:
                i = info.get(pick["commons"])
                if not i:
                    failed.append((slug, "Commons metadata unavailable")); continue
                name, w, h = save(curl(i["url"], binary=True), slug)
                rec = {"file": name, "px": [w, h], "source": "wikimedia-commons",
                       "source_title": pick["commons"], "source_url": i["page"],
                       "author": i["author"], "date": i["date"],
                       "license": i["license"], "license_url": i["license_url"],
                       "rights": "Reusable under the licence shown.",
                       "credit": "%s%s. Wikimedia Commons, %s." % (
                           i["author"],
                           ", " + i["date"] if i["date"] else "",
                           i["license"])}
            else:
                name, w, h = save(curl(pick["url"], binary=True), slug)
                rec = {"file": name, "px": [w, h],
                       "source": pick.get("source", "cited-by-zimmer"),
                       "source_title": pick["caption"], "source_url": pick["page"],
                       "author": pick.get("author", "Unattributed"), "date": "",
                       "license": pick.get("license", "Not cleared"),
                       "license_url": "",
                       "rights": pick.get("rights", ARCHIVAL),
                       "credit": pick.get("credit",
                           "%s. Early A.A. archival photograph, via %s. "
                           "Displayed for in-group study only." % (
                               pick["caption"],
                               urllib.parse.urlparse(pick["page"]).netloc))}
            manifest[slug] = rec
            print("  %-20s %-14s %4dx%-4d  %s"
                  % (slug, rec["license"], w, h, name))
        except (RuntimeError, OSError) as e:
            failed.append((slug, str(e)[:120]))

    json.dump(dict(sorted(manifest.items())), open(MANIFEST, "w"),
              indent=2, ensure_ascii=False)
    open(MANIFEST, "a").write("\n")
    print("wrote %s (%d entries)" % (MANIFEST, len(manifest)))
    for slug, why in failed:
        print("  ! %-20s %s" % (slug, why))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())

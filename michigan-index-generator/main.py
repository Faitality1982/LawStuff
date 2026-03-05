"""
Michigan Index of Authorities Generator
Uses docx2pdf (Microsoft Word) + pdfplumber with footer-region page detection.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading, os, sys, re, webbrowser, tempfile, shutil
from collections import defaultdict

try:
    import pdfplumber
    from docx2pdf import convert as docx2pdf_convert
except ImportError as e:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing dependency",
        f"{e}\n\nRun:  pip install pdfplumber docx2pdf")
    sys.exit(1)

# ── Regex patterns ────────────────────────────────────────────────────────────
REPORTER     = r'(?:Mich\.?\s*App\.?|Mich\.?|U\.?S\.?|NW[23]?d|S\.?\s*Ct\.?|SCt\.?)'
IN_RE_PAT    = re.compile(
    r'(?<![A-Za-z])(In\s+re\s+[A-Z][A-Za-z/]+(?:\s+[A-Z][A-Za-z]+)?(?:\s+Minors)?)'
    r'\s*,\s*(\d+\s+' + REPORTER + r'[^(]{0,150}\(\d{4}\))')
V_IN_RE_PAT  = re.compile(
    r'(?<![A-Za-z])([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){0,5}'
    r'\s+v[s]?\.?\s+[A-Z][A-Za-z]+(?:\s+[A-Z]?[A-Za-z]+){0,5}'
    r'\s*\(In\s+re\s+[A-Za-z/]+\))'
    r'\s*,\s*(\d+\s+' + REPORTER + r'[^(]{0,150}\(\d{4}\))')
SIMPLE_V_PAT = re.compile(
    r'(?<![A-Za-z])([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z\.]+){0,3}'
    r'\s+v[s]?\.?\s+[A-Z][A-Za-z]+(?:\s+[A-Z]?[A-Za-z\.]+){0,3})'
    r'\s*,\s*(\d+\s+' + REPORTER + r'[^(]{0,150}\(\d{4}\))')

# MCL: capture full subsection spec (no lookahead — we handle lists in code)
MCL_PAT      = re.compile(
    r'(MCL\s+\d+[A-Z]?\.\d+[A-Za-z]?'
    r'(?:\(\d+\))?(?:\([a-z]\))?(?:\([ivxIVX]+\))?)')
MCL_LIST_FOLLOW = re.compile(r'\s*,\s*\(')   # ", (" after match → list form

MCR_PAT      = re.compile(r'(MCR\s+\d+\.\d+(?:\([A-Za-z0-9]+\))*)')

# Pages to skip entirely
SKIP_PAT     = re.compile(
    r'TABLE OF CONTENTS|INDEX OF AUTHORITIES|CERTIFICATE OF COMP'
    r'|STATEMENT OF QUESTIONS INVOLVED',
    re.IGNORECASE)
# Pages that end the skip zone
BODY_PAT     = re.compile(
    r'STATEMENT OF FACTS|STATEMENT OF JURISDICTION', re.IGNORECASE)
# Argument section starts body (but don't use plain "ARGUMENT" alone — too easy to false-trigger)
ARG_PAT      = re.compile(r'^\s*ARGUMENT\s*$', re.MULTILINE)

MCL_SKIP     = re.compile(r'MCL\s+7\.\d')
ROMAN_RE     = re.compile(r'^(x{0,3}(?:ix|iv|v?i{0,3}))$', re.IGNORECASE)
ARABIC_RE    = re.compile(r'^\d+$')

CANON_NAMES = {
    'Family Independence Agency v Boursaw (In re Boursaw)':
        'Family Independent Agency v Boursaw (In re Boursaw)',
    'Family Independence Agency v Sours (In re Sours)':
        'Family Independent Agency v Sours (In re Sours)',
}

def clean_name(raw):
    raw = raw.strip().rstrip('.,;')
    raw = re.sub(r'^In\s+(?!re\s)', '', raw)
    return CANON_NAMES.get(raw, raw)

def sort_pages(pages):
    # Roman numeral values — covers i through xx (enough for any brief front matter)
    ROMAN_VALS = {
        'i':1,'ii':2,'iii':3,'iv':4,'v':5,'vi':6,'vii':7,'viii':8,'ix':9,
        'x':10,'xi':11,'xii':12,'xiii':13,'xiv':14,'xv':15,
        'xvi':16,'xvii':17,'xviii':18,'xix':19,'xx':20,
    }
    def key(p):
        pl = p.lower()
        rv = ROMAN_VALS.get(pl, 0)
        if rv:                return (0, rv)   # roman first, in order
        try:                  return (1, int(p)) # arabic second
        except ValueError:    return (2, p)
    return sorted(pages, key=key)

# ── Sorting keys per skill spec ───────────────────────────────────────────────

def case_sort_key(name):
    """Sort cases alphabetically by the meaningful part of the name.
    'In re Archer' → 'archer'
    'Family Independent Agency v Boursaw (In re Boursaw)' → 'boursaw'
    'Santosky v Kramer' → 'santosky v kramer'
    """
    # "X (In re Y)" → sort by Y
    m = re.search(r'\(In\s+re\s+([^)]+)\)', name, re.IGNORECASE)
    if m:
        return m.group(1).strip().lower()
    # "In re X" → sort by X
    m = re.match(r'In\s+re\s+(.*)', name, re.IGNORECASE)
    if m:
        return m.group(1).strip().lower()
    # everything else → full name lowercased
    return name.lower()

def _subsection_token_key(token):
    """Convert a subsection token like '(3)', '(c)', '(i)' to a numeric sort tuple."""
    inner = token.strip('()')
    # Roman numerals
    roman = {'i':1,'ii':2,'iii':3,'iv':4,'v':5,'vi':6,'vii':7,'viii':8,'ix':9,'x':10}
    if inner.lower() in roman:
        return (2, roman[inner.lower()])
    # Digits
    if inner.isdigit():
        return (0, int(inner))
    # Letters
    if inner.isalpha():
        return (1, ord(inner.lower()))
    return (3, inner)

def mcl_sort_key(name):
    """Sort MCL citations numerically: 712A.1 < 712A.19b < 712A.19b(3)(c)(i)."""
    # Strip "MCL " prefix
    rest = re.sub(r'^MCL\s+', '', name, flags=re.IGNORECASE).strip()
    # Split base section from subsections: "712A.19b" + ["(3)","(c)","(i)"]
    m = re.match(r'(\d+[Aa]?\.\d+[a-zA-Z]?)(.*)', rest)
    if not m:
        return (name,)
    base_str, subs_str = m.group(1), m.group(2)
    # Parse base: major.minor_alpha  e.g. "712A.19b" → (712, 'A', 19, 'b')
    bm = re.match(r'(\d+)([A-Za-z]?)\.(\d+)([A-Za-z]?)', base_str)
    if bm:
        base_key = (int(bm.group(1)), bm.group(2).upper(),
                    int(bm.group(3)), bm.group(4).lower())
    else:
        base_key = (0, '', 0, '')
    # Parse subsection tokens
    tokens = re.findall(r'\([^)]*\)', subs_str)
    sub_keys = tuple(_subsection_token_key(t) for t in tokens)
    return base_key + sub_keys

def mcr_sort_key(name):
    """Sort MCR citations numerically: 3.977 < 3.977(F) < 3.977(F)(1)(b) < 7.203."""
    rest = re.sub(r'^MCR\s+', '', name, flags=re.IGNORECASE).strip()
    m = re.match(r'(\d+)\.(\d+)(.*)', rest)
    if not m:
        return (name,)
    major, minor, subs_str = int(m.group(1)), int(m.group(2)), m.group(3)
    tokens = re.findall(r'\([^)]*\)', subs_str)
    sub_keys = tuple(_subsection_token_key(t) for t in tokens)
    return (major, minor) + sub_keys

# ── Get printed page number from PDF footer region ───────────────────────────
def get_page_label(pdf_page):
    h = float(pdf_page.height)
    w = float(pdf_page.width)

    # Crop bottom 10% where Word puts the footer page number
    footer = pdf_page.crop((0, h * 0.90, w, h))
    for line in (footer.extract_text() or '').split('\n'):
        token = line.strip()
        if ROMAN_RE.match(token):
            return token.lower()
        if ARABIC_RE.match(token) and 1 <= int(token) <= 999:
            return token

    # Fallback: bottom-up full-page scan
    for line in reversed([(l.strip()) for l in
                          (pdf_page.extract_text() or '').split('\n') if l.strip()]):
        if ROMAN_RE.match(line):
            return line.lower()
        if ARABIC_RE.match(line) and 1 <= int(line) <= 999:
            return line
    return None

# ── Core extraction ───────────────────────────────────────────────────────────
def extract_index(input_path, progress_cb=None):
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".pdf":
        # Native PDF — use directly
        pdf_path = input_path
        tmp_dir  = None
        if progress_cb: progress_cb("Scanning pages for citations…")
    else:
        # DOCX — convert via Microsoft Word first
        if progress_cb: progress_cb("Converting DOCX → PDF via Microsoft Word…")
        tmp_dir  = tempfile.mkdtemp()
        pdf_path = os.path.join(tmp_dir, "brief.pdf")
        try:
            docx2pdf_convert(input_path, pdf_path)
        except Exception as e:
            raise RuntimeError(
                f"Could not convert DOCX to PDF.\n"
                f"Make sure Microsoft Word is installed and not currently open.\n\n{e}")
        if not os.path.exists(pdf_path):
            raise RuntimeError("PDF was not created. Is Microsoft Word installed?")
        if progress_cb: progress_cb("Scanning pages for citations…")

    cases    = {}
    statutes = defaultdict(set)
    rules    = defaultdict(set)
    mcl_bare = defaultdict(set)

    # Track skip state: skip front-matter sections except jurisdiction
    in_skip   = False
    body_seen = False

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            if progress_cb and i % 5 == 0:
                progress_cb(f"Scanning page {i+1} of {total}…")

            raw   = page.extract_text() or ''
            label = get_page_label(page)
            if label is None:
                continue

            # --- Skip-zone management ---
            # Once we've seen body content, never skip again
            if BODY_PAT.search(raw) or ARG_PAT.search(raw):
                body_seen = True
                in_skip   = False

            if not body_seen:
                if SKIP_PAT.search(raw):
                    in_skip = True
                # Statement of Jurisdiction (page iii) contains MCR/MCL citations
                # we DO want — let it through
                if re.search(r'STATEMENT OF JURISDICTION', raw, re.IGNORECASE):
                    in_skip = False

            if in_skip:
                continue

            flat = re.sub(r'[\n\r]+', ' ', raw)
            flat = re.sub(r'\s+',    ' ', flat)

            # ── Cases ──
            seen = []
            for pat in [V_IN_RE_PAT, IN_RE_PAT, SIMPLE_V_PAT]:
                for m in pat.finditer(flat):
                    if not any(m.start() < e and m.end() > s for s, e in seen):
                        name = clean_name(m.group(1))
                        cite = m.group(2).strip().rstrip('.,;')
                        if name not in cases:
                            cases[name] = {'pages': set(), 'cite': cite}
                        cases[name]['pages'].add(label)
                        seen.append((m.start(), m.end()))

            # ── Statutes ──
            for m in MCL_PAT.finditer(flat):
                val = m.group(1).strip()
                if MCL_SKIP.match(val):
                    continue
                # List form: "MCL 712A.19b(3)(a), (b), ..." → record bare parent only
                is_list = bool(MCL_LIST_FOLLOW.match(flat[m.end():]))
                if is_list:
                    bare_m = re.match(r'(MCL\s+\d+[A-Z]?\.\d+[A-Za-z]?)', val)
                    if bare_m:
                        mcl_bare[bare_m.group(1).strip()].add(label)
                elif re.match(r'MCL\s+712A\.19b$', val):
                    mcl_bare[val].add(label)
                else:
                    statutes[val].add(label)

            # ── Court Rules ──
            for m in MCR_PAT.finditer(flat):
                rules[m.group(1).strip()].add(label)

    shutil.rmtree(tmp_dir, ignore_errors=True)

    # Collapse bare MCL prefix: merge its pages INTO subsections, keep it too
    mcl_keys = list(statutes.keys())
    filtered_statutes = {}
    for k in sorted(mcl_keys, key=len, reverse=True):
        is_prefix = any(o != k and o.startswith(k+'(') for o in mcl_keys)
        if not is_prefix:
            filtered_statutes[k] = statutes[k]
        else:
            for o in mcl_keys:
                if o != k and o.startswith(k+'('):
                    statutes[o].update(statutes[k])
    # Always keep the bare 712A.19b entry (it represents the statute chapter itself)
    for k, v in mcl_bare.items():
        filtered_statutes[k] = v

    return {
        'cases':    {k: {'pages': sort_pages(v['pages']), 'cite': v['cite']}
                     for k, v in sorted(cases.items(), key=lambda x: case_sort_key(x[0]))},
        'statutes': {k: sort_pages(v)
                     for k, v in sorted(filtered_statutes.items(), key=lambda x: mcl_sort_key(x[0]))},
        'rules':    {k: sort_pages(v)
                     for k, v in sorted(rules.items(), key=lambda x: mcr_sort_key(x[0]))},
    }

# ── HTML output ───────────────────────────────────────────────────────────────
def build_html(data):
    def ehtml(display_name, ps, italic=False):
        n = f'<em>{display_name}</em>' if italic else display_name
        return (f'<div class="entry">'
                f'<span class="entry-name">{n}</span>'
                f'<span class="dots"></span>'
                f'<span class="pages">{ps}</span>'
                f'</div>\n')
    def ejs(display_name, ps):
        safe = display_name.replace("'", "\\'")
        return f"    ['{safe}', '{ps}']"

    ch, cj = '', []
    for name, v in data['cases'].items():
        ps = ', '.join(v['pages'])
        # Full entry: "Name, Citation"
        display = f"{name}, {v['cite']}" if v['cite'] else name
        ch += ehtml(display, ps, italic=True)
        cj.append(ejs(display, ps))

    sh, sj = '', []
    for name, pages in data['statutes'].items():
        ps = ', '.join(pages)
        sh += ehtml(name, ps)
        sj.append(ejs(name, ps))

    rh, rj = '', []
    for name, pages in data['rules'].items():
        ps = ', '.join(pages)
        rh += ehtml(name, ps)
        rj.append(ejs(name, ps))

    cj_s = ',\n'.join(cj)
    sj_s = ',\n'.join(sj)
    rj_s = ',\n'.join(rj)

    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Index of Authorities</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:Arial,sans-serif;font-size:12pt;max-width:760px;margin:30px auto;padding:20px 30px 50px;color:#000;background:#fff}}
h1{{text-align:center;font-size:13pt;font-weight:bold;text-decoration:underline;margin-bottom:24px;letter-spacing:.04em}}
h2{{font-size:12pt;font-weight:normal;text-decoration:underline;margin:20px 0 8px}}
.entry{{display:flex;align-items:baseline;margin-bottom:5px;line-height:1.5}}
.entry-name{{flex-shrink:0;max-width:78%}}
.entry-name em{{font-style:italic}}
.dots{{flex:1;border-bottom:1px dotted #000;margin:0 4px 3px;min-width:12px}}
.pages{{flex-shrink:0;white-space:nowrap}}
.copy-btn{{display:block;margin:0 0 24px auto;padding:7px 18px;background:#1a3a6b;color:#fff;border:none;border-radius:4px;font-size:11pt;cursor:pointer;font-family:Arial,sans-serif}}
.copy-btn:hover{{background:#0f2548}}
.copy-btn.copied{{background:#2a7a3b}}
@media print{{.copy-btn{{display:none}}}}
</style></head><body>
<button class="copy-btn" id="copyBtn" onclick="copyIndex()">&#128203; Copy as Plain Text</button>
<h1>INDEX OF AUTHORITIES</h1>
<h2>CASE LAW</h2>
{ch}
<h2>STATUTES &amp; OTHER AUTHORITIES</h2>
{sh}
<h2>MICHIGAN COURT RULES</h2>
{rh}
<script>
const CASES=[{cj_s}];
const STATS=[{sj_s}];
const RULES=[{rj_s}];
function fmt(n,p){{return n+'.'.repeat(Math.max(3,70-n.length-p.length))+p;}}
function copyIndex(){{
  const L=['INDEX OF AUTHORITIES','','CASE LAW',''];
  CASES.forEach(([n,p])=>L.push(fmt(n,p)));
  L.push('','STATUTES & OTHER AUTHORITIES','');
  STATS.forEach(([n,p])=>L.push(fmt(n,p)));
  L.push('','MICHIGAN COURT RULES','');
  RULES.forEach(([n,p])=>L.push(fmt(n,p)));
  navigator.clipboard.writeText(L.join('\\n')).then(()=>{{
    const b=document.getElementById('copyBtn');
    b.textContent='✓ Copied!';b.classList.add('copied');
    setTimeout(()=>{{b.textContent='📋 Copy as Plain Text';b.classList.remove('copied');}},2500);
  }}).catch(()=>{{const t=document.createElement('textarea');t.value=L.join('\\n');
    document.body.appendChild(t);t.select();document.execCommand('copy');document.body.removeChild(t);alert('Copied!');}});
}}
</script></body></html>"""

# ── GUI ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Michigan Index of Authorities Generator")
        self.geometry("560x330")
        self.resizable(False, False)
        self.configure(bg="#f4f6fb")
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="Michigan Index of Authorities",
                 font=("Arial",15,"bold"), bg="#1a3a6b", fg="white", pady=14).pack(fill="x")
        frame = tk.Frame(self, bg="#f4f6fb", padx=30, pady=20)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Select your appellate brief (.docx or .pdf):",
                 font=("Arial",11), bg="#f4f6fb").pack(anchor="w")
        row = tk.Frame(frame, bg="#f4f6fb"); row.pack(fill="x", pady=8)
        self.path_var = tk.StringVar()
        tk.Entry(row, textvariable=self.path_var, font=("Arial",10), width=42).pack(side="left", padx=(0,8))
        tk.Button(row, text="Browse…", command=self._browse,
                  bg="#1a3a6b", fg="white", relief="flat", font=("Arial",10), padx=10).pack(side="left")
        self.status = tk.StringVar(value="Ready.")
        tk.Label(frame, textvariable=self.status, font=("Arial",9), fg="#555", bg="#f4f6fb").pack(anchor="w")
        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=480)
        self.progress.pack(pady=(6,12))
        tk.Button(frame, text="Generate Index of Authorities", command=self._run,
                  bg="#1a3a6b", fg="white", font=("Arial",12,"bold"), relief="flat", padx=20, pady=8).pack()

        # ── Donate button ──────────────────────────────────────────────────
        donate_frame = tk.Frame(self, bg="#f4f6fb", pady=6)
        donate_frame.pack(fill="x")
        donate_lbl = tk.Label(
            donate_frame,
            text="💛  Support this tool — Donate via PayPal",
            font=("Arial", 9, "underline"),
            fg="#0070ba", bg="#f4f6fb", cursor="hand2"
        )
        donate_lbl.pack()
        donate_lbl.bind("<Button-1>",
            lambda e: webbrowser.open(
                "https://www.paypal.com/donate/?hosted_button_id=4AM22J9AMNCEE"))

    def _browse(self):
        path = filedialog.askopenfilename(title="Select Brief",
            filetypes=[
                ("Brief files", "*.docx *.pdf"),
                ("Word Documents", "*.docx"),
                ("PDF Files", "*.pdf"),
                ("All files", "*.*"),
            ])
        if path: self.path_var.set(path)

    def _run(self):
        path = self.path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("No file", "Please select a valid .docx or .pdf file first.")
            return
        self.progress.start(12); self.status.set("Working…")
        threading.Thread(target=self._worker, args=(path,), daemon=True).start()

    def _worker(self, path):
        try:
            data = extract_index(path, progress_cb=lambda s: self.status.set(s))
            html = build_html(data)
            out  = tempfile.NamedTemporaryFile(suffix=".html", delete=False,
                       mode="w", encoding="utf-8", prefix="index_of_authorities_")
            out.write(html); out.close()
            self.after(0, lambda p=out.name, d=data: self._done(p, d))
        except Exception as ex:
            msg = str(ex)
            self.after(0, lambda m=msg: self._error(m))

    def _done(self, html_path, data):
        self.progress.stop()
        nc, ns, nr = len(data['cases']), len(data['statutes']), len(data['rules'])
        self.status.set(f"Done!  {nc} cases · {ns} statutes · {nr} rules")
        webbrowser.open(f"file:///{html_path}")
        messagebox.showinfo("Index Generated",
            f"Index of Authorities opened in your browser.\n\n"
            f"  Cases:    {nc}\n  Statutes: {ns}\n  Rules:    {nr}\n\n"
            f"Use the 'Copy as Plain Text' button to paste into Word.")

    def _error(self, msg):
        self.progress.stop(); self.status.set("Error — see dialog.")
        messagebox.showerror("Error", msg)

if __name__ == "__main__":
    App().mainloop()

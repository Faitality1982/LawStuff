# Michigan Index of Authorities Generator

A desktop tool that reads a Michigan appellate brief (`.docx`), extracts all case citations, statutes (MCL), and court rules (MCR) with their page numbers, and generates a formatted **Index of Authorities** — ready to copy into Word.

## Features
- Pick any `.docx` appellate brief
- OCR-based page number detection via `pdfplumber`
- Extracts *In re* cases, *X v Y* cases, MCL statutes, MCR court rules
- Skips Table of Contents / existing Index sections to avoid duplicates
- Opens result in browser with dot-leader formatting
- One-click **Copy as Plain Text** button for pasting into Word

## Running from Source

```bash
pip install pdfplumber python-docx
# Also requires LibreOffice installed (for DOCX → PDF conversion)
python main.py
```

## Building the .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name MichiganIndexGenerator main.py
# Output: dist/MichiganIndexGenerator.exe
```

## Requirements
- Python 3.8+
- LibreOffice (for DOCX → PDF conversion)
- `pdfplumber`, `python-docx`, `tkinter` (standard library)

## Usage
1. Run `MichiganIndexGenerator.exe`
2. Click **Browse…** and select your `.docx` brief
3. Click **Generate Index of Authorities**
4. The index opens in your browser
5. Click **Copy as Plain Text** → paste into your Word document

## Citation Patterns Supported
- `In re [Name], [volume] [reporter] [pages] ([year])`
- `[Agency] v [Party] (In re [Name]), ...`
- `[Name] v [Name], ...`
- `MCL ###.##(subsections)`
- `MCR ###.###(subsections)`

---
*Built for Michigan appellate practice — St. Clair County*

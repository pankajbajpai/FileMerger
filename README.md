# File Merger
Merging pdf/excel/word doc into single file

## Acknowledgement
1. Claude.ai is used heavily to design the code

## 
A self-contained Python package with a local web UI to merge or convert:

- **PDF**, **Excel** (`.xlsx`, `.xls`, `.csv`), and **Word** (`.docx`) files —
  in any mix — into **one output file**.
- The output format is decided entirely by the extension you type into the
  output file name (must be `.pdf`, `.docx`, or `.xlsx`).
  - If all selected files belong to the same family as the output (e.g. all
    Excel-family files → `.xlsx`), a high-fidelity native merge is used
    (each source file/sheet becomes its own sheet for Excel output).
  - If files are mixed types, or a single file's type differs from the
    output extension, each file is converted (best-effort: text and tables
    preserved, original layout/formatting is not) into the target format
    before merging.
  - A single file with the *same* extension as the output is just carried
    through untouched (effectively a rename/passthrough) — no lossy
    conversion happens unless the output extension actually differs.

Password-protected inputs are detected automatically — you're only asked for
a password for a file that's actually protected, never for the rest. Nothing
is uploaded anywhere — everything runs on `127.0.0.1` on your own machine.

## Install

### Windows / Linux / macOS
```bash
pip install .
```
(run from inside this folder), or without installing the package:
```bash
pip install -r requirements.txt
```

### Android (via Termux)
```bash
pkg install python
pip install .
```

## Run

If installed with `pip install .`:
```bash
filemerger
```

Or, without installing, from inside this folder:
```bash
pip install -r requirements.txt
python -m filemerger.cli
```

This starts a local server (default `http://127.0.0.1:5678`) and opens it in
your default browser automatically. On Android/Termux, or if the browser
doesn't open automatically, open that address manually in Chrome/Firefox.

Use `filemerger --no-browser` to start the server without auto-opening a browser.

## Usage

1. Add one or more files (any mix of PDF / Excel / Word).
2. Type an output file name **including its extension**, e.g. `combined.docx`.
3. Click **Merge & Download**.
   - If any file turns out to be password protected, the app will tell you
     exactly which one(s) and show a password box only for those — fill it
     in and click **Merge & Download** again.
4. Your browser downloads the result; use its "Save As" prompt to choose
   the destination folder.

## Notes & limitations

- Legacy `.doc` (old binary Word format) is **not** supported as input —
  only `.docx`. Convert `.doc` to `.docx` first (e.g. in Word or
  LibreOffice) if needed.
- Supported output extensions: `.pdf`, `.docx`, `.xlsx` only.
- `.xls`/`.xlsx`/`.docx` password protection is decrypted via
  `msoffcrypto-tool`; encrypted PDFs are decrypted via `pypdf`.
- Cross-format conversion (e.g. PDF → DOCX, DOCX → XLSX) is best-effort:
  text and tables carry over, but original layout, images, and formatting
  do not.
- Maximum total upload size per merge request is 512 MB (adjustable in
  `filemerger/app.py`).

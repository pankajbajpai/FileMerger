# File Merger
Merging pdf/excel/word doc into single file

## Acknowledgement
1. Claude.ai is used heavily to design the code

## -------------------------------------------------------------------------------
A self-contained Python package with a local web UI to merge:

- **PDF** files → single `.pdf`
- **Excel** files (`.xlsx`, `.xls`, `.csv`) → single `.xlsx` (each source file/sheet
  becomes its own sheet in the output workbook)
- **Word** files (`.docx`) → single `.docx`

Password-protected inputs are decrypted in memory using a password you supply,
then merged. No files are uploaded anywhere — everything runs on `127.0.0.1`
on your own machine.

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

1. Choose a file type (PDF / Excel / Word).
2. Add two or more files of that type.
3. Mark any password-protected file and enter its password.
4. Optionally set an output file name.
5. Click **Merge & Download** — your browser will download the merged file;
   use your browser's "Save As" prompt to choose the destination folder.

## Notes & limitations

- Legacy `.doc` (old binary Word format) is **not** supported — only `.docx`.
  Convert `.doc` to `.docx` first (e.g. in Word or LibreOffice) if needed.
- `.xls`/`.xlsx` password protection is decrypted via `msoffcrypto-tool`.
- Output for Excel-family merges is always `.xlsx` (the most capable format
  among xlsx/xls/csv).
- Output for Word-family merges is always `.docx`.
- Maximum total upload size per merge request is 512 MB (adjustable in
  `filemerger/app.py`).

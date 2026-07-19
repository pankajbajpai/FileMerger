"""
Generic cross-format conversion + merge.

Used whenever the input files are not all the same "family" as the requested
output extension (e.g. merging a .pdf and a .docx into one .docx, or merging
a single .xlsx into a .pdf). Each input file is reduced to a simple list of
content blocks (heading / paragraph / table), and then all files' blocks are
written out sequentially (with a page/section break between files) in the
target format.

This is a best-effort conversion: exact layout/formatting of the original
files is not preserved, but text, tables, and reading order are.
"""

import io
from xml.sax.saxutils import escape as _xml_escape

import openpyxl
from pypdf import PdfReader
from docx import Document as DocxDocument
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Table,
    TableStyle,
    PageBreak,
    Spacer,
)

from .excel_merger import (
    _safe_sheet_name,
    _load_xlsx_sheets,
    _load_xls_rows,
    _load_csv_rows,
    _is_encrypted as _excel_is_encrypted,
    _decrypt_office_file as _excel_decrypt,
)
from .docx_merger import _is_encrypted as _docx_is_encrypted, _decrypt as _docx_decrypt

SUPPORTED_OUTPUT_EXTS = ("pdf", "docx", "xlsx")


class ConversionError(Exception):
    pass


# --------------------------------------------------------------------------
# Extraction: turn each input file into a list of content blocks
#   {"type": "heading" | "paragraph", "text": str}
#   {"type": "table", "rows": [[cell, ...], ...], "name": optional str}
# --------------------------------------------------------------------------

def _extract_pdf_blocks(entry):
    stream = entry["stream"]
    name = entry["name"]
    password = entry.get("password")
    stream.seek(0)
    try:
        reader = PdfReader(stream)
    except Exception as exc:
        raise ConversionError(f"Could not read '{name}': {exc}") from exc

    if reader.is_encrypted:
        if not password:
            raise ConversionError(
                f"'{name}' is password protected. Please supply its password."
            )
        if reader.decrypt(password) == 0:
            raise ConversionError(f"Incorrect password for '{name}'.")

    blocks = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = line.strip()
            if line:
                blocks.append({"type": "paragraph", "text": line})
    if not blocks:
        blocks.append({"type": "paragraph", "text": "[No extractable text]"})
    return blocks


def _extract_docx_blocks(entry):
    stream = entry["stream"]
    name = entry["name"]
    password = entry.get("password")
    stream.seek(0)

    if _docx_is_encrypted(stream):
        if not password:
            raise ConversionError(
                f"'{name}' is password protected. Please supply its password."
            )
        stream = _docx_decrypt(stream, password, name)

    try:
        doc = DocxDocument(stream)
    except Exception as exc:
        raise ConversionError(f"Could not read '{name}': {exc}") from exc

    blocks = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style_name = (p.style.name or "").lower() if p.style else ""
        block_type = "heading" if style_name.startswith("heading") or style_name == "title" else "paragraph"
        blocks.append({"type": block_type, "text": text})
    for t in doc.tables:
        rows = [[cell.text for cell in row.cells] for row in t.rows]
        blocks.append({"type": "table", "rows": rows})
    return blocks


def _extract_excel_blocks(entry, used_names):
    stream = entry["stream"]
    name = entry["name"]
    ext = entry["ext"].lower()
    password = entry.get("password")
    base_name = name.rsplit(".", 1)[0]
    stream.seek(0)

    if ext in ("xlsx", "xls") and _excel_is_encrypted(stream):
        if not password:
            raise ConversionError(
                f"'{name}' is password protected. Please supply its password."
            )
        stream = _excel_decrypt(stream, password)
        ext = "xlsx"

    blocks = []
    try:
        if ext == "xlsx":
            for sheet_name, ws in _load_xlsx_sheets(stream, base_name, used_names):
                rows = [list(row) for row in ws.iter_rows(values_only=True)]
                blocks.append({"type": "heading", "text": sheet_name})
                blocks.append({"type": "table", "rows": rows, "name": sheet_name})
        elif ext == "xls":
            for sheet_name, rows in _load_xls_rows(stream, base_name, used_names):
                blocks.append({"type": "heading", "text": sheet_name})
                blocks.append({"type": "table", "rows": rows, "name": sheet_name})
        elif ext == "csv":
            for sheet_name, rows in _load_csv_rows(stream, base_name, used_names):
                blocks.append({"type": "table", "rows": rows, "name": sheet_name})
        else:
            raise ConversionError(f"Unsupported spreadsheet type: {ext}")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Could not read '{name}': {exc}") from exc
    return blocks


def extract_blocks(entry, used_sheet_names):
    ext = entry["ext"].lower()
    if ext == "pdf":
        return _extract_pdf_blocks(entry)
    if ext == "docx":
        return _extract_docx_blocks(entry)
    if ext in ("xlsx", "xls", "csv"):
        return _extract_excel_blocks(entry, used_sheet_names)
    raise ConversionError(f"Unsupported file type: '{entry['name']}'")


# --------------------------------------------------------------------------
# Writers: turn (filename, blocks) pairs into a single output file
# --------------------------------------------------------------------------

def _write_pdf(files_blocks):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    for idx, (fname, blocks) in enumerate(files_blocks):
        if idx > 0:
            story.append(PageBreak())
        for b in blocks:
            if b["type"] == "heading":
                story.append(Paragraph(_xml_escape(b["text"]), styles["Heading2"]))
                story.append(Spacer(1, 6))
            elif b["type"] == "paragraph":
                story.append(Paragraph(_xml_escape(b["text"]), styles["Normal"]))
                story.append(Spacer(1, 4))
            elif b["type"] == "table":
                rows = [
                    ["" if c is None else str(c) for c in row] for row in b["rows"]
                ]
                if rows:
                    table = Table(rows)
                    table.setStyle(
                        TableStyle(
                            [
                                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                                ("FONTSIZE", (0, 0), (-1, -1), 8),
                            ]
                        )
                    )
                    story.append(table)
                    story.append(Spacer(1, 10))

    if not story:
        story.append(Paragraph("Empty document", styles["Normal"]))

    doc.build(story)
    buf.seek(0)
    return buf


def _write_docx(files_blocks):
    out_doc = DocxDocument()
    for idx, (fname, blocks) in enumerate(files_blocks):
        if idx > 0:
            out_doc.add_page_break()
        for b in blocks:
            if b["type"] == "heading":
                out_doc.add_heading(b["text"], level=2)
            elif b["type"] == "paragraph":
                out_doc.add_paragraph(b["text"])
            elif b["type"] == "table":
                rows = b["rows"]
                if not rows:
                    continue
                ncols = max(len(r) for r in rows)
                table = out_doc.add_table(rows=len(rows), cols=ncols)
                try:
                    table.style = "Table Grid"
                except KeyError:
                    pass
                for r, row in enumerate(rows):
                    for c in range(ncols):
                        val = row[c] if c < len(row) else ""
                        table.cell(r, c).text = "" if val is None else str(val)
    buf = io.BytesIO()
    out_doc.save(buf)
    buf.seek(0)
    return buf


def _write_excel(files_blocks):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used_names = set()

    for fname, blocks in files_blocks:
        base = fname.rsplit(".", 1)[0]
        text_lines = []
        table_index = 0

        for b in blocks:
            if b["type"] in ("heading", "paragraph"):
                text_lines.append([b["text"]])
            elif b["type"] == "table":
                table_index += 1
                label = b.get("name") or (
                    base if table_index == 1 else f"{base}_{table_index}"
                )
                sheet_name = _safe_sheet_name(label, used_names)
                ws = wb.create_sheet(title=sheet_name)
                for row in b["rows"]:
                    ws.append(list(row))

        if text_lines:
            label = _safe_sheet_name(
                base if table_index == 0 else f"{base}_text", used_names
            )
            ws = wb.create_sheet(title=label)
            for row in text_lines:
                ws.append(row)

    if len(wb.worksheets) == 0:
        wb.create_sheet(title="Sheet1")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def convert_and_merge(files, output_ext):
    """
    files: list of dicts {"stream", "name", "ext", "password"}
    output_ext: one of "pdf", "docx", "xlsx"

    Returns BytesIO of the merged/converted output file.
    """
    if output_ext not in SUPPORTED_OUTPUT_EXTS:
        raise ConversionError(
            f"Unsupported output extension '.{output_ext}'. "
            f"Supported outputs: {', '.join('.' + e for e in SUPPORTED_OUTPUT_EXTS)}"
        )
    if not files:
        raise ConversionError("No files supplied.")

    used_sheet_names = set()
    files_blocks = []
    for entry in files:
        blocks = extract_blocks(entry, used_sheet_names)
        files_blocks.append((entry["name"], blocks))

    if output_ext == "pdf":
        return _write_pdf(files_blocks)
    if output_ext == "docx":
        return _write_docx(files_blocks)
    return _write_excel(files_blocks)

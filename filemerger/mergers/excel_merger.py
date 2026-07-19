"""
Merge Excel-family files (.xlsx, .xls, .csv) into a single .xlsx workbook.
Each source file (or, for multi-sheet source workbooks, each source sheet)
becomes its own sheet in the output workbook.
"""

import io
import csv
import re

import openpyxl
import xlrd
import msoffcrypto


class ExcelMergeError(Exception):
    pass


INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")


def _safe_sheet_name(name, used_names):
    name = INVALID_SHEET_CHARS.sub("_", name)[:31] or "Sheet"
    base = name
    i = 1
    while name in used_names:
        suffix = f"_{i}"
        name = (base[: 31 - len(suffix)]) + suffix
        i += 1
    used_names.add(name)
    return name


def _decrypt_office_file(stream, password):
    """Returns a BytesIO of the decrypted file, or raises ExcelMergeError."""
    stream.seek(0)
    office_file = msoffcrypto.OfficeFile(stream)
    try:
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception as exc:
        raise ExcelMergeError("Incorrect password or unsupported encryption.") from exc


def _is_encrypted(stream):
    stream.seek(0)
    try:
        office_file = msoffcrypto.OfficeFile(stream)
        return office_file.is_encrypted()
    except Exception:
        return False
    finally:
        stream.seek(0)


def _load_xlsx_sheets(stream, base_name, used_names):
    stream.seek(0)
    wb = openpyxl.load_workbook(stream, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        label = base_name if len(wb.worksheets) == 1 else f"{base_name}_{ws.title}"
        sheets.append((_safe_sheet_name(label, used_names), ws))
    return sheets


def _load_xls_rows(stream, base_name, used_names):
    stream.seek(0)
    book = xlrd.open_workbook(file_contents=stream.read())
    sheets = []
    for sheet in book.sheets():
        rows = []
        for r in range(sheet.nrows):
            rows.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
        label = base_name if book.nsheets == 1 else f"{base_name}_{sheet.name}"
        sheets.append((_safe_sheet_name(label, used_names), rows))
    return sheets


def _load_csv_rows(stream, base_name, used_names):
    stream.seek(0)
    raw = stream.read()
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ExcelMergeError(f"Could not decode CSV file '{base_name}'.")

    rows = list(csv.reader(text.splitlines()))
    label = _safe_sheet_name(base_name, used_names)
    return [(label, rows)]


def merge_excels(files):
    """
    files: list of dicts:
        {
            "stream": file-like object,
            "name": original filename,
            "ext": "xlsx" | "xls" | "csv",
            "password": str or None,
        }

    Returns: BytesIO containing the merged .xlsx workbook.
    """
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)  # remove default blank sheet
    used_names = set()

    for f in files:
        stream = f["stream"]
        name = f.get("name", "file")
        ext = f["ext"].lower()
        password = f.get("password")
        base_name = name.rsplit(".", 1)[0]

        if ext in ("xlsx", "xls") and _is_encrypted(stream):
            if not password:
                raise ExcelMergeError(
                    f"'{name}' is password protected. Please supply its password."
                )
            stream = _decrypt_office_file(stream, password)
            # msoffcrypto always decrypts old xls encryption into a readable
            # OLE/xls or xlsx stream; openpyxl can read the modern result.
            ext = "xlsx"

        try:
            if ext == "xlsx":
                sheets = _load_xlsx_sheets(stream, base_name, used_names)
                for sheet_name, ws in sheets:
                    out_ws = out_wb.create_sheet(title=sheet_name)
                    for row in ws.iter_rows(values_only=True):
                        out_ws.append(list(row))
            elif ext == "xls":
                sheets = _load_xls_rows(stream, base_name, used_names)
                for sheet_name, rows in sheets:
                    out_ws = out_wb.create_sheet(title=sheet_name)
                    for row in rows:
                        out_ws.append(row)
            elif ext == "csv":
                sheets = _load_csv_rows(stream, base_name, used_names)
                for sheet_name, rows in sheets:
                    out_ws = out_wb.create_sheet(title=sheet_name)
                    for row in rows:
                        out_ws.append(row)
            else:
                raise ExcelMergeError(f"Unsupported spreadsheet type: {ext}")
        except ExcelMergeError:
            raise
        except Exception as exc:
            raise ExcelMergeError(f"Could not read '{name}': {exc}") from exc

    if len(out_wb.worksheets) == 0:
        raise ExcelMergeError("No sheets found to merge.")

    out = io.BytesIO()
    out_wb.save(out)
    out.seek(0)
    return out

"""Merge multiple PDF files into a single PDF, with optional per-file passwords."""

import io
from pypdf import PdfReader, PdfWriter
from pypdf.errors import FileNotDecryptedError


class PdfMergeError(Exception):
    pass


def merge_pdfs(files):
    """
    files: list of dicts:
        {
            "stream": file-like object (bytes) positioned at 0,
            "name": original filename,
            "password": str or None,
        }

    Returns: BytesIO containing the merged PDF.
    """
    writer = PdfWriter()

    for f in files:
        stream = f["stream"]
        name = f.get("name", "file.pdf")
        password = f.get("password")

        stream.seek(0)
        try:
            reader = PdfReader(stream)
        except Exception as exc:
            raise PdfMergeError(f"Could not read '{name}': {exc}") from exc

        if reader.is_encrypted:
            if not password:
                raise PdfMergeError(
                    f"'{name}' is password protected. Please supply its password."
                )
            try:
                result = reader.decrypt(password)
                if result == 0:
                    raise PdfMergeError(f"Incorrect password for '{name}'.")
            except FileNotDecryptedError as exc:
                raise PdfMergeError(f"Incorrect password for '{name}'.") from exc

        try:
            for page in reader.pages:
                writer.add_page(page)
        except Exception as exc:
            raise PdfMergeError(f"Could not merge pages from '{name}': {exc}") from exc

    if len(writer.pages) == 0:
        raise PdfMergeError("No pages found to merge.")

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out

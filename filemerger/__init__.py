"""
filemerger
==========

A self-contained, cross-platform (Windows / Linux / macOS / Android-via-Termux)
file merging tool with a local web UI.

Supports merging:
  - PDF files            -> single .pdf
  - Excel files (xlsx/xls/csv) -> single .xlsx (each source file becomes a sheet)
  - Word files (docx)    -> single .docx

Password-protected input files are decrypted (given a user-supplied password)
before merging.
"""

__version__ = "1.0.0"

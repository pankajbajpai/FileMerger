import io
import json

from flask import Flask, render_template, request, send_file, jsonify
from pypdf import PdfReader

from .mergers.pdf_merger import merge_pdfs, PdfMergeError
from .mergers.excel_merger import merge_excels, ExcelMergeError, _is_encrypted as _excel_is_encrypted
from .mergers.docx_merger import merge_docx, DocxMergeError, _is_encrypted as _docx_is_encrypted
from .mergers.converter import convert_and_merge, ConversionError, SUPPORTED_OUTPUT_EXTS

INPUT_EXTS = {"pdf", "xlsx", "xls", "csv", "docx"}

# Which "family" each input extension belongs to. A merge can use the fast,
# high-fidelity path only when every input file's family matches the
# requested output family.
FAMILY_OF_EXT = {
    "pdf": "pdf",
    "docx": "word",
    "xlsx": "excel",
    "xls": "excel",
    "csv": "excel",
}
FAMILY_OF_OUTPUT = {"pdf": "pdf", "docx": "word", "xlsx": "excel"}

MIME_BY_EXT = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
}


def create_app():
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024  # 512 MB safety cap

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/api/merge", methods=["POST"])
    def api_merge():
        uploaded = request.files.getlist("files")
        if not uploaded:
            return jsonify(error="No files were uploaded."), 400

        output_name = (request.form.get("output_name") or "").strip()
        if not output_name or "." not in output_name:
            return jsonify(
                error="Please provide an output file name that includes an "
                f"extension (one of: {', '.join('.' + e for e in SUPPORTED_OUTPUT_EXTS)})."
            ), 400
        output_ext = output_name.rsplit(".", 1)[-1].lower()
        if output_ext not in SUPPORTED_OUTPUT_EXTS:
            return jsonify(
                error=f"Unsupported output extension '.{output_ext}'. "
                f"Supported outputs: {', '.join('.' + e for e in SUPPORTED_OUTPUT_EXTS)}."
            ), 400

        try:
            passwords = json.loads(request.form.get("passwords", "{}"))
        except (TypeError, ValueError):
            passwords = {}

        prepared = []
        for idx, storage in enumerate(uploaded):
            name = storage.filename or f"file_{idx}"
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            if ext not in INPUT_EXTS:
                return jsonify(
                    error=f"'{name}' has an unsupported file type. Supported "
                    f"input types: {', '.join('.' + e for e in sorted(INPUT_EXTS))}."
                ), 400
            buf = io.BytesIO(storage.read())
            prepared.append(
                {
                    "stream": buf,
                    "name": name,
                    "ext": ext,
                    "password": passwords.get(str(idx)) or None,
                }
            )

        # --- Only ask for passwords on files that are actually protected ---
        missing_pw = []
        for idx, entry in enumerate(prepared):
            if entry["password"]:
                continue
            if _needs_password(entry):
                missing_pw.append({"index": idx, "name": entry["name"]})
        if missing_pw:
            return (
                jsonify(
                    error="Some files are password protected.",
                    password_required=missing_pw,
                ),
                422,
            )

        # --- Choose fast (same-family) path or generic conversion path ---
        families = {FAMILY_OF_EXT[e["ext"]] for e in prepared}
        output_family = FAMILY_OF_OUTPUT[output_ext]

        try:
            if families == {output_family}:
                if output_family == "pdf":
                    result = merge_pdfs(prepared)
                elif output_family == "word":
                    result = merge_docx(prepared)
                else:
                    result = merge_excels(prepared)
            else:
                result = convert_and_merge(prepared, output_ext)
        except (PdfMergeError, ExcelMergeError, DocxMergeError, ConversionError) as exc:
            return jsonify(error=str(exc)), 422
        except Exception as exc:  # pragma: no cover - safety net
            return jsonify(error=f"Unexpected error while merging: {exc}"), 500

        if not output_name.lower().endswith(f".{output_ext}"):
            output_name = f"{output_name}.{output_ext}"

        return send_file(
            result,
            mimetype=MIME_BY_EXT[output_ext],
            as_attachment=True,
            download_name=output_name,
        )

    return app


def _needs_password(entry):
    stream = entry["stream"]
    ext = entry["ext"]
    stream.seek(0)
    try:
        if ext == "pdf":
            return PdfReader(stream).is_encrypted
        if ext in ("xlsx", "xls"):
            return _excel_is_encrypted(stream)
        if ext == "docx":
            return _docx_is_encrypted(stream)
        return False
    except Exception:
        return False
    finally:
        stream.seek(0)

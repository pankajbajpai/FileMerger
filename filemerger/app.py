import io
import json
import datetime

from flask import Flask, render_template, request, send_file, jsonify

from .mergers.pdf_merger import merge_pdfs, PdfMergeError
from .mergers.excel_merger import merge_excels, ExcelMergeError
from .mergers.docx_merger import merge_docx, DocxMergeError

EXCEL_EXTS = {"xlsx", "xls", "csv"}
PDF_EXTS = {"pdf"}
WORD_EXTS = {"docx"}

ALLOWED_BY_TYPE = {
    "pdf": PDF_EXTS,
    "excel": EXCEL_EXTS,
    "word": WORD_EXTS,
}


def create_app():
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024  # 512 MB safety cap

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/api/merge", methods=["POST"])
    def api_merge():
        file_type = request.form.get("file_type")
        if file_type not in ALLOWED_BY_TYPE:
            return jsonify(error="Invalid or missing file_type."), 400

        uploaded = request.files.getlist("files")
        if not uploaded or len(uploaded) < 1:
            return jsonify(error="No files were uploaded."), 400

        try:
            passwords = json.loads(request.form.get("passwords", "{}"))
        except (TypeError, ValueError):
            passwords = {}

        allowed_exts = ALLOWED_BY_TYPE[file_type]
        prepared = []
        for idx, storage in enumerate(uploaded):
            name = storage.filename or f"file_{idx}"
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            if ext not in allowed_exts:
                return jsonify(
                    error=f"'{name}' does not match the selected file type "
                    f"({', '.join(sorted(allowed_exts))})."
                ), 400
            buf = io.BytesIO(storage.read())
            entry = {"stream": buf, "name": name, "password": passwords.get(str(idx)) or None}
            if file_type == "excel":
                entry["ext"] = ext
            prepared.append(entry)

        try:
            if file_type == "pdf":
                result = merge_pdfs(prepared)
                mimetype = "application/pdf"
                default_ext = "pdf"
            elif file_type == "excel":
                result = merge_excels(prepared)
                mimetype = (
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                )
                default_ext = "xlsx"
            else:  # word
                result = merge_docx(prepared)
                mimetype = (
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"
                )
                default_ext = "docx"
        except (PdfMergeError, ExcelMergeError, DocxMergeError) as exc:
            return jsonify(error=str(exc)), 422
        except Exception as exc:  # pragma: no cover - safety net
            return jsonify(error=f"Unexpected error while merging: {exc}"), 500

        out_name = request.form.get("output_name", "").strip()
        if not out_name:
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"merged_{file_type}_{stamp}"
        if not out_name.lower().endswith(f".{default_ext}"):
            out_name = f"{out_name}.{default_ext}"

        return send_file(
            result,
            mimetype=mimetype,
            as_attachment=True,
            download_name=out_name,
        )

    return app

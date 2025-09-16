import os
import tempfile
from flask import (
    Flask,
    request,
    send_file,
    render_template_string,
    abort,
    after_this_request,
)
from conversion_utils import write_temp_file, pdf_to_docx, docx_to_pdf

app = Flask(__name__)

# Simple HTML form
HTML = """
<!doctype html>
<title>PDF⇄Word Converter</title>
<h1>Upload PDF or Word</h1>
<form method=post enctype=multipart/form-data>
  <input type=file name=file><br><br>
  <button type=submit>Convert</button>
</form>
"""

# Serve the OpenAPI spec so Copilot can discover your API
@app.route("/swagger.yaml", methods=["GET"])
def serve_swagger():
    swagger_path = os.path.join(os.path.dirname(__file__), "swagger.yaml")
    if not os.path.exists(swagger_path):
        abort(404, "swagger.yaml not found")
    return send_file(swagger_path, mimetype="application/x-yaml")


@app.route("/", methods=["GET", "POST"])
def convert():
    if request.method == "GET":
        return render_template_string(HTML)

    # POST: user uploaded a file
    file = request.files.get("file")
    if not file:
        return abort(400, "No file uploaded")

    filename = file.filename
    data     = file.read()

    # Save to temp and convert
    suffix = os.path.splitext(filename)[1].lower()
    tmp_in = write_temp_file(data, suffix=suffix)

    if suffix == ".pdf":
        out_path = pdf_to_docx(tmp_in, original_name=filename)
        mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif suffix in (".doc", ".docx"):
        out_path = docx_to_pdf(tmp_in, original_name=filename)
        mimetype = "application/pdf"
    else:
        # Cleanup tmp_in before aborting
        try:
            os.remove(tmp_in)
        except Exception:
            pass
        return abort(400, "Unsupported file type")

    # Schedule cleanup of temp files after response
    @after_this_request
    def cleanup(response):
        try:
            os.remove(tmp_in)
        except Exception:
            pass
        try:
            os.remove(out_path)
        except Exception:
            pass
        return response

    # Send converted file back with the original base filename
    return send_file(
        out_path,
        as_attachment=True,
        download_name=os.path.basename(out_path),
        mimetype=mimetype,
    )


if __name__ == "__main__":
    # For demo purposes; not for production use
    app.run(port=5000)

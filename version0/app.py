from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import uuid
from pdf_to_json import extract_pdf_content, generate_slide_data
from main import create_presentation_from_json
from pptx import Presentation

BASE_UPLOAD_FOLDER = "uploads"
BASE_OUTPUT_FOLDER = "output"

app = Flask(__name__)

os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET"])
def upload_form():
    return render_template("upload.html")

@app.route("/upload", methods=["POST"])
def handle_upload():
    file = request.files["pdf"]
    if not file or file.filename == "":
        return "No file selected", 400

    # Generate unique ID for this session
    session_id = uuid.uuid4().hex
    upload_folder = os.path.join(BASE_UPLOAD_FOLDER, session_id)
    output_folder = os.path.join(BASE_OUTPUT_FOLDER, session_id)
    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    original_name = file.filename
    pdf_path = os.path.join(upload_folder, original_name)
    pptx_name = original_name.replace(".pdf", ".pptx")
    pptx_path = os.path.join(output_folder, pptx_name)

    file.save(pdf_path)

    try:
        content = extract_pdf_content(pdf_path)
        slide_data = generate_slide_data(content)
        prs = create_presentation_from_json(slide_data)
        prs.save(pptx_path)

        return redirect(url_for("result_page", session_id=session_id, filename=pptx_name))
    except Exception as e:
        return f"Error during processing: {str(e)}", 500

@app.route("/result/<session_id>/<filename>")
def result_page(session_id, filename):
    return render_template("result.html", session_id=session_id, filename=filename)

@app.route("/download/<session_id>/<filename>")
def download_file(session_id, filename):
    file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)

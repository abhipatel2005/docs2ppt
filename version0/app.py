from flask import Flask, render_template, request, send_file
import os
import uuid
from pdf_to_json import extract_pdf_content, generate_slide_data
from main import create_presentation_from_json
from pptx import Presentation

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET"])
def upload_form():
    return render_template("upload.html")

@app.route("/upload", methods=["POST"])
def handle_upload():
    if "pdf" not in request.files:
        return "No file part", 400

    file = request.files["pdf"]
    if file.filename == "":
        return "No selected file", 400

    # Save uploaded PDF
    filename = f"{uuid.uuid4().hex}_{file.filename}"
    pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(pdf_path)

    # Generate JSON from PDF using Gemini
    try:
        content = extract_pdf_content(pdf_path)
        slide_data = generate_slide_data(content)

        # Generate PPTX using main.py's logic
        prs = create_presentation_from_json(slide_data)
        output_path = os.path.join(app.config["OUTPUT_FOLDER"], filename.replace(".pdf", ".pptx"))
        prs.save(output_path)

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return f"Error during processing: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)

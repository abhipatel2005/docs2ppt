# # from flask import Flask, render_template, request, redirect, url_for, send_file
# # import os
# # import uuid
# # from pdf_to_json import extract_pdf_content, generate_slide_data
# # from main import create_presentation_from_json
# # from pptx import Presentation
# # import sys
# # print("🧪 Python version:", sys.version)

# # BASE_UPLOAD_FOLDER = "uploads"
# # BASE_OUTPUT_FOLDER = "output"

# # app = Flask(__name__)

# # os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
# # os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)

# # @app.route("/", methods=["GET"])
# # def upload_form():
# #     return render_template("upload.html")

# # @app.route("/upload", methods=["POST"])
# # def handle_upload():
# #     file = request.files["pdf"]
# #     if not file or file.filename == "":
# #         return "No file selected", 400

# #     # Generate unique ID for this session
# #     session_id = uuid.uuid4().hex
# #     upload_folder = os.path.join(BASE_UPLOAD_FOLDER, session_id)
# #     output_folder = os.path.join(BASE_OUTPUT_FOLDER, session_id)
# #     os.makedirs(upload_folder, exist_ok=True)
# #     os.makedirs(output_folder, exist_ok=True)

# #     original_name = file.filename
# #     pdf_path = os.path.join(upload_folder, original_name)
# #     pptx_name = original_name.replace(".pdf", ".pptx")
# #     pptx_path = os.path.join(output_folder, pptx_name)

# #     file.save(pdf_path)

# #     try:
# #         content = extract_pdf_content(pdf_path)
# #         slide_data = generate_slide_data(content)
# #         prs = create_presentation_from_json(slide_data)
# #         prs.save(pptx_path)

# #         return redirect(url_for("result_page", session_id=session_id, filename=pptx_name))
# #     except Exception as e:
# #         return f"Error during processing: {str(e)}", 500

# # @app.route("/result/<session_id>/<filename>")
# # def result_page(session_id, filename):
# #     return render_template("result.html", session_id=session_id, filename=filename)

# # @app.route("/download/<session_id>/<filename>")
# # def download_file(session_id, filename):
# #     file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
# #     return send_file(file_path, as_attachment=True)

# # if __name__ == "__main__":
# #     app.run(debug=True)

# from flask import Flask, render_template, request, redirect, url_for, send_file
# import os
# import uuid
# from pdf_to_json import extract_pdf_content, generate_slide_data
# from main import create_presentation_from_json
# from pptx import Presentation
# import sys
# print("🧪 Python version:", sys.version)

# BASE_UPLOAD_FOLDER = "uploads"
# BASE_OUTPUT_FOLDER = "output"

# app = Flask(__name__)

# os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)

# @app.route("/", methods=["GET"])
# def upload_form():
#     return render_template("upload.html")

# @app.route("/upload", methods=["POST"])
# def handle_upload():
#     file = request.files["pdf"]
#     if not file or file.filename == "":
#         return "No file selected", 400

#     # Generate unique ID for this session
#     session_id = uuid.uuid4().hex
#     upload_folder = os.path.join(BASE_UPLOAD_FOLDER, session_id)
#     output_folder = os.path.join(BASE_OUTPUT_FOLDER, session_id)
#     os.makedirs(upload_folder, exist_ok=True)
#     os.makedirs(output_folder, exist_ok=True)

#     original_name = file.filename
#     pdf_path = os.path.join(upload_folder, original_name)
#     pptx_name = original_name.replace(".pdf", ".pptx")
#     pptx_path = os.path.join(output_folder, pptx_name)

#     file.save(pdf_path)

#     try:
#         content = extract_pdf_content(pdf_path)
#         slide_data = generate_slide_data(content)
#         prs = create_presentation_from_json(slide_data)
#         prs.save(pptx_path)

#         return redirect(url_for("result_page", session_id=session_id, filename=pptx_name))
#     except Exception as e:
#         return f"Error during processing: {str(e)}", 500

# @app.route("/result/<session_id>/<filename>")
# def result_page(session_id, filename):
#     return render_template("result.html", session_id=session_id, filename=filename)

# @app.route("/download/<session_id>/<filename>")
# def download_file(session_id, filename):
#     file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
#     return send_file(file_path, as_attachment=True)

# if __name__ == "__main__":
#     # Get port from environment variable (Render sets this)
#     port = int(os.environ.get("PORT", 5000))
#     # Important: bind to 0.0.0.0 to make the server publicly accessible
#     app.run(host="0.0.0.0", port=port, debug=True)


#! code with cleanup

# from flask import Flask, render_template, request, redirect, url_for, send_file
# import os
# import uuid
# import sys
# import threading
# import time
# from datetime import datetime, timedelta
# from pdf_to_json import extract_pdf_content, generate_slide_data
# from main import create_presentation_from_json
# from pptx import Presentation

# print("🧪 Python version:", sys.version)

# BASE_UPLOAD_FOLDER = "uploads"
# BASE_OUTPUT_FOLDER = "output"
# PDF_IMAGES_FOLDER = "pdf_images"

# app = Flask(__name__)
# os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)
# os.makedirs(PDF_IMAGES_FOLDER, exist_ok=True)

# # --- CLEANUP FUNCTION ---
# EXPIRATION_MINUTES = 30

# def is_file_expired(file_path, expiration_minutes):
#     file_mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
#     return datetime.now() - file_mtime > timedelta(minutes=expiration_minutes)

# def cleanup_old_files_and_folders(base_folder, expiration_minutes):
#     for session_id in os.listdir(base_folder):
#         folder_path = os.path.join(base_folder, session_id)
#         if not os.path.isdir(folder_path):
#             continue

#         try:
#             # Delete old files
#             for root, dirs, files in os.walk(folder_path):
#                 for filename in files:
#                     file_path = os.path.join(root, filename)
#                     if is_file_expired(file_path, expiration_minutes):
#                         os.remove(file_path)
#                         print(f"🗑️ Deleted file: {file_path}")

#             # Remove empty folders
#             for root, dirs, files in os.walk(folder_path, topdown=False):
#                 if not os.listdir(root):  # empty dir
#                     os.rmdir(root)
#                     print(f"📁 Deleted empty folder: {root}")

#         except Exception as e:
#             print(f"⚠️ Error cleaning {folder_path}: {e}")

# def background_cleanup_loop():
#     while True:
#         print("🧹 Running background cleanup...")
#         cleanup_old_files_and_folders(BASE_UPLOAD_FOLDER, EXPIRATION_MINUTES)
#         cleanup_old_files_and_folders(BASE_OUTPUT_FOLDER, EXPIRATION_MINUTES)
#         cleanup_old_files_and_folders(PDF_IMAGES_FOLDER, EXPIRATION_MINUTES)
#         time.sleep(300)  # every 5 minutes

# @app.before_request
# def start_background_cleanup():
#     if not hasattr(start_background_cleanup, "thread"):
#         thread = threading.Thread(target=background_cleanup_loop)
#         thread.daemon = True
#         thread.start()
#         start_background_cleanup.thread = thread

# # --- ROUTES ---
# @app.route("/", methods=["GET"])
# def upload_form():
#     return render_template("upload.html")

# @app.route("/upload", methods=["POST"])
# def handle_upload():
#     file = request.files["pdf"]
#     if not file or file.filename == "":
#         return "No file selected", 400

#     session_id = uuid.uuid4().hex
#     upload_folder = os.path.join(BASE_UPLOAD_FOLDER, session_id)
#     output_folder = os.path.join(BASE_OUTPUT_FOLDER, session_id)
#     os.makedirs(upload_folder, exist_ok=True)
#     os.makedirs(output_folder, exist_ok=True)

#     original_name = file.filename
#     pdf_path = os.path.join(upload_folder, original_name)
#     pptx_name = original_name.replace(".pdf", ".pptx")
#     pptx_path = os.path.join(output_folder, pptx_name)

#     file.save(pdf_path)

#     try:
#         content = extract_pdf_content(pdf_path)
#         slide_data = generate_slide_data(content)
#         prs = create_presentation_from_json(slide_data)
#         prs.save(pptx_path)
#         return redirect(url_for("result_page", session_id=session_id, filename=pptx_name))
#     except Exception as e:
#         return f"Error during processing: {str(e)}", 500

# @app.route("/result")
# def result_page():
#     session_id = request.args.get("session_id")
#     filename = request.args.get("filename")
#     return render_template("result.html", session_id=session_id, filename=filename)

# @app.route("/download")
# def download_file():
#     session_id = request.args.get("session_id")
#     filename = request.args.get("filename")
#     file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
#     return send_file(file_path, as_attachment=True)

# if __name__ == "__main__":
#     port = int(os.environ.get("PORT", 5000))
#     app.run(host="0.0.0.0", port=port, debug=True)

#!properly running the cleanup function and 15 min deletion policy

from flask import Flask, render_template, request, redirect, url_for, send_file, render_template_string
import os
import uuid
import sys
import threading
import time
from datetime import datetime, timedelta
from pdf_to_json import extract_pdf_content, extract_docx_content, generate_slide_data
from main import create_presentation_from_json
from pptx import Presentation

print("🧪 Python version:", sys.version)

BASE_UPLOAD_FOLDER = "uploads"
BASE_OUTPUT_FOLDER = "output"
PDF_IMAGES_FOLDER = "pdf_images"
DOCX_IMAGES_FOLDER = "docx_extracted"
EXPIRATION_MINUTES = 15
EXPIRATION_MINUTES_FOR_IMAGES = 5

app = Flask(__name__)
os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)
os.makedirs(PDF_IMAGES_FOLDER, exist_ok=True)
os.makedirs(DOCX_IMAGES_FOLDER, exist_ok=True)

# --- CLEANUP ---
def is_file_expired(file_path, expiration_minutes):
    file_mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
    return datetime.now() - file_mtime > timedelta(minutes=expiration_minutes)

def cleanup_old_files_and_folders(base_folder, expiration_minutes):
    for session_id in os.listdir(base_folder):
        folder_path = os.path.join(base_folder, session_id)
        if not os.path.isdir(folder_path):
            continue
        try:
            for root, dirs, files in os.walk(folder_path):
                for filename in files:
                    file_path = os.path.join(root, filename)
                    if is_file_expired(file_path, expiration_minutes):
                        os.remove(file_path)
                        print(f"🗑️ Deleted file: {file_path}")
            for root, dirs, files in os.walk(folder_path, topdown=False):
                if not os.listdir(root):
                    os.rmdir(root)
                    print(f"📁 Deleted empty folder: {root}")
        except Exception as e:
            print(f"⚠️ Error cleaning {folder_path}: {e}")

def background_cleanup_loop():
    while True:
        print("🧹 Running background cleanup...")
        cleanup_old_files_and_folders(PDF_IMAGES_FOLDER, EXPIRATION_MINUTES)
        cleanup_old_files_and_folders(BASE_UPLOAD_FOLDER, EXPIRATION_MINUTES)
        cleanup_old_files_and_folders(BASE_OUTPUT_FOLDER, EXPIRATION_MINUTES)
        cleanup_old_files_and_folders(DOCX_IMAGES_FOLDER, EXPIRATION_MINUTES)
        time.sleep(300)  # every 5 minutes

@app.before_request
def start_background_cleanup():
    if not hasattr(start_background_cleanup, "thread"):
        thread = threading.Thread(target=background_cleanup_loop)
        thread.daemon = True
        thread.start()
        start_background_cleanup.thread = thread

# --- ROUTES ---
@app.route("/", methods=["GET"])
def upload_form():
    return render_template("upload.html")

@app.route("/upload", methods=["POST"])
def handle_upload():
    file = request.files["file"]
    if not file or file.filename == "":
        return "No file selected", 400

    session_id = uuid.uuid4().hex
    upload_folder = os.path.join(BASE_UPLOAD_FOLDER, session_id)
    output_folder = os.path.join(BASE_OUTPUT_FOLDER, session_id)
    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    original_name = file.filename
    file_path = os.path.join(upload_folder, original_name)
    pptx_name = os.path.splitext(original_name)[0] + ".pptx"
    pptx_path = os.path.join(output_folder, pptx_name)

    file.save(file_path)

    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".pdf":
            content = extract_pdf_content(file_path)
        elif ext in [".docx", ".doc"]:
            content = extract_docx_content(file_path)
        else:
            return "Unsupported file type. Please upload PDF or DOCX.", 400

        slide_data = generate_slide_data(content)
        prs = create_presentation_from_json(slide_data)
        prs.save(pptx_path)

        return redirect(url_for("result_page", session_id=session_id, filename=pptx_name))
    except Exception as e:
        return f"Error during processing: {str(e)}", 500

@app.route("/result")
def result_page():
    session_id = request.args.get("session_id")
    filename = request.args.get("filename")
    return render_template("result.html", session_id=session_id, filename=filename)

# @app.route("/download")
# def download_file():
#     session_id = request.args.get("session_id")
#     filename = request.args.get("filename")
#     file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
#     return send_file(file_path, as_attachment=True)

@app.route("/download")
def download_file():
    session_id = request.args.get("session_id")
    filename = request.args.get("filename")
    file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)

    if not os.path.exists(file_path):
        # Custom message for missing file
        return render_template_string("""
            <!DOCTYPE html>
            <html>
            <head>
                <title>File Not Found</title>
                <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
            </head>
            <body class="bg-light d-flex flex-column justify-content-center align-items-center" style="height: 100vh;">
                <div class="text-center">
                    <h2 class="text-danger">⏳ File Expired</h2>
                    <p class="lead">This file was automatically deleted due to our 15-minute retention policy.</p>
                    <p>You can re-upload your PDF to generate the presentation again.</p>
                    <a href="/" class="btn btn-primary">Go Back to Upload</a>
                </div>
            </body>
            </html>
        """), 404

    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

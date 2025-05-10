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

from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import uuid
import sys
import threading
import time
from datetime import datetime, timedelta
from pdf_to_json import extract_pdf_content, generate_slide_data
from main import create_presentation_from_json
from pptx import Presentation

print("🧪 Python version:", sys.version)

BASE_UPLOAD_FOLDER = "uploads"
BASE_OUTPUT_FOLDER = "output"

app = Flask(__name__)
os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)

# --- CLEANUP FUNCTION ---
EXPIRATION_MINUTES = 30

def is_file_expired(file_path, expiration_minutes):
    file_mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
    return datetime.now() - file_mtime > timedelta(minutes=expiration_minutes)

def cleanup_old_files_and_folders(base_folder, expiration_minutes):
    for session_id in os.listdir(base_folder):
        folder_path = os.path.join(base_folder, session_id)
        if not os.path.isdir(folder_path):
            continue

        try:
            # Delete old files
            for root, dirs, files in os.walk(folder_path):
                for filename in files:
                    file_path = os.path.join(root, filename)
                    if is_file_expired(file_path, expiration_minutes):
                        os.remove(file_path)
                        print(f"🗑️ Deleted file: {file_path}")

            # Remove empty folders
            for root, dirs, files in os.walk(folder_path, topdown=False):
                if not os.listdir(root):  # empty dir
                    os.rmdir(root)
                    print(f"📁 Deleted empty folder: {root}")

        except Exception as e:
            print(f"⚠️ Error cleaning {folder_path}: {e}")

def background_cleanup_loop():
    while True:
        print("🧹 Running background cleanup...")
        cleanup_old_files_and_folders(BASE_UPLOAD_FOLDER, EXPIRATION_MINUTES)
        cleanup_old_files_and_folders(BASE_OUTPUT_FOLDER, EXPIRATION_MINUTES)
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
    file = request.files["pdf"]
    if not file or file.filename == "":
        return "No file selected", 400

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

@app.route("/result")
def result_page():
    session_id = request.args.get("session_id")
    filename = request.args.get("filename")
    return render_template("result.html", session_id=session_id, filename=filename)

@app.route("/download")
def download_file():
    session_id = request.args.get("session_id")
    filename = request.args.get("filename")
    file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

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

# from flask import Flask, render_template, request, redirect, url_for, send_file, render_template_string
# import os
# import uuid
# import sys
# import threading
# import time
# from datetime import datetime, timedelta
# from pdf_to_json import extract_pdf_content, extract_docx_content, generate_slide_data
# from main import create_presentation_from_json
# from pptx import Presentation

# print("🧪 Python version:", sys.version)

# BASE_UPLOAD_FOLDER = "uploads"
# BASE_OUTPUT_FOLDER = "output"
# PDF_IMAGES_FOLDER = "pdf_images"
# DOCX_IMAGES_FOLDER = "docx_extracted"
# EXPIRATION_MINUTES = 15
# EXPIRATION_MINUTES_FOR_IMAGES = 5

# app = Flask(__name__)
# os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)
# os.makedirs(PDF_IMAGES_FOLDER, exist_ok=True)
# os.makedirs(DOCX_IMAGES_FOLDER, exist_ok=True)

# # --- CLEANUP ---
# def is_file_expired(file_path, expiration_minutes):
#     file_mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
#     return datetime.now() - file_mtime > timedelta(minutes=expiration_minutes)

# def cleanup_old_files_and_folders(base_folder, expiration_minutes):
#     for session_id in os.listdir(base_folder):
#         folder_path = os.path.join(base_folder, session_id)
#         if not os.path.isdir(folder_path):
#             continue
#         try:
#             for root, dirs, files in os.walk(folder_path):
#                 for filename in files:
#                     file_path = os.path.join(root, filename)
#                     if is_file_expired(file_path, expiration_minutes):
#                         os.remove(file_path)
#                         print(f"🗑️ Deleted file: {file_path}")
#             for root, dirs, files in os.walk(folder_path, topdown=False):
#                 if not os.listdir(root):
#                     os.rmdir(root)
#                     print(f"📁 Deleted empty folder: {root}")
#         except Exception as e:
#             print(f"⚠️ Error cleaning {folder_path}: {e}")

# def background_cleanup_loop():
#     while True:
#         print("🧹 Running background cleanup...")
#         cleanup_old_files_and_folders(PDF_IMAGES_FOLDER, EXPIRATION_MINUTES)
#         cleanup_old_files_and_folders(BASE_UPLOAD_FOLDER, EXPIRATION_MINUTES)
#         cleanup_old_files_and_folders(BASE_OUTPUT_FOLDER, EXPIRATION_MINUTES)
#         cleanup_old_files_and_folders(DOCX_IMAGES_FOLDER, EXPIRATION_MINUTES)
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
#     file = request.files["file"]
#     if not file or file.filename == "":
#         return "No file selected", 400

#     session_id = uuid.uuid4().hex
#     upload_folder = os.path.join(BASE_UPLOAD_FOLDER, session_id)
#     output_folder = os.path.join(BASE_OUTPUT_FOLDER, session_id)
#     os.makedirs(upload_folder, exist_ok=True)
#     os.makedirs(output_folder, exist_ok=True)

#     original_name = file.filename
#     file_path = os.path.join(upload_folder, original_name)
#     pptx_name = os.path.splitext(original_name)[0] + ".pptx"
#     pptx_path = os.path.join(output_folder, pptx_name)

#     file.save(file_path)

#     try:
#         ext = os.path.splitext(file_path)[1].lower()
#         if ext == ".pdf":
#             content = extract_pdf_content(file_path)
#         elif ext in [".docx", ".doc"]:
#             content = extract_docx_content(file_path)
#         else:
#             return "Unsupported file type. Please upload PDF or DOCX.", 400

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

# # @app.route("/download")
# # def download_file():
# #     session_id = request.args.get("session_id")
# #     filename = request.args.get("filename")
# #     file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
# #     return send_file(file_path, as_attachment=True)

# @app.route("/download")
# def download_file():
#     session_id = request.args.get("session_id")
#     filename = request.args.get("filename")
#     file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)

#     if not os.path.exists(file_path):
#         # Custom message for missing file
#         return render_template_string("""
#             <!DOCTYPE html>
#             <html>
#             <head>
#                 <title>File Not Found</title>
#                 <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
#             </head>
#             <body class="bg-light d-flex flex-column justify-content-center align-items-center" style="height: 100vh;">
#                 <div class="text-center">
#                     <h2 class="text-danger">⏳ File Expired</h2>
#                     <p class="lead">This file was automatically deleted due to our 15-minute retention policy.</p>
#                     <p>You can re-upload your PDF to generate the presentation again.</p>
#                     <a href="/" class="btn btn-primary">Go Back to Upload</a>
#                 </div>
#             </body>
#             </html>
#         """), 404

#     return send_file(file_path, as_attachment=True)

# if __name__ == "__main__":
#     port = int(os.environ.get("PORT", 5000))
#     app.run(host="0.0.0.0", port=port, debug=True)


#!testing the cleanup function and 15 min deletion policy very strctly

from flask import Flask, render_template, request, redirect, url_for, send_file, render_template_string
import os
import uuid
import sys
import threading
import time
import logging
from datetime import datetime, timedelta
from pdf_to_json import extract_pdf_content, extract_docx_content, generate_slide_data
from main import create_presentation_from_json
from pptx import Presentation

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

logger.info("🧪 Python version: %s", sys.version)

BASE_UPLOAD_FOLDER = "uploads"
BASE_OUTPUT_FOLDER = "output"
PDF_IMAGES_FOLDER = "pdf_images"
DOCX_IMAGES_FOLDER = "docx_extracted"
EXPIRATION_MINUTES = 15  # Main content expiration (15 minutes)
CLEANUP_INTERVAL = 300   # Cleanup every 5 minutes (300 seconds)

# Track active sessions to prevent cleanup during processing
active_sessions = set()
active_sessions_lock = threading.Lock()

app = Flask(__name__)
os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
os.makedirs(BASE_OUTPUT_FOLDER, exist_ok=True)
os.makedirs(PDF_IMAGES_FOLDER, exist_ok=True)
os.makedirs(DOCX_IMAGES_FOLDER, exist_ok=True)

# --- SESSION MANAGEMENT ---
def register_active_session(session_id):
    """Register a session as active to prevent cleanup"""
    with active_sessions_lock:
        active_sessions.add(session_id)
        logger.info(f"📝 Registered active session: {session_id}")

def unregister_active_session(session_id):
    """Mark a session as no longer active"""
    with active_sessions_lock:
        if session_id in active_sessions:
            active_sessions.remove(session_id)
            logger.info(f"✓ Unregistered session: {session_id}")

def is_session_active(session_id):
    """Check if a session is currently active"""
    with active_sessions_lock:
        return session_id in active_sessions

# --- IMPROVED CLEANUP ---
def is_file_expired(file_path, expiration_minutes):
    """Check if a file is older than the expiration time"""
    try:
        file_mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
        elapsed = datetime.now() - file_mtime
        is_expired = elapsed > timedelta(minutes=expiration_minutes)
        
        if is_expired:
            logger.debug(f"File {file_path} is {elapsed.total_seconds()/60:.1f} minutes old (limit: {expiration_minutes})")
        
        return is_expired
    except Exception as e:
        logger.error(f"Error checking expiration for {file_path}: {e}")
        return False  # Don't delete files if there's an error

def is_folder_expired(folder_path, expiration_minutes):
    """Check if a folder is older than the expiration time based on its contents"""
    if not os.path.exists(folder_path):
        return False
        
    try:
        # Check if the folder has any files
        any_files = False
        newest_file_time = datetime.fromtimestamp(0)  # Initialize with Unix epoch
        
        for root, _, files in os.walk(folder_path):
            for filename in files:
                any_files = True
                file_path = os.path.join(root, filename)
                file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                if file_time > newest_file_time:
                    newest_file_time = file_time
        
        if not any_files:
            # Empty directory - check the folder creation time
            folder_time = datetime.fromtimestamp(os.path.getctime(folder_path))
            return datetime.now() - folder_time > timedelta(minutes=expiration_minutes)
        else:
            # Folder with files - check the newest file's timestamp
            return datetime.now() - newest_file_time > timedelta(minutes=expiration_minutes)
    except Exception as e:
        logger.error(f"Error checking folder expiration for {folder_path}: {e}")
        return False  # Don't delete folders if there's an error

def cleanup_old_files_and_folders(base_folder, expiration_minutes):
    """Clean up expired files and folders while respecting active sessions"""
    if not os.path.exists(base_folder):
        return
        
    try:
        # First pass: collect sessions to clean up
        sessions_to_clean = []
        
        for session_id in os.listdir(base_folder):
            folder_path = os.path.join(base_folder, session_id)
            if not os.path.isdir(folder_path):
                continue
                
            # Skip active sessions
            if is_session_active(session_id):
                logger.info(f"⏳ Skipping cleanup for active session: {session_id}")
                continue
                
            # Check if the folder is expired
            if is_folder_expired(folder_path, expiration_minutes):
                sessions_to_clean.append(session_id)
                
        # Second pass: clean up expired sessions
        for session_id in sessions_to_clean:
            folder_path = os.path.join(base_folder, session_id)
            
            # Double-check the session is still not active
            if is_session_active(session_id):
                continue
                
            # Clean up files
            try:
                logger.info(f"🧹 Cleaning up session {session_id} in {base_folder}")
                
                # Delete files
                for root, _, files in os.walk(folder_path):
                    for filename in files:
                        file_path = os.path.join(root, filename)
                        if is_file_expired(file_path, expiration_minutes):
                            os.remove(file_path)
                            logger.info(f"🗑️ Deleted file: {file_path}")
                
                # Remove empty directories (bottom-up)
                for root, dirs, files in os.walk(folder_path, topdown=False):
                    if not os.listdir(root):  # Check if directory is empty
                        os.rmdir(root)
                        logger.info(f"📁 Deleted empty folder: {root}")
            except Exception as e:
                logger.error(f"⚠️ Error cleaning {folder_path}: {e}")
    except Exception as e:
        logger.error(f"⚠️ Error during cleanup of {base_folder}: {e}")

def background_cleanup_loop():
    """Background thread that periodically cleans up expired files"""
    while True:
        try:
            logger.info("🧹 Running background cleanup...")
            
            # Clean up folders in order of least to most important
            cleanup_old_files_and_folders(PDF_IMAGES_FOLDER, EXPIRATION_MINUTES)
            cleanup_old_files_and_folders(DOCX_IMAGES_FOLDER, EXPIRATION_MINUTES)
            cleanup_old_files_and_folders(BASE_UPLOAD_FOLDER, EXPIRATION_MINUTES)
            cleanup_old_files_and_folders(BASE_OUTPUT_FOLDER, EXPIRATION_MINUTES)
            
            logger.info("✅ Cleanup completed")
        except Exception as e:
            logger.error(f"⚠️ Error in cleanup loop: {e}")
        
        # Sleep for the cleanup interval
        time.sleep(CLEANUP_INTERVAL)

# Start the cleanup thread when the app starts
cleanup_thread = None

# Replace @app.before_first_request with code that runs when the app starts
def start_background_cleanup():
    """Start the background cleanup thread"""
    global cleanup_thread
    if cleanup_thread is None or not cleanup_thread.is_alive():
        cleanup_thread = threading.Thread(target=background_cleanup_loop)
        cleanup_thread.daemon = True
        cleanup_thread.start()
        logger.info("🧵 Started background cleanup thread")

# --- ROUTES ---
@app.route("/", methods=["GET"])
def upload_form():
    # Start cleanup thread on first request
    start_background_cleanup()
    return render_template("upload.html")

@app.route("/upload", methods=["POST"])
def handle_upload():
    file = request.files["file"]
    if not file or file.filename == "":
        return "No file selected", 400

    session_id = uuid.uuid4().hex
    upload_folder = os.path.join(BASE_UPLOAD_FOLDER, session_id)
    output_folder = os.path.join(BASE_OUTPUT_FOLDER, session_id)
    
    # Register this session as active before creating directories
    register_active_session(session_id)
    
    try:
        os.makedirs(upload_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)

        original_name = file.filename
        file_path = os.path.join(upload_folder, original_name)
        pptx_name = os.path.splitext(original_name)[0] + ".pptx"
        pptx_path = os.path.join(output_folder, pptx_name)

        file.save(file_path)
        logger.info(f"📄 Saved uploaded file: {file_path}")

        # Process the file
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".pdf":
            content = extract_pdf_content(file_path)
        elif ext in [".docx", ".doc"]:
            content = extract_docx_content(file_path)
        else:
            unregister_active_session(session_id)
            return "Unsupported file type. Please upload PDF or DOCX.", 400

        slide_data = generate_slide_data(content)
        prs = create_presentation_from_json(slide_data)
        prs.save(pptx_path)
        logger.info(f"💾 Saved presentation: {pptx_path}")
        
        # Unregister the session now that processing is complete
        unregister_active_session(session_id)
        
        return redirect(url_for("result_page", session_id=session_id, filename=pptx_name))
    except Exception as e:
        logger.error(f"⚠️ Error processing file: {e}")
        unregister_active_session(session_id)
        return f"Error during processing: {str(e)}", 500

@app.route("/result")
def result_page():
    session_id = request.args.get("session_id")
    filename = request.args.get("filename")
    
    # Check if the file exists first
    file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)
    if not os.path.exists(file_path):
        return render_template_string("""
            <!DOCTYPE html>
            <html>
            <head>
                <title>File Not Found</title>
                <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
            </head>
            <body class="bg-light d-flex flex-column justify-content-center align-items-center" style="height: 100vh;">
                <div class="text-center">
                    <h2 class="text-danger">⚠️ File Not Found</h2>
                    <p class="lead">The requested file could not be located.</p>
                    <p>You can try uploading your file again.</p>
                    <a href="/" class="btn btn-primary">Go Back to Upload</a>
                </div>
            </body>
            </html>
        """), 404
    
    # Register session as active while viewing results
    register_active_session(session_id)
    
    return render_template("result.html", session_id=session_id, filename=filename)

@app.route("/download")
def download_file():
    session_id = request.args.get("session_id")
    filename = request.args.get("filename")
    file_path = os.path.join(BASE_OUTPUT_FOLDER, session_id, filename)

    # Register this session as active before download
    register_active_session(session_id)

    if not os.path.exists(file_path):
        # Unregister session if file doesn't exist
        unregister_active_session(session_id)
        
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
                    <p>You can re-upload your file to generate the presentation again.</p>
                    <a href="/" class="btn btn-primary">Go Back to Upload</a>
                </div>
            </body>
            </html>
        """), 404

    try:
        # We'll unregister after the download starts
        response = send_file(file_path, as_attachment=True)
        
        # Setup response callback to unregister session after download starts
        @response.call_on_close
        def on_close():
            unregister_active_session(session_id)
            
        return response
    except Exception as e:
        unregister_active_session(session_id)
        logger.error(f"⚠️ Error sending file {file_path}: {e}")
        return "Error serving file", 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    start_background_cleanup()
    app.run(host="0.0.0.0", port=port, debug=False)
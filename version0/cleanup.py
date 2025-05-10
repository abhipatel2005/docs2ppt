# cleanup.py
import os
from datetime import datetime, timedelta

BASE_FOLDERS = ["uploads", "output"]
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
            # Step 1: Delete old files
            for root, dirs, files in os.walk(folder_path):
                for filename in files:
                    file_path = os.path.join(root, filename)
                    if is_file_expired(file_path, expiration_minutes):
                        os.remove(file_path)
                        print(f"🗑️ Deleted file: {file_path}")

            # Step 2: Remove empty folders
            for root, dirs, files in os.walk(folder_path, topdown=False):
                if not os.listdir(root):  # empty directory
                    os.rmdir(root)
                    print(f"📁 Deleted empty folder: {root}")

        except Exception as e:
            print(f"⚠️ Error cleaning {folder_path}: {e}")

if __name__ == "__main__":
    for folder in BASE_FOLDERS:
        cleanup_old_files_and_folders(folder, EXPIRATION_MINUTES)

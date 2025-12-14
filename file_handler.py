import os
import tempfile
from werkzeug.utils import secure_filename

# You can adjust this per converter type if you want

ALLOWED_EXTENSIONS = {
    # Image types
    '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif',
    '.webp', '.heif', '.heic',

    # Document types
    '.pdf', '.docx', '.xlsx', '.pptx',

    # === ADD THESE ===
    '.zip',
    '.csv',
    '.json'
}


MAX_FILE_SIZE_MB = 50  # prevent excessively large uploads


def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def save_uploaded_files(request):
    """
    Saves uploaded files from Flask request.files['images'] (or other fields)
    into a temporary directory.

    Returns:
        tuple: (temp_dir, saved_filepaths)
    """
    temp_dir = tempfile.mkdtemp()
    saved_filepaths = []

    uploaded_files = []
    # Try multiple common field names
    for key in ('images', 'files', 'pdfs', 'documents'):
        if key in request.files:
            uploaded_files.extend(request.files.getlist(key))

    if not uploaded_files:
        print("No files found in request.")
        return temp_dir, []

    for file in uploaded_files:
        if not file or not file.filename:
            continue

        filename = secure_filename(file.filename)

        if not allowed_file(filename):
            print(f"Skipped unsupported file: {filename}")
            continue

        # Limit file size before writing to disk
        file.seek(0, os.SEEK_END)
        size_mb = file.tell() / (1024 * 1024)
        file.seek(0)
        if size_mb > MAX_FILE_SIZE_MB:
            print(f"File too large ({size_mb:.2f} MB): {filename}")
            continue

        full_path = os.path.join(temp_dir, filename)
        try:
            file.save(full_path)
            saved_filepaths.append(full_path)
            print(f"✅ Saved: {full_path}")
        except Exception as e:
            print(f"⚠️ Could not save file {filename}: {e}")

    return temp_dir, saved_filepaths

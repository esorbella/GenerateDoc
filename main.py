import os
import shutil
import tempfile
from converter import converter
from flask import Flask, request, send_file, after_this_request
from zipfile import ZipFile

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_files = request.files.getlist('file')
    if not uploaded_files or uploaded_files[0].filename == '':
        return "No files uploaded", 400

    # Create a temp directory to store converted files
    temp_dir = tempfile.mkdtemp()
    output_files = []

    for uploaded_file in uploaded_files:
        file_contents = uploaded_file.read()
        # 👇 Pass original filename for context, if useful
        output_path = converter(file_contents)
        output_files.append(output_path)

    # Create zip file
    zip_path = os.path.join(temp_dir, "converted_files.zip")
    with ZipFile(zip_path, 'w') as zipf:
        for file_path in output_files:
            zipf.write(file_path, arcname=os.path.basename(file_path))

    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            app.logger.error(f"Error cleaning up temp files: {e}")
        return response

    return send_file(
        zip_path,
        as_attachment=True,
        download_name="converted_files.zip"
    )

if __name__ == '__main__':
    app.run(debug=True)
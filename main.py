from flask import Flask, request, send_file, render_template, after_this_request
from converter import converter
import io
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_file = request.files['file']
    if not uploaded_file:
        return "No file uploaded", 400

    # Read file content (example: assume text file for this demo)
    file_contents = uploaded_file.read()

    # 👇 Here's your "converter" line — do actual processing here
    #processed_contents = file_contents.upper()  # EXAMPLE: convert text to uppercase
    converter(file_contents)
    # Return the processed file

    filepath = "generated_schedule.docx"

    @after_this_request
    def remove_file(response):
        try:
            os.remove(filepath)
        except Exception as e:
            app.logger.error(f"Error deleting file: {e}")
        return response

    return send_file(
        filepath,
        as_attachment=True,
        download_name="generated_schedule.docx"
    )

if __name__ == '__main__':
    app.run(debug=True)

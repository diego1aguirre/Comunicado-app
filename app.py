import os
import uuid
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
from processor import process_comunicado

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max

UPLOAD_FOLDER = '/tmp/comunicado_uploads'
OUTPUT_FOLDER = '/tmp/comunicado_outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'No file selected'}), 400
    if not file.filename.lower().endswith('.docx'):
        return jsonify({'error': 'Only .docx files are supported'}), 400

    # Save uploaded file
    uid = uuid.uuid4().hex
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}_input.docx')
    output_path = os.path.join(OUTPUT_FOLDER, f'{uid}_plain.docx')
    file.save(input_path)

    try:
        output_filename = process_comunicado(input_path, output_path)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


if __name__ == '__main__':
    app.run(debug=True, port=8080)

import os
import uuid
import zipfile
import subprocess
import shutil
from flask import Flask, request, render_template, send_file, jsonify
from processor import process_comunicado

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max

UPLOAD_FOLDER = '/tmp/comunicado_uploads'
OUTPUT_FOLDER = '/tmp/comunicado_outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


_SOFFICE_CANDIDATES = [
    'soffice',
    'libreoffice',
    '/usr/bin/soffice',
    '/usr/lib/libreoffice/program/soffice',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',  # macOS
]

def _find_soffice():
    """Return the path to the soffice/libreoffice binary, or raise."""
    # Allow explicit override via env var
    override = os.environ.get('SOFFICE_PATH')
    if override:
        return override
    for candidate in _SOFFICE_CANDIDATES:
        path = shutil.which(candidate) or (candidate if os.path.isfile(candidate) else None)
        if path:
            return path
    raise FileNotFoundError(
        'LibreOffice not found. Install it or set the SOFFICE_PATH environment variable.'
    )


def _convert_to_pdf(docx_path, out_dir):
    """Convert a .docx to PDF using LibreOffice headless and return the PDF path."""
    soffice = _find_soffice()
    subprocess.run(
        [soffice, '--headless', '--convert-to', 'pdf', '--outdir', out_dir, docx_path],
        check=True,
        capture_output=True,
    )
    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    return os.path.join(out_dir, pdf_name)


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

    want_plain = request.form.get('plain') == 'true'
    want_pdf   = request.form.get('pdf')   == 'true'

    if not want_plain and not want_pdf:
        return jsonify({'error': 'Selecciona al menos una salida'}), 400

    # Save uploaded file
    uid = uuid.uuid4().hex
    work_dir = os.path.join(OUTPUT_FOLDER, uid)
    os.makedirs(work_dir, exist_ok=True)

    input_path  = os.path.join(UPLOAD_FOLDER, f'{uid}_input.docx')
    docx_output = os.path.join(work_dir, f'{uid}_plain.docx')
    file.save(input_path)

    try:
        docx_filename = process_comunicado(input_path, docx_output)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

    final_docx = os.path.join(work_dir, docx_filename)
    os.rename(docx_output, final_docx)

    # ── docx only ──────────────────────────────────────────────────────────
    if want_plain and not want_pdf:
        return send_file(
            final_docx,
            as_attachment=True,
            download_name=docx_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    # ── PDF needed — convert first ─────────────────────────────────────────
    try:
        pdf_path = _convert_to_pdf(final_docx, work_dir)
    except Exception as e:
        return jsonify({'error': f'PDF conversion failed: {e}'}), 500

    pdf_filename = os.path.splitext(docx_filename)[0] + '.pdf'

    # ── pdf only ───────────────────────────────────────────────────────────
    if want_pdf and not want_plain:
        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=pdf_filename,
            mimetype='application/pdf'
        )

    # ── both — return a ZIP ────────────────────────────────────────────────
    zip_filename = os.path.splitext(docx_filename)[0] + '.zip'
    zip_path = os.path.join(work_dir, zip_filename)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(final_docx, docx_filename)
        zf.write(pdf_path, pdf_filename)

    return send_file(
        zip_path,
        as_attachment=True,
        download_name=zip_filename,
        mimetype='application/zip'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(debug=False, host='0.0.0.0', port=port)

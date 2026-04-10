import os
import re
import uuid
import subprocess
import shutil
from flask import Flask, request, render_template, send_file, jsonify
from processor import process_comunicado

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max
app.config['TEMPLATES_AUTO_RELOAD'] = True

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

    # Optional custom filename (sanitized: keep letters, digits, spaces, hyphens, underscores)
    custom_name = request.form.get('output_name', '').strip()
    custom_name = re.sub(r'[^\w\s\-]', '', custom_name, flags=re.UNICODE).strip()
    custom_name = custom_name[:80] or None  # cap length; None = use auto name

    if not want_plain and not want_pdf:
        return jsonify({'error': 'Selecciona al menos una salida'}), 400

    uid = uuid.uuid4().hex
    work_dir = os.path.join(OUTPUT_FOLDER, uid)
    os.makedirs(work_dir, exist_ok=True)

    # Save the original upload — used for PDF conversion as-is
    original_path = os.path.join(work_dir, 'original.docx')
    file.save(original_path)

    # ── Plain .docx (reformatted) ───────────────────────────────────────────
    final_docx, docx_filename = None, None
    if want_plain:
        docx_output = os.path.join(work_dir, 'plain.docx')
        try:
            docx_filename = process_comunicado(original_path, docx_output)
        except Exception as e:
            return jsonify({'error': str(e)}), 500
        if custom_name:
            docx_filename = custom_name + '.docx'
        final_docx = os.path.join(work_dir, docx_filename)
        os.rename(docx_output, final_docx)

    # ── PDF (original input converted directly) ────────────────────────────
    pdf_path, pdf_filename = None, None
    if want_pdf:
        try:
            pdf_path = _convert_to_pdf(original_path, work_dir)
        except Exception as e:
            return jsonify({'error': f'PDF conversion failed: {e}'}), 500
        # Name the PDF: custom > plain docx base > upload filename
        if custom_name:
            base_name = custom_name
        elif docx_filename:
            base_name = os.path.splitext(docx_filename)[0]
        else:
            base_name = os.path.splitext(file.filename)[0]
        pdf_filename = base_name + '.pdf'
        os.rename(pdf_path, os.path.join(work_dir, pdf_filename))
        pdf_path = os.path.join(work_dir, pdf_filename)

    # ── Return ─────────────────────────────────────────────────────────────
    if want_plain:
        return send_file(final_docx, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    return send_file(pdf_path, as_attachment=True, download_name=pdf_filename,
                     mimetype='application/pdf')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(debug=False, host='0.0.0.0', port=port)

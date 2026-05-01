"""
DOCX Image Extractor Service
Extracts text + images from .docx files and returns them with position markers.
"""

from flask import Flask, request, jsonify
from docx import Document
from docx.oxml.ns import qn
import zipfile
import base64
import io
import os
import re

app = Flask(__name__)


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'docx-image-extractor'})


@app.route('/extract', methods=['POST'])
def extract_docx():
    """
    POST /extract
    Accepts EITHER:
      - multipart/form-data with a file field named "file"
      - JSON body: { "file_base64": "...", "filename": "book.docx" }
    Returns: {
        "text_with_markers": "...",
        "images": [...],
        "image_count": N,
        ...
    }
    """
    try:
        docx_bytes = None
        filename = 'unknown.docx'

        # Method 1: multipart file upload (preferred for n8n)
        if 'file' in request.files:
            uploaded_file = request.files['file']
            docx_bytes = uploaded_file.read()
            filename = uploaded_file.filename or 'upload.docx'

        # Method 2: JSON with base64 (kept for compatibility)
        elif request.is_json:
            data = request.json
            if not data or 'file_base64' not in data:
                return jsonify({'error': 'Missing file_base64 or file upload'}), 400
            docx_b64 = data['file_base64']
            filename = data.get('filename', 'unknown.docx')
            try:
                docx_bytes = base64.b64decode(docx_b64)
            except Exception as e:
                return jsonify({'error': f'Invalid base64: {str(e)}'}), 400

        else:
            return jsonify({
                'error': 'No file provided. Send as multipart/form-data with field "file", or as JSON with field "file_base64".'
            }), 400

        if not docx_bytes or len(docx_bytes) < 100:
            return jsonify({
                'error': f'File is empty or too small ({len(docx_bytes) if docx_bytes else 0} bytes)'
            }), 400

        docx_io = io.BytesIO(docx_bytes)

        # ------------------------------------------------------------
        # STEP 1: Extract images from the ZIP structure
        # ------------------------------------------------------------
        images = []
        image_filenames_in_order = []

        try:
            with zipfile.ZipFile(docx_io, 'r') as z:
                media_files = sorted([
                    f for f in z.namelist()
                    if f.startswith('word/media/')
                ])

                for media_path in media_files:
                    image_data = z.read(media_path)
                    image_name = os.path.basename(media_path)

                    images.append({
                        'name': image_name,
                        'original_path': media_path,
                        'data_base64': base64.b64encode(image_data).decode('utf-8'),
                        'mime_type': guess_mime(image_name),
                        'size_bytes': len(image_data)
                    })
                    image_filenames_in_order.append(image_name)
        except zipfile.BadZipFile:
            preview = docx_bytes[:50] if docx_bytes else b''
            return jsonify({
                'error': f'File is not a valid .docx (not a ZIP). Received {len(docx_bytes)} bytes. Preview: {preview}'
            }), 400

        # ------------------------------------------------------------
        # STEP 2: Walk through document and insert image markers
        # ------------------------------------------------------------
        docx_io.seek(0)
        doc = Document(docx_io)

        text_parts = []
        image_counter = 0
        marker_to_image = {}

        for para in doc.paragraphs:
            para_has_image = False

            for run in para.runs:
                drawings = run.element.findall('.//' + qn('w:drawing'))
                for _ in drawings:
                    image_counter += 1
                    marker = f'[IMAGE_{image_counter:03d}]'
                    text_parts.append(marker)
                    if image_counter <= len(image_filenames_in_order):
                        marker_to_image[marker] = image_filenames_in_order[image_counter - 1]
                    para_has_image = True

            if para.text.strip():
                text_parts.append(para.text)
            elif not para_has_image:
                text_parts.append('')

        full_text = '\n\n'.join(text_parts)
        full_text = re.sub(r'\n{3,}', '\n\n', full_text).strip()

        return jsonify({
            'text_with_markers': full_text,
            'images': images,
            'image_count': image_counter,
            'marker_to_image': marker_to_image,
            'filename': filename,
            'received_bytes': len(docx_bytes),
            'success': True
        })

    except Exception as e:
        import traceback
        return jsonify({
            'error': str(e),
            'traceback': traceback.format_exc(),
            'success': False
        }), 500


def guess_mime(filename):
    """Guess MIME type from file extension."""
    ext = filename.lower().rsplit('.', 1)[-1]
    return {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'webp': 'image/webp',
        'bmp': 'image/bmp',
        'tiff': 'image/tiff',
        'tif': 'image/tiff',
    }.get(ext, 'application/octet-stream')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)

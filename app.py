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
    Body: { "file_base64": "...", "filename": "book.docx" }
    Returns: { 
        "text_with_markers": "...",
        "images": [{"name": "image1.png", "data_base64": "...", "mime_type": "image/png"}],
        "image_count": N
    }
    """
    try:
        data = request.json
        if not data or 'file_base64' not in data:
            return jsonify({'error': 'Missing file_base64'}), 400

        docx_b64 = data['file_base64']
        filename = data.get('filename', 'unknown.docx')

        # Decode the file
        try:
            docx_bytes = base64.b64decode(docx_b64)
        except Exception as e:
            return jsonify({'error': f'Invalid base64: {str(e)}'}), 400

        docx_io = io.BytesIO(docx_bytes)

        # ------------------------------------------------------------
        # STEP 1: Extract images from the ZIP structure
        # ------------------------------------------------------------
        images = []
        image_filenames_in_order = []

        with zipfile.ZipFile(docx_io, 'r') as z:
            # Get all media files
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

        # ------------------------------------------------------------
        # STEP 2: Walk through document and insert image markers
        # ------------------------------------------------------------
        docx_io.seek(0)
        doc = Document(docx_io)

        text_parts = []
        image_counter = 0
        marker_to_image = {}

        # Walk paragraphs and insert markers where images appear
        for para in doc.paragraphs:
            para_has_image = False

            # Check for inline images (drawings)
            for run in para.runs:
                drawings = run.element.findall('.//' + qn('w:drawing'))
                for _ in drawings:
                    image_counter += 1
                    marker = f'[IMAGE_{image_counter:03d}]'
                    text_parts.append(marker)
                    if image_counter <= len(image_filenames_in_order):
                        marker_to_image[marker] = image_filenames_in_order[image_counter - 1]
                    para_has_image = True

            # Add paragraph text
            if para.text.strip():
                text_parts.append(para.text)
            elif not para_has_image:
                text_parts.append('')

        full_text = '\n\n'.join(text_parts)

        # Clean up: collapse 3+ consecutive blank lines into 2
        full_text = re.sub(r'\n{3,}', '\n\n', full_text).strip()

        return jsonify({
            'text_with_markers': full_text,
            'images': images,
            'image_count': image_counter,
            'marker_to_image': marker_to_image,
            'filename': filename,
            'success': True
        })

    except Exception as e:
        return jsonify({
            'error': str(e),
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

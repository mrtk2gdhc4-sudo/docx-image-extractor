import os
import io
import re
import base64
import zipfile
import gc
from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.oxml.ns import qn

app = Flask(__name__)

# --- CONFIGURATION ---
BATCH_SIZE = 15  # Reduced to prevent OpenAI timeout and memory spikes
MAX_TOKENS_PER_BATCH = 3000 
_openai_client = None

def get_openai_client():
    global _openai_client
    if _openai_client is None:
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            raise RuntimeError('OPENAI_API_KEY not set')
        from openai import OpenAI
        _openai_client = OpenAI(api_key=api_key)
    return _openai_client

# --- HELPERS ---

def update_paragraph_text(para, new_text):
    """Updates paragraph text while preserving images/drawings."""
    text_only_runs = [run for run in para.runs if not run.element.findall('.//' + qn('w:drawing'))]
    if not text_only_runs:
        return
    for run in text_only_runs:
        run.text = ''
    text_only_runs[0].text = new_text

def edit_paragraphs_batch(batch_items):
    """Sends a smaller batch to OpenAI."""
    if not batch_items:
        return {}

    lines = [f"[{idx}] {text.replace('\n', ' ')}" for idx, text in batch_items]
    user_message = '\n'.join(lines)

    try:
        response = get_openai_client().chat.completions.create(
            model='gpt-4o',
            temperature=0.3,
            messages=[
                {'role': 'system', 'content': EDIT_SYSTEM_PROMPT},
                {'role': 'user', 'content': user_message},
            ],
        )
        edited_content = response.choices[0].message.content.strip()
        
        results = {}
        for line in edited_content.split('\n'):
            match = re.match(r'^\[(\d+)\]\s*(.*)$', line.strip())
            if match:
                results[int(match.group(1))] = match.group(2)
        return results
    except Exception as e:
        print(f"Batch processing error: {e}")
        return {}

# --- PROMPTS ---
EDIT_SYSTEM_PROMPT = """Professional copyeditor: Fix grammar/typos. Preserve voice and dialogue. 
Return format: [N] edited text. Do not merge paragraphs."""

# --- ROUTES ---

@app.route('/edit-docx', methods=['POST'])
def edit_docx():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file'}), 400

        uploaded_file = request.files['file']
        filename = uploaded_file.filename
        
        # 1. Load Document into memory once
        doc = Document(io.BytesIO(uploaded_file.read()))
        
        # 2. Filter for actual text to save API costs and memory
        paragraphs_to_edit = []
        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if len(text) > 2: # Ignore empty lines/page breaks
                paragraphs_to_edit.append((idx, text))

        total_to_edit = len(paragraphs_to_edit)
        print(f"Processing {total_to_edit} paragraphs for {filename}")

        # 3. Process in smaller batches with Garbage Collection
        all_edits = {}
        for i in range(0, total_to_edit, BATCH_SIZE):
            batch = paragraphs_to_edit[i : i + BATCH_SIZE]
            edits = edit_paragraphs_batch(batch)
            all_edits.update(edits)
            
            # Forced cleanup to prevent Render memory crashes
            del batch
            gc.collect() 
            print(f"Progress: {min(i + BATCH_SIZE, total_to_edit)}/{total_to_edit}")

        # 4. Apply edits back to the doc object
        applied_count = 0
        for idx, original_text in paragraphs_to_edit:
            if idx in all_edits:
                new_text = all_edits[idx]
                # Safety check: ensure AI didn't hallucinate/delete the whole paragraph
                if len(new_text) > (len(original_text) * 0.3):
                    update_paragraph_text(doc.paragraphs[idx], new_text)
                    applied_count += 1

        # 5. Save and Return
        out_io = io.BytesIO()
        doc.save(out_io)
        out_io.seek(0)
        
        return send_file(
            out_io,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f"edited_{filename}"
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# (Keep your existing /health, /extract, and /detect-trim logic below)

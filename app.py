import os
import io
import re
import base64
import zipfile
import gc
import anthropic  # <-- Updated from openai
from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.oxml.ns import qn

app = Flask(__name__)

# --- CONFIGURATION ---
BATCH_SIZE = 15 
_anthropic_client = None

def get_anthropic_client():
    global _anthropic_client
    if _anthropic_client is None:
        api_key = os.environ.get('ANTHROPIC_API_KEY') # <-- Uses the key we added to Render
        if not api_key:
            raise RuntimeError('ANTHROPIC_API_KEY not set')
        _anthropic_client = anthropic.Anthropic(api_key=api_key)
    return _anthropic_client

# --- HELPERS ---

def update_paragraph_text(para, new_text):
    """Updates paragraph text while preserving images/drawings."""
    text_only_runs = [run for run in para.runs if not run.element.findall('.//' + qn('w:drawing'))]
    if not text_only_runs:
        return
    for run in text_only_runs:
        run.text = ''
    text_only_runs[0].text = new_text

def edit_paragraphs_batch(batch_items, dynamic_system_prompt):
    """Sends a batch to Claude with the custom prompt from n8n."""
    if not batch_items:
        return {}

    lines = [f"[{idx}] {text.replace('\n', ' ')}" for idx, text in batch_items]
    user_message = '\n'.join(lines)

    # Use the prompt from n8n, or fall back to a default if missing
    final_prompt = dynamic_system_prompt or "Professional copyeditor: Fix grammar/typos. Preserve voice. Return format: [N] edited text."

    try:
        # The Claude-specific call
        response = get_anthropic_client().messages.create(
            model="claude-3-5-sonnet-20240620",
            max_tokens=4000,
            temperature=0.3,
            system=final_prompt, # This is where the magic happens
            messages=[
                {'role': 'user', 'content': user_message},
            ],
        )
        edited_content = response.content[0].text.strip()
        
        results = {}
        for line in edited_content.split('\n'):
            match = re.match(r'^\[(\d+)\]\s*(.*)$', line.strip())
            if match:
                results[int(match.group(1))] = match.group(2)
        return results
    except Exception as e:
        print(f"Batch processing error: {e}")
        return {}

# --- ROUTES ---

@app.route('/edit-docx', methods=['POST'])
def edit_docx():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file'}), 400

        # Get the system_prompt sent from n8n
        dynamic_system_prompt = request.form.get('system_prompt')

        uploaded_file = request.files['file']
        filename = uploaded_file.filename
        
        # 1. Load Document
        doc = Document(io.BytesIO(uploaded_file.read()))
        
        # 2. Filter text
        paragraphs_to_edit = []
        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if len(text) > 2: 
                paragraphs_to_edit.append((idx, text))

        total_to_edit = len(paragraphs_to_edit)
        print(f"Processing {total_to_edit} paragraphs via Claude for {filename}")

        # 3. Process in batches
        all_edits = {}
        for i in range(0, total_to_edit, BATCH_SIZE):
            batch = paragraphs_to_edit[i : i + BATCH_SIZE]
            edits = edit_paragraphs_batch(batch, dynamic_system_prompt)
            all_edits.update(edits)
            
            del batch
            gc.collect() 
            print(f"Progress: {min(i + BATCH_SIZE, total_to_edit)}/{total_to_edit}")

        # 4. Apply edits
        applied_count = 0
        for idx, original_text in paragraphs_to_edit:
            if idx in all_edits:
                new_text = all_edits[idx]
                if len(new_text) > (len(original_text) * 0.3):
                    update_paragraph_text(doc.paragraphs[idx], new_text)
                    applied_count += 1

        # 5. Return
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

if __name__ == '__main__':
    app.run(debug=True)

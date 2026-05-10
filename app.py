import os
import io
import re
import gc
import anthropic
from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import Inches

app = Flask(__name__)

_anthropic_client = None

def get_anthropic_client():
    global _anthropic_client
    if _anthropic_client is None:
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        _anthropic_client = anthropic.Anthropic(api_key=api_key)
    return _anthropic_client

def apply_formatting(doc, width, height):
    """Automatically sets the physical trim size of the book."""
    for section in doc.sections:
        section.page_width = Inches(float(width))
        section.page_height = Inches(float(height))
        # Standard professional margins
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

def edit_paragraphs_batch(batch_items, dynamic_system_prompt):
    if not batch_items:
        return {}

    lines = [f"[{idx}] {text}" for idx, text in batch_items]
    user_message = "Please edit the following paragraphs:\n" + '\n'.join(lines)

    # We force the AI to be an ACTIVE editor here
    base_instructions = """You are a Master Book Editor. 
    EXPECTATION: You must actively repair the text while keeping the author's soul intact.
    REQUIRED ACTIONS:
    - FIX OCR: Change 'err0rs' to 'errors', 'Th1s' to 'This'.
    - FIX SPACING: Add spaces after punctuation and repair 'smushedtext' into 'smushed text'.
    - FIX REPETITION: Delete accidental double words like 'the the'.
    - NORMALIZE: Ensure lists and dialogue are punctuated professionally.
    - REJECT: Do not preserve obvious technical failures as 'voice'."""

    final_prompt = f"{base_instructions}\n\nCLIENT SPECIFIC RULES:\n{dynamic_system_prompt}"

    try:
        response = get_anthropic_client().messages.create(
            model="claude-3-5-sonnet-20240620",
            max_tokens=4000,
            temperature=0.2, # Lower temperature = more precise editing
            system=final_prompt,
            messages=[{'role': 'user', 'content': user_message}],
        )
        edited_content = response.content[0].text.strip()
        
        results = {}
        for line in edited_content.split('\n'):
            match = re.match(r'^\[(\d+)\]\s*(.*)$', line.strip())
            if match:
                results[int(match.group(1))] = match.group(2)
        return results
    except Exception as e:
        print(f"Error: {e}")
        return {}

@app.route('/edit-docx', methods=['POST'])
def edit_docx():
    try:
        # 1. Get data from n8n
        uploaded_file = request.files['file']
        sys_prompt = request.form.get('system_prompt', "")
        t_width = request.form.get('trim_width', 6)
        t_height = request.form.get('trim_height', 9)

        doc = Document(io.BytesIO(uploaded_file.read()))
        
        # 2. APPLY TRIM SIZE IMMEDIATELY
        apply_formatting(doc, t_width, t_height)

        paragraphs_to_edit = [(i, p.text.strip()) for i, p in enumerate(doc.paragraphs) if len(p.text.strip()) > 1]

        # 3. Process with AI
        all_edits = {}
        batch_size = 10 # Smaller batches for higher quality
        for i in range(0, len(paragraphs_to_edit), batch_size):
            batch = paragraphs_to_edit[i : i + batch_size]
            edits = edit_paragraphs_batch(batch, sys_prompt)
            all_edits.update(edits)
            gc.collect()

        # 4. Apply back to doc (Removed the strict length filter)
        for idx, _ in paragraphs_to_edit:
            if idx in all_edits:
                doc.paragraphs[idx].text = all_edits[idx]

        out_io = io.BytesIO()
        doc.save(out_io)
        out_io.seek(0)
        
        return send_file(out_io, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', as_attachment=True, download_name="edited_book.docx")

    except Exception as e:
        return jsonify({'error': str(e)}), 500

import os
import io
import re
import gc
import anthropic
from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import Inches

app = Flask(__name__)

# Cache the client
_anthropic_client = None

def get_anthropic_client():
    global _anthropic_client
    if _anthropic_client is None:
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        _anthropic_client = anthropic.Anthropic(api_key=api_key)
    return _anthropic_client

def apply_formatting(doc, width, height):
    """Sets physical trim and professional margins."""
    for section in doc.sections:
        section.page_width = Inches(float(width))
        section.page_height = Inches(float(height))
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

def edit_paragraphs_batch(batch_items, dynamic_system_prompt):
    if not batch_items:
        return {}

    lines = [f"[{idx}] {text}" for idx, text in batch_items]
    user_message = "ACT AS AN EDITOR. FIX ALL ERRORS IN THESE PARAGRAPHS:\n" + '\n'.join(lines)

    # FORCE Claude to be aggressive
    system_instruction = (
        "You are an ELITE MANUSCRIPT EDITOR. You have zero tolerance for technical errors.\n"
        "MANDATORY ACTIONS:\n"
        "- FIX OCR: 'Th1s' -> 'This', 'err0rs' -> 'errors'.\n"
        "- FIX REPETITION: Remove double words like 'however however'.\n"
        "- FIX SPACING: Repair 'word,word' and 'smushedtext'.\n"
        "Stay true to the author's story, but CLEAN THE TEXT COMPLETELY.\n"
        f"CONTEXT: {dynamic_system_prompt}\n"
        "Format: [N] edited text."
    )

    try:
        client = get_anthropic_client()
        response = client.messages.create(
            model="claude-3-5-sonnet-20240620",
            max_tokens=4000,
            temperature=0, # 0 = NO CREATIVITY, ONLY ACCURACY
            system=system_instruction,
            messages=[{'role': 'user', 'content': user_message}],
        )
        
        results = {}
        for line in response.content[0].text.strip().split('\n'):
            match = re.match(r'^\[(\d+)\]\s*(.*)$', line.strip())
            if match:
                results[int(match.group(1))] = match.group(2)
        return results
    except Exception as e:
        print(f"Claude Error: {e}")
        return {}

@app.route('/edit-docx', methods=['POST'])
def edit_docx():
    try:
        uploaded_file = request.files['file']
        sys_prompt = request.form.get('system_prompt', "Professional editing.")
        t_width = request.form.get('trim_width', 6)
        t_height = request.form.get('trim_height', 9)

        doc = Document(io.BytesIO(uploaded_file.read()))
        
        # 1. Physical Formatting
        apply_formatting(doc, t_width, t_height)

        # 2. Extract content (Index, Text)
        to_edit = []
        for i, p in enumerate(doc.paragraphs):
            clean_text = p.text.strip()
            if len(clean_text) > 1:
                to_edit.append((i, clean_text))

        # 3. Batch Edit
        batch_size = 8 # Smaller batches = higher attention to detail
        for i in range(0, len(to_edit), batch_size):
            batch = to_edit[i : i + batch_size]
            edits = edit_paragraphs_batch(batch, sys_prompt)
            
            # 4. DIRECT OVERWRITE
            for idx, _ in batch:
                if idx in edits:
                    # This replaces the text while keeping the paragraph object
                    doc.paragraphs[idx].text = edits[idx]
            
            gc.collect()

        # 5. Final Export
        out_io = io.BytesIO()
        doc.save(out_io)
        out_io.seek(0)
        
        return send_file(out_io, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', as_attachment=True, download_name="edited_manuscript.docx")

    except Exception as e:
        print(f"Global Error: {e}")
        return jsonify({'error': str(e)}), 500

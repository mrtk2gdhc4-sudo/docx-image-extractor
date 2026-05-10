import os
import io
import json
import anthropic
from flask import Flask, request, jsonify, send_file
from docx import Document

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Anthropic client
# ---------------------------------------------------------------------------
client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

CHUNK_SIZE = 30  # paragraphs per API call

# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

# ---------------------------------------------------------------------------
# /edit-docx
# Accepts: multipart/form-data
#   file          — the .docx file
#   system_prompt — (optional) custom editing instructions from n8n
# Returns: edited .docx file
# ---------------------------------------------------------------------------
@app.route("/edit-docx", methods=["POST"])
def edit_docx():
    # --- Validate input ---
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files["file"]
    custom_prompt = request.form.get("system_prompt", "").strip()

    # --- Build system prompt ---
    if custom_prompt:
        system_prompt = custom_prompt
    else:
        system_prompt = (
            "You are a professional book editor. "
            "Line edit the following text: fix grammar, punctuation, and awkward phrasing. "
            "Preserve the author's voice. "
            "Return ONLY the edited text. No commentary or explanations. "
            "Preserve all paragraph breaks exactly. "
            "Do not summarize or truncate. Return every word, edited."
        )

    # --- Load docx ---
    try:
        file_bytes = uploaded_file.read()
        doc = Document(io.BytesIO(file_bytes))
    except Exception as e:
        return jsonify({"error": f"Failed to read docx: {str(e)}"}), 400

    # --- Collect paragraphs (preserve images/tables by index) ---
    paragraphs = doc.paragraphs
    total = len(paragraphs)

    # --- Edit in chunks ---
    for start in range(0, total, CHUNK_SIZE):
        chunk = paragraphs[start : start + CHUNK_SIZE]

        # Build text block — skip empty and image-marker-only paragraphs
        texts = []
        indices = []
        for i, para in enumerate(chunk):
            text = para.text.strip()
            if text:
                texts.append(text)
                indices.append(i)

        if not texts:
            continue

        chunk_text = "\n\n".join(texts)

        # --- Call Claude ---
        try:
            message = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=4096,
                system=system_prompt,
                messages=[
                    {
                        "role": "user",
                        "content": (
                            f"Edit the following text. "
                            f"Return ONLY the edited paragraphs separated by blank lines. "
                            f"Same number of paragraphs as input ({len(texts)}). "
                            f"Do not add or remove paragraphs.\n\n"
                            f"{chunk_text}"
                        )
                    }
                ]
            )
            edited_text = message.content[0].text.strip()
        except Exception as e:
            return jsonify({"error": f"Claude API error at chunk {start}: {str(e)}"}), 500

        # --- Write edited text back to paragraphs ---
        edited_paras = [p.strip() for p in edited_text.split("\n\n") if p.strip()]

        for j, idx in enumerate(indices):
            if j < len(edited_paras):
                # Preserve runs formatting by only changing text of first run
                para = chunk[idx]
                if para.runs:
                    # Clear all runs, put edited text in first run
                    edited = edited_paras[j]
                    para.runs[0].text = edited
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.text = edited_paras[j]

    # --- Save and return ---
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="edited.docx"
    )


# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

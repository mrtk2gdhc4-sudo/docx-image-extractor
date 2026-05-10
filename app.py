import os
import io
import anthropic
from flask import Flask, request, jsonify, send_file
from docx import Document

app = Flask(__name__)
client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

CHUNK_SIZE = 20

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/edit-docx", methods=["POST"])
def edit_docx():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files["file"]
    custom_prompt = request.form.get("system_prompt", "").strip()

    if custom_prompt:
        system_prompt = custom_prompt
    else:
        system_prompt = (
            "You are a professional book editor. "
            "Fix grammar, punctuation, and awkward phrasing. "
            "Preserve the author's voice. "
            "Return ONLY the edited text with paragraphs separated by <<<PARA>>>. "
            "Do not summarize or truncate."
        )

    try:
        file_bytes = uploaded_file.read()
        doc = Document(io.BytesIO(file_bytes))
    except Exception as e:
        return jsonify({"error": f"Failed to read docx: {str(e)}"}), 400

    paragraphs = doc.paragraphs
    total = len(paragraphs)

    for start in range(0, total, CHUNK_SIZE):
        chunk = paragraphs[start: start + CHUNK_SIZE]

        # Build indexed list — skip truly empty paragraphs
        texts = []
        indices = []
        for i, para in enumerate(chunk):
            text = para.text.strip()
            if text:
                texts.append(text)
                indices.append(i)

        if not texts:
            continue

        # Join with separator Claude must preserve
        chunk_text = " <<<PARA>>> ".join(texts)

        try:
            message = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=4096,
                system=system_prompt + (
                    "\n\nCRITICAL: The input paragraphs are separated by <<<PARA>>>. "
                    "You MUST return the same number of paragraphs separated by <<<PARA>>>. "
                    "Do NOT merge paragraphs. Do NOT add or remove <<<PARA>>> markers. "
                    "Edit each paragraph individually and return them in order."
                ),
                messages=[{
                    "role": "user",
                    "content": (
                        f"Edit these {len(texts)} paragraphs. "
                        f"Return exactly {len(texts)} edited paragraphs separated by <<<PARA>>>.\n\n"
                        f"{chunk_text}"
                    )
                }]
            )
            edited_text = message.content[0].text.strip()
        except Exception as e:
            return jsonify({"error": f"Claude API error at chunk {start}: {str(e)}"}), 500

        # Split on separator
        edited_paras = [p.strip() for p in edited_text.split("<<<PARA>>>")]

        # Write back
        for j, idx in enumerate(indices):
            if j < len(edited_paras):
                para = chunk[idx]
                new_text = edited_paras[j]
                if para.runs:
                    para.runs[0].text = new_text
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.text = new_text

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="edited.docx"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

import os
import io
import time
import json
import requests
import anthropic
from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.oxml.ns import qn
import copy

app = Flask(__name__)
client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
CLOUDCONVERT_API_KEY = os.environ.get("CLOUDCONVERT_API_KEY")

CHUNK_SIZE = 20

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


# ─── ORIGINAL EDIT ENDPOINT (kept for small books) ──────────────────────────

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
        texts = []
        indices = []
        for i, para in enumerate(chunk):
            text = para.text.strip()
            if text:
                texts.append(text)
                indices.append(i)
        if not texts:
            continue
        chunk_text = " <<<PARA>>> ".join(texts)
        try:
            message = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=4096,
                system=system_prompt + (
                    "\n\nCRITICAL: Input paragraphs are separated by <<<PARA>>>. "
                    "Return same number of paragraphs separated by <<<PARA>>>. "
                    "Do NOT merge paragraphs."
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

        edited_paras = [p.strip() for p in edited_text.split("<<<PARA>>>")]
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


# ─── NEW: EXTRACT PARAGRAPHS ─────────────────────────────────────────────────

@app.route("/extract-paragraphs", methods=["POST"])
def extract_paragraphs():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    try:
        file_bytes = request.files["file"].read()
        doc = Document(io.BytesIO(file_bytes))
    except Exception as e:
        return jsonify({"error": f"Failed to read docx: {str(e)}"}), 400

    paragraphs = []
    for i, para in enumerate(doc.paragraphs):
        paragraphs.append({
            "index": i,
            "text": para.text,
            "style": para.style.name,
            "empty": len(para.text.strip()) == 0
        })

    return jsonify({
        "total": len(paragraphs),
        "paragraphs": paragraphs
    })


# ─── NEW: EDIT BATCH ─────────────────────────────────────────────────────────

@app.route("/edit-batch", methods=["POST"])
def edit_batch():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON body"}), 400

    paragraphs = data.get("paragraphs", [])
    system_prompt = data.get("system_prompt", "You are a professional book editor. Fix grammar and punctuation. Preserve the author's voice.")

    # Only edit non-empty paragraphs
    editable = [p for p in paragraphs if not p.get("empty") and p.get("text", "").strip()]

    if not editable:
        return jsonify({"paragraphs": paragraphs})

    texts = [p["text"] for p in editable]
    chunk_text = " <<<PARA>>> ".join(texts)

    try:
        message = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=4096,
            system=system_prompt + (
                "\n\nCRITICAL: Input paragraphs are separated by <<<PARA>>>. "
                "Return exactly the same number of paragraphs separated by <<<PARA>>>. "
                "Do NOT merge, add, or remove paragraphs."
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
        return jsonify({"error": f"Claude API error: {str(e)}"}), 500

    edited_paras = [p.strip() for p in edited_text.split("<<<PARA>>>")]

    # Map edited text back to original paragraph indices
    result = list(paragraphs)
    for j, para in enumerate(editable):
        if j < len(edited_paras):
            idx = next(i for i, p in enumerate(result) if p["index"] == para["index"])
            result[idx] = {**para, "text": edited_paras[j], "edited": True}

    return jsonify({"paragraphs": result})


# ─── NEW: REBUILD DOCX ───────────────────────────────────────────────────────

@app.route("/rebuild-docx", methods=["POST"])
def rebuild_docx():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    if "paragraphs" not in request.form:
        return jsonify({"error": "No paragraphs data"}), 400

    try:
        file_bytes = request.files["file"].read()
        doc = Document(io.BytesIO(file_bytes))
    except Exception as e:
        return jsonify({"error": f"Failed to read docx: {str(e)}"}), 400

    try:
        edited_paragraphs = json.loads(request.form["paragraphs"])
    except Exception as e:
        return jsonify({"error": f"Invalid paragraphs JSON: {str(e)}"}), 400

    # Build lookup by index
    edited_map = {p["index"]: p["text"] for p in edited_paragraphs if p.get("edited")}

    # Write back to doc
    for i, para in enumerate(doc.paragraphs):
        if i in edited_map:
            new_text = edited_map[i]
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


# ─── CONVERT TO PDF ──────────────────────────────────────────────────────────

@app.route("/convert-to-pdf", methods=["POST"])
def convert_to_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    if not CLOUDCONVERT_API_KEY:
        return jsonify({"error": "CLOUDCONVERT_API_KEY not set"}), 500

    uploaded_file = request.files["file"]
    file_bytes = uploaded_file.read()
    filename = request.form.get("filename", "document.docx")
    page_width = request.form.get("page_width", "6")
    page_height = request.form.get("page_height", "9")

    headers = {
        "Authorization": f"Bearer {CLOUDCONVERT_API_KEY}",
        "Content-Type": "application/json"
    }

    job_payload = {
        "tasks": {
            "import-file": {"operation": "import/upload"},
            "convert-file": {
                "operation": "convert",
                "input": "import-file",
                "output_format": "pdf",
                "engine": "libreoffice",
                "page_width": float(page_width),
                "page_height": float(page_height)
            },
            "export-file": {
                "operation": "export/url",
                "input": "convert-file"
            }
        }
    }

    try:
        job_resp = requests.post(
            "https://api.cloudconvert.com/v2/jobs",
            json=job_payload,
            headers=headers,
            timeout=30
        )
        job_data = job_resp.json()
    except Exception as e:
        return jsonify({"error": f"CloudConvert job creation failed: {str(e)}"}), 500

    if job_resp.status_code != 201:
        return jsonify({"error": "CloudConvert job creation failed", "details": job_data}), 500

    tasks = job_data.get("data", {}).get("tasks", [])
    upload_task = next((t for t in tasks if t.get("name") == "import-file"), None)
    if not upload_task:
        return jsonify({"error": "No upload task found"}), 500

    upload_url = upload_task.get("result", {}).get("form", {}).get("url")
    upload_params = upload_task.get("result", {}).get("form", {}).get("parameters", {})
    if not upload_url:
        return jsonify({"error": "No upload URL from CloudConvert"}), 500

    try:
        upload_resp = requests.post(
            upload_url,
            data=upload_params,
            files={"file": (filename, file_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            timeout=60
        )
    except Exception as e:
        return jsonify({"error": f"File upload failed: {str(e)}"}), 500

    job_id = job_data["data"]["id"]
    pdf_url = None

    for _ in range(30):
        time.sleep(3)
        try:
            status_resp = requests.get(
                f"https://api.cloudconvert.com/v2/jobs/{job_id}",
                headers=headers,
                timeout=15
            )
            status_data = status_resp.json()
            job_status = status_data.get("data", {}).get("status")
            if job_status == "finished":
                export_task = next(
                    (t for t in status_data["data"]["tasks"] if t.get("name") == "export-file"),
                    None
                )
                if export_task:
                    pdf_url = export_task["result"]["files"][0]["url"]
                break
            elif job_status == "error":
                return jsonify({"error": "CloudConvert conversion failed", "details": status_data}), 500
        except Exception as e:
            return jsonify({"error": f"Polling failed: {str(e)}"}), 500

    if not pdf_url:
        return jsonify({"error": "PDF URL not found after conversion"}), 500

    try:
        pdf_resp = requests.get(pdf_url, timeout=60)
        pdf_bytes = pdf_resp.content
    except Exception as e:
        return jsonify({"error": f"Failed to download PDF: {str(e)}"}), 500

    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename.replace(".docx", ".pdf")
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

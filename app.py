import os
import io
import time
import json
import uuid
import threading
import requests
import anthropic
from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
CLOUDCONVERT_API_KEY = os.environ.get("CLOUDCONVERT_API_KEY")

CHUNK_SIZE = 20
jobs = {}

WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'


def para_has_drawing(para):
    for run in para.runs:
        for tag in ['{%s}inline' % WP_NS, '{%s}anchor' % WP_NS]:
            if run._r.find('.//' + tag) is not None:
                return True
    return False


def update_text_safely(para, new_text):
    runs_with_drawings = set()
    for i, run in enumerate(para.runs):
        for tag in ['{%s}inline' % WP_NS, '{%s}anchor' % WP_NS]:
            if run._r.find('.//' + tag) is not None:
                runs_with_drawings.add(i)

    if not runs_with_drawings:
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""
        else:
            para.text = new_text
    else:
        first_done = False
        for i, run in enumerate(para.runs):
            if i not in runs_with_drawings:
                if not first_done:
                    run.text = new_text
                    first_done = True
                else:
                    run.text = ""


def apply_house_style(doc, page_width_inches=6, page_height_inches=9):
    # 1. Normal style
    try:
        normal = doc.styles['Normal']
        normal.font.name = 'Times New Roman'
        normal.font.size = Pt(12)
        normal.paragraph_format.line_spacing = 1.5
        normal.paragraph_format.space_before = Pt(6)
        normal.paragraph_format.space_after = Pt(6)
        normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        normal.paragraph_format.first_line_indent = Inches(0.3)
    except Exception:
        pass

    # 2. Heading styles
    for i in range(1, 4):
        try:
            h = doc.styles[f'Heading {i}']
            h.font.name = 'Times New Roman'
            h.font.size = Pt(18)
            h.font.bold = True
            h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.paragraph_format.space_before = Pt(6)
            h.paragraph_format.space_after = Pt(18)
            h.paragraph_format.first_line_indent = Inches(0)
        except Exception:
            pass

    # 3. Auto-resize images wider than page margins
    EMU_PER_INCH = 914400
    max_width_emu = int((page_width_inches - 1.0) * EMU_PER_INCH)  # 0.5in margin each side

    for para in doc.paragraphs:
        for run in para.runs:
            for drawing_tag in ['{%s}inline' % WP_NS, '{%s}anchor' % WP_NS]:
                drawing = run._r.find('.//' + drawing_tag)
                if drawing is None:
                    continue
                extent = drawing.find('{%s}extent' % WP_NS)
                if extent is None:
                    continue
                try:
                    cx = int(extent.get('cx', 0))
                    cy = int(extent.get('cy', 0))
                except (TypeError, ValueError):
                    continue
                if cx == 0 or cx <= max_width_emu:
                    continue
                scale = max_width_emu / cx
                new_cx = max_width_emu
                new_cy = int(cy * scale)
                extent.set('cx', str(new_cx))
                extent.set('cy', str(new_cy))
                for sp_pr in run._r.findall('.//{%s}spPr' % PIC_NS):
                    xfrm = sp_pr.find('{%s}xfrm' % A_NS)
                    if xfrm is not None:
                        ext = xfrm.find('{%s}ext' % A_NS)
                        if ext is not None:
                            ext.set('cx', str(new_cx))
                            ext.set('cy', str(new_cy))

    # 4. Find body start — skip title page
    body_start = 0
    for i, para in enumerate(doc.paragraphs):
        if len(para.text.strip()) > 100:
            body_start = i
            break

    # 5. Paragraph-level formatting
    for i, para in enumerate(doc.paragraphs):
        if i < body_start:
            continue
        style_name = para.style.name.lower()

        if any(x in style_name for x in ['toc', 'table of', 'index', 'caption', 'header', 'footer', 'vellum']):
            pass
        elif para_has_drawing(para):
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.first_line_indent = Inches(0)
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
        elif 'heading' in style_name:
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(18)
            para.paragraph_format.first_line_indent = Inches(0)
        else:
            para.paragraph_format.line_spacing = 1.5
            para.paragraph_format.first_line_indent = Inches(0.3)
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(6)
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    return doc


def cloudconvert_job(job_payload, file_bytes, filename, mime_type):
    headers = {
        "Authorization": f"Bearer {CLOUDCONVERT_API_KEY}",
        "Content-Type": "application/json"
    }
    job_resp = requests.post(
        "https://api.cloudconvert.com/v2/jobs",
        json=job_payload, headers=headers, timeout=30
    )
    job_data = job_resp.json()
    if job_resp.status_code != 201:
        return None, f"CloudConvert job creation failed: {job_data}"

    tasks = job_data.get("data", {}).get("tasks", [])
    upload_task = next((t for t in tasks if t.get("name") == "import-file"), None)
    if not upload_task:
        return None, "No upload task found"

    upload_url = upload_task.get("result", {}).get("form", {}).get("url")
    upload_params = upload_task.get("result", {}).get("form", {}).get("parameters", {})
    if not upload_url:
        return None, "No upload URL from CloudConvert"

    requests.post(
        upload_url, data=upload_params,
        files={"file": (filename, file_bytes, mime_type)},
        timeout=60
    )

    job_id = job_data["data"]["id"]
    for _ in range(60):
        time.sleep(5)
        status_resp = requests.get(
            f"https://api.cloudconvert.com/v2/jobs/{job_id}",
            headers=headers, timeout=15
        )
        status_data = status_resp.json()
        job_status = status_data.get("data", {}).get("status")
        if job_status == "finished":
            export_task = next(
                (t for t in status_data["data"]["tasks"] if t.get("name") == "export-file"), None
            )
            if export_task:
                return export_task["result"]["files"][0]["url"], None
            return None, "No export task found"
        elif job_status == "error":
            return None, f"CloudConvert error: {status_data}"

    return None, "Timeout waiting for CloudConvert"


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/format-docx", methods=["POST"])
def format_docx():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        file_bytes = request.files["file"].read()
        doc = Document(io.BytesIO(file_bytes))
    except Exception as e:
        return jsonify({"error": f"Failed to read docx: {str(e)}"}), 400
    doc = apply_house_style(doc)
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="formatted.docx"
    )


@app.route("/convert-pdf-to-docx", methods=["POST"])
def convert_pdf_to_docx():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    if not CLOUDCONVERT_API_KEY:
        return jsonify({"error": "CLOUDCONVERT_API_KEY not set"}), 500

    file_bytes = request.files["file"].read()
    filename = request.form.get("filename", "document.pdf")

    job_payload = {
        "tasks": {
            "import-file": {"operation": "import/upload"},
            "convert-file": {
                "operation": "convert",
                "input": "import-file",
                "output_format": "docx"
            },
            "export-file": {
                "operation": "export/url",
                "input": "convert-file"
            }
        }
    }

    docx_url, error = cloudconvert_job(job_payload, file_bytes, filename, "application/pdf")
    if error:
        return jsonify({"error": error}), 500

    try:
        docx_resp = requests.get(docx_url, timeout=60)
        docx_bytes = docx_resp.content
    except Exception as e:
        return jsonify({"error": f"Failed to download docx: {str(e)}"}), 500

    return send_file(
        io.BytesIO(docx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename.replace(".pdf", ".docx")
    )


@app.route("/resize-pdf", methods=["POST"])
def resize_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    if not CLOUDCONVERT_API_KEY:
        return jsonify({"error": "CLOUDCONVERT_API_KEY not set"}), 500

    file_bytes = request.files["file"].read()
    filename = request.form.get("filename", "document.pdf")
    page_width = request.form.get("page_width", "6")
    page_height = request.form.get("page_height", "9")

    job_payload = {
        "tasks": {
            "import-file": {"operation": "import/upload"},
            "convert-file": {
                "operation": "convert",
                "input": "import-file",
                "output_format": "pdf",
                "engine": "ghostscript",
                "page_width": float(page_width),
                "page_height": float(page_height)
            },
            "export-file": {
                "operation": "export/url",
                "input": "convert-file"
            }
        }
    }

    pdf_url, error = cloudconvert_job(job_payload, file_bytes, filename, "application/pdf")
    if error:
        return jsonify({"error": error}), 500

    try:
        pdf_resp = requests.get(pdf_url, timeout=60)
        pdf_bytes = pdf_resp.content
    except Exception as e:
        return jsonify({"error": f"Failed to download PDF: {str(e)}"}), 500

    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename
    )


@app.route("/edit-docx", methods=["POST"])
def edit_docx():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files["file"]
    custom_prompt = request.form.get("system_prompt", "").strip()

    system_prompt = custom_prompt if custom_prompt else (
        "You are a professional book editor. Fix grammar, punctuation, and awkward phrasing. "
        "Preserve the author's voice. Return ONLY the edited text with paragraphs separated by <<<PARA>>>. "
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
        texts, indices = [], []
        for i, para in enumerate(chunk):
            if para_has_drawing(para):
                continue
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
                messages=[{"role": "user", "content": (
                    f"Edit these {len(texts)} paragraphs. "
                    f"Return exactly {len(texts)} edited paragraphs separated by <<<PARA>>>.\n\n"
                    f"{chunk_text}"
                )}]
            )
            edited_text = message.content[0].text.strip()
        except Exception as e:
            return jsonify({"error": f"Claude API error at chunk {start}: {str(e)}"}), 500

        edited_paras = [p.strip() for p in edited_text.split("<<<PARA>>>") if p.strip()]
        for j, idx in enumerate(indices):
            if j < len(edited_paras):
                update_text_safely(chunk[idx], edited_paras[j])

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="edited.docx"
    )


def process_job(job_id, file_bytes, system_prompt):
    jobs[job_id]["status"] = "processing"
    try:
        doc = Document(io.BytesIO(file_bytes))
        paragraphs = doc.paragraphs
        total = len(paragraphs)
        jobs[job_id]["total_chunks"] = (total // CHUNK_SIZE) + 1
        jobs[job_id]["completed_chunks"] = 0

        for start in range(0, total, CHUNK_SIZE):
            chunk = paragraphs[start: start + CHUNK_SIZE]
            texts, indices = [], []
            for i, para in enumerate(chunk):
                if para_has_drawing(para):
                    continue
                text = para.text.strip()
                if text:
                    texts.append(text)
                    indices.append(i)

            if not texts:
                jobs[job_id]["completed_chunks"] += 1
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
                    messages=[{"role": "user", "content": (
                        f"Edit these {len(texts)} paragraphs. "
                        f"Return exactly {len(texts)} edited paragraphs separated by <<<PARA>>>.\n\n"
                        f"{chunk_text}"
                    )}]
                )
                edited_text = message.content[0].text.strip()
                if not edited_text or len(edited_text) < 10:
                    jobs[job_id]["completed_chunks"] += 1
                    continue
            except Exception:
                jobs[job_id]["completed_chunks"] += 1
                continue

            edited_paras = [p.strip() for p in edited_text.split("<<<PARA>>>") if p.strip()]
            for j, idx in enumerate(indices):
                if j < len(edited_paras):
                    update_text_safely(chunk[idx], edited_paras[j])

            jobs[job_id]["completed_chunks"] += 1

        output = io.BytesIO()
        doc.save(output)
        jobs[job_id]["result"] = output.getvalue()
        jobs[job_id]["status"] = "done"

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(e)


@app.route("/edit-docx-async", methods=["POST"])
def edit_docx_async():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file_bytes = request.files["file"].read()
    system_prompt = request.form.get("system_prompt", "").strip()

    if not system_prompt:
        system_prompt = (
            "You are a professional book editor. Fix grammar, punctuation, and awkward phrasing. "
            "Preserve the author's voice. Return ONLY the edited text with paragraphs separated by <<<PARA>>>. "
            "Do not summarize or truncate."
        )

    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "status": "queued",
        "created_at": time.time(),
        "total_chunks": 0,
        "completed_chunks": 0,
        "result": None,
        "error": None
    }

    thread = threading.Thread(target=process_job, args=(job_id, file_bytes, system_prompt))
    thread.daemon = True
    thread.start()

    return jsonify({"job_id": job_id, "status": "queued"})


@app.route("/job-status/<job_id>", methods=["GET"])
def job_status(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    return jsonify({
        "job_id": job_id,
        "status": job["status"],
        "total_chunks": job["total_chunks"],
        "completed_chunks": job["completed_chunks"],
        "error": job.get("error")
    })


@app.route("/job-result/<job_id>", methods=["GET"])
def job_result(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "done":
        return jsonify({"error": f"Job not done yet — status: {job['status']}"}), 400
    result_bytes = job["result"]
    del jobs[job_id]
    return send_file(
        io.BytesIO(result_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="edited.docx"
    )


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
    return jsonify({"total": len(paragraphs), "paragraphs": paragraphs})


@app.route("/edit-batch", methods=["POST"])
def edit_batch():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON body"}), 400

    paragraphs = data.get("paragraphs", [])
    system_prompt = data.get("system_prompt", "You are a professional book editor.")
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
            messages=[{"role": "user", "content": (
                f"Edit these {len(texts)} paragraphs. "
                f"Return exactly {len(texts)} edited paragraphs separated by <<<PARA>>>.\n\n"
                f"{chunk_text}"
            )}]
        )
        edited_text = message.content[0].text.strip()
    except Exception as e:
        return jsonify({"error": f"Claude API error: {str(e)}"}), 500

    edited_paras = [p.strip() for p in edited_text.split("<<<PARA>>>") if p.strip()]
    result = list(paragraphs)
    for j, para in enumerate(editable):
        if j < len(edited_paras):
            idx = next(i for i, p in enumerate(result) if p["index"] == para["index"])
            result[idx] = {**para, "text": edited_paras[j], "edited": True}

    return jsonify({"paragraphs": result})


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

    edited_map = {p["index"]: p["text"] for p in edited_paragraphs if p.get("edited")}
    for i, para in enumerate(doc.paragraphs):
        if i in edited_map:
            if para_has_drawing(para):
                continue
            update_text_safely(para, edited_map[i])

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="edited.docx"
    )


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

    try:
        doc = Document(io.BytesIO(file_bytes))
        doc = apply_house_style(doc, page_width_inches=float(page_width), page_height_inches=float(page_height))
        styled_output = io.BytesIO()
        doc.save(styled_output)
        file_bytes = styled_output.getvalue()
    except Exception:
        pass

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

    pdf_url, error = cloudconvert_job(
        job_payload, file_bytes, filename,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    if error:
        return jsonify({"error": error}), 500

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

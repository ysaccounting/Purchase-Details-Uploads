import io
import os
import json
import zipfile
import threading
import uuid
from flask import Flask, request, jsonify, send_file, render_template
from processor import process_file

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# In-memory job store (fine for single-worker Railway deploy)
jobs = {}
jobs_lock = threading.Lock()


def run_job(job_id, file_bytes):
    try:
        result = process_file(file_bytes)
        with jobs_lock:
            jobs[job_id] = {"status": "done", "result": result}
    except Exception as e:
        with jobs_lock:
            jobs[job_id] = {"status": "error", "message": str(e)}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    file_bytes = f.read()

    job_id = str(uuid.uuid4())
    with jobs_lock:
        jobs[job_id] = {"status": "processing"}

    thread = threading.Thread(target=run_job, args=(job_id, file_bytes), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"status": "not_found"}), 404
    if job["status"] == "error":
        return jsonify({"status": "error", "message": job["message"]})
    if job["status"] == "processing":
        return jsonify({"status": "processing"})

    result = job["result"]
    date_range = result["date_range"]
    companies = list(result["companies"].keys())
    return jsonify({
        "status": "done",
        "date_range": date_range,
        "companies": companies,
    })


@app.route("/download/<job_id>/combined")
def download_combined(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Not ready"}), 404
    result = job["result"]
    date_range = result["date_range"]
    filename = f"Purchase Details - {date_range}.xlsx"
    return send_file(
        io.BytesIO(result["combined"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/download/<job_id>/company/<company>")
def download_company(job_id, company):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Not ready"}), 404
    result = job["result"]
    file_bytes = result["companies"].get(company)
    if file_bytes is None:
        return jsonify({"error": "Company not found"}), 404
    date_range = result["date_range"]
    filename = f"Purchase Details - {company} - {date_range}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/download/<job_id>/all")
def download_all_zip(job_id):
    """Download all files (combined + per-company) as a single zip."""
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "Not ready"}), 404
    result = job["result"]
    date_range = result["date_range"]

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"Purchase Details - {date_range}.xlsx", result["combined"])
        for company, file_bytes in result["companies"].items():
            zf.writestr(f"Purchase Details - {company} - {date_range}.xlsx", file_bytes)
    zip_buf.seek(0)

    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"Purchase Details - {date_range}.zip",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

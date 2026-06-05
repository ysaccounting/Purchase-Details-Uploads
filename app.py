import io
import os
import json
import zipfile
import threading
import uuid
import tempfile
import pickle
from flask import Flask, request, jsonify, send_file, render_template
from processor import process_files, build_filtered_outputs

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

JOBS_DIR = os.path.join(tempfile.gettempdir(), "ticketvault_jobs")
os.makedirs(JOBS_DIR, exist_ok=True)


def job_dir(job_id):
    return os.path.join(JOBS_DIR, job_id)


def write_job_status(job_id, status, message=None):
    d = job_dir(job_id)
    os.makedirs(d, exist_ok=True)
    payload = {"status": status}
    if message:
        payload["message"] = message
    with open(os.path.join(d, "status.json"), "w") as f:
        json.dump(payload, f)


def read_job_status(job_id):
    path = os.path.join(job_dir(job_id), "status.json")
    if not os.path.exists(path):
        return None
    with open(path) as f:
        return json.load(f)


def read_meta(job_id):
    path = os.path.join(job_dir(job_id), "meta.json")
    if not os.path.exists(path):
        return None
    with open(path) as f:
        return json.load(f)


def run_job(job_id, file_list):
    try:
        result = process_files(file_list)
        d = job_dir(job_id)

        # Pickle DataFrames for deferred filtered output building
        with open(os.path.join(d, "dataframes.pkl"), "wb") as f:
            pickle.dump({
                "df_raw":       result["_df_raw"],
                "df_cancelled": result["_df_cancelled"],
                "all_df":       result["_all_df"],
                "summary_df":   result["_summary_df"],
                "company_dfs":  result["_company_dfs"],
            }, f)

        meta = {
            "date_range":    result["date_range"],
            "all_companies": result["all_companies"],
            "stats":         result["stats"],
        }
        with open(os.path.join(d, "meta.json"), "w") as f:
            json.dump(meta, f)

        write_job_status(job_id, "done")

    except Exception as e:
        import traceback
        write_job_status(job_id, "error", traceback.format_exc())


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("file")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files provided"}), 400
    file_list = [(f.read(), f.filename) for f in files if f.filename]
    job_id = str(uuid.uuid4())
    write_job_status(job_id, "processing")
    threading.Thread(target=run_job, args=(job_id, file_list), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    job = read_job_status(job_id)
    if not job:
        return jsonify({"status": "not_found"}), 404
    if job["status"] in ("error", "processing"):
        return jsonify(job)
    meta = read_meta(job_id)
    if not meta:
        return jsonify({"status": "error", "message": "Result files missing"}), 500
    return jsonify({
        "status": "done",
        "date_range":    meta["date_range"],
        "all_companies": meta["all_companies"],
        "stats":         meta["stats"],
    })


@app.route("/configure/<job_id>", methods=["POST"])
def configure(job_id):
    """Accept selected companies, build filtered output files, save to disk."""
    meta = read_meta(job_id)
    if not meta:
        return jsonify({"error": "Job not found"}), 404

    data = request.get_json()
    selected_companies = data.get("selected_companies", meta["all_companies"])

    # Load pickled DataFrames
    pkl_path = os.path.join(job_dir(job_id), "dataframes.pkl")
    if not os.path.exists(pkl_path):
        return jsonify({"error": "Data not found"}), 404
    with open(pkl_path, "rb") as f:
        dfs = pickle.load(f)

    combined_bytes, company_files = build_filtered_outputs(
        dfs["df_raw"], dfs["df_cancelled"], dfs["all_df"],
        dfs["summary_df"], dfs["company_dfs"], selected_companies
    )

    d = job_dir(job_id)

    # Write combined
    with open(os.path.join(d, "combined.xlsx"), "wb") as f:
        f.write(combined_bytes)

    # Write company files
    companies_dir = os.path.join(d, "companies")
    os.makedirs(companies_dir, exist_ok=True)
    # Clear old company files
    for fn in os.listdir(companies_dir):
        os.remove(os.path.join(companies_dir, fn))
    for company, file_bytes in company_files.items():
        safe = company.replace("/", "_").replace("\\", "_")
        with open(os.path.join(companies_dir, f"{safe}.xlsx"), "wb") as f:
            f.write(file_bytes)

    # Update meta with selection
    meta["selected_companies"] = selected_companies
    meta["available_companies"] = list(company_files.keys())
    with open(os.path.join(d, "meta.json"), "w") as f:
        json.dump(meta, f)

    return jsonify({
        "status": "ready",
        "date_range":           meta["date_range"],
        "available_companies":  meta["available_companies"],
        "selected_companies":   selected_companies,
    })


@app.route("/download/<job_id>/combined")
def download_combined(job_id):
    meta = read_meta(job_id)
    if not meta:
        return jsonify({"error": "Job not found"}), 404
    path = os.path.join(job_dir(job_id), "combined.xlsx")
    if not os.path.exists(path):
        return jsonify({"error": "File not found — please generate first"}), 404
    filename = f"Purchase Details - Combined - {meta['date_range']}.xlsx"
    return send_file(path, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/download/<job_id>/company/<company>")
def download_company(job_id, company):
    meta = read_meta(job_id)
    if not meta:
        return jsonify({"error": "Job not found"}), 404
    safe = company.replace("/", "_").replace("\\", "_")
    path = os.path.join(job_dir(job_id), "companies", f"{safe}.xlsx")
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    filename = f"Purchase Details - {company} - {meta['date_range']}.xlsx"
    return send_file(path, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/download/<job_id>/all")
def download_all_zip(job_id):
    meta = read_meta(job_id)
    if not meta:
        return jsonify({"error": "Job not found"}), 404
    d = job_dir(job_id)
    dr = meta["date_range"]
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        combined = os.path.join(d, "combined.xlsx")
        if os.path.exists(combined):
            zf.write(combined, f"Purchase Details - Combined - {dr}.xlsx")
        companies_dir = os.path.join(d, "companies")
        if os.path.exists(companies_dir):
            for company in meta.get("available_companies", []):
                safe = company.replace("/", "_").replace("\\", "_")
                cp = os.path.join(companies_dir, f"{safe}.xlsx")
                if os.path.exists(cp):
                    zf.write(cp, f"Purchase Details - {company} - {dr}.xlsx")
    zip_buf.seek(0)
    return send_file(zip_buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"Purchase Details - {dr}.zip")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

"""
Compensation Grade EIB Builder - Web Application
Flask-based web interface for generating Workday EIB templates.
"""
import os
import io
import json
from datetime import date
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from eib_engine import (
    generate_workday_eib,
    parse_uploaded_excel,
    parse_job_code_table,
    FACTORS,
    CAREER_BANDS_EXECUTIVE,
    CAREER_BANDS_FULL,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB max upload
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

JOB_CODE_TABLE_LINK = (
    "https://tmobileusa.sharepoint.com/sites/Compensation2/Compensation%20Repository/"
    "Forms/AllItems.aspx?id=%2Fsites%2FCompensation2%2FCompensation%20Repository%2F"
    "Workday%20Reports%2FJob%20Code%20Table&viewid=76927db9%2Dcec5%2D42db%2D8a70%2D"
    "bc8d668a500c"
)


@app.route("/")
def index():
    return render_template(
        "index.html",
        career_bands_executive=CAREER_BANDS_EXECUTIVE,
        career_bands_full=CAREER_BANDS_FULL,
        default_date=date.today().replace(month=1, day=1).isoformat(),
        job_code_table_link=JOB_CODE_TABLE_LINK,
    )


@app.route("/generate", methods=["POST"])
def generate():
    """Generate EIB from manual form input."""
    data = request.get_json()
    template_type = data.get("template_type", "New")
    employment_type = data.get("employment_type", "Exempt or TTC")
    jobs = data.get("jobs", [])

    if not jobs:
        return jsonify({"error": "No jobs provided."}), 400

    # Validate
    for i, job in enumerate(jobs):
        if not job.get("job_code"):
            return jsonify({"error": f"Job #{i+1}: Job Code is required."}), 400
        if not job.get("job_title"):
            return jsonify({"error": f"Job #{i+1}: Job Title is required."}), 400
        if not job.get("career_band"):
            return jsonify({"error": f"Job #{i+1}: Career Band is required."}), 400
        try:
            market_50th = float(job["national_market_50th"])
            if market_50th <= 0:
                return jsonify({"error": f"Job #{i+1}: National Market 50th must be a positive number greater than 0."}), 400
        except (ValueError, TypeError):
            return jsonify({"error": f"Job #{i+1}: National Market 50th must be a valid number."}), 400

    df = generate_workday_eib(jobs, template_type, employment_type)

    # Convert to JSON for preview
    preview = df.head(100).to_dict(orient="records")
    total_rows = len(df)
    factors_count = len(FACTORS.get(employment_type, []))

    return jsonify({
        "preview": preview,
        "total_rows": total_rows,
        "jobs_count": len(jobs),
        "factors_count": factors_count,
        "columns": list(df.columns),
    })


@app.route("/download", methods=["POST"])
def download():
    """Download the generated EIB as Excel."""
    data = request.get_json()
    template_type = data.get("template_type", "New")
    employment_type = data.get("employment_type", "Exempt or TTC")
    jobs = data.get("jobs", [])

    if not jobs:
        return jsonify({"error": "No jobs provided."}), 400

    df = generate_workday_eib(jobs, template_type, employment_type)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Compensation Grade EIB")
    output.seek(0)

    filename = f"Compensation_Grade_EIB_{template_type}_{employment_type.replace(' ', '_')}_{date.today().isoformat()}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/upload", methods=["POST"])
def upload():
    """Upload an Excel file and parse jobs from it."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    try:
        jobs, template_type, employment_type = parse_uploaded_excel(filepath)
        return jsonify({
            "jobs": jobs,
            "template_type": template_type,
            "employment_type": employment_type,
        })
    except Exception as e:
        return jsonify({"error": f"Failed to parse file: {str(e)}"}), 400
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)


@app.route("/upload-job-code-table", methods=["POST"])
def upload_job_code_table():
    """Upload a Job Code Table file and return parsed job records for dropdown population."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    try:
        result = parse_job_code_table(filepath)
        job_records = result["jobs"]
        if not job_records:
            return jsonify({"error": f"No job records found. Columns detected: {', '.join(result['columns_found'])}"}), 400
        return jsonify({
            "jobs": job_records,
            "columns_found": result["columns_found"],
            "mapped": result["mapped"],
        })
    except Exception as e:
        return jsonify({"error": f"Failed to parse Job Code Table: {str(e)}"}), 400
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)


@app.route("/download-csv", methods=["POST"])
def download_csv():
    """Download the generated EIB as CSV."""
    data = request.get_json()
    template_type = data.get("template_type", "New")
    employment_type = data.get("employment_type", "Exempt or TTC")
    jobs = data.get("jobs", [])

    if not jobs:
        return jsonify({"error": "No jobs provided."}), 400

    df = generate_workday_eib(jobs, template_type, employment_type)

    output = io.BytesIO()
    df.to_csv(output, index=False)
    output.seek(0)

    filename = f"Compensation_Grade_EIB_{template_type}_{employment_type.replace(' ', '_')}_{date.today().isoformat()}.csv"
    return send_file(
        output,
        mimetype="text/csv",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)

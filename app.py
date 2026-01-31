from flask import Flask, render_template, request, send_file
import os
import zipfile
import subprocess
import shutil
from pathlib import Path
from services.summary import build_summary
import services.loaders as loaders

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# =========================
# Index (GET فقط)
# =========================
@app.route("/", methods=["GET"])
def index():
    # Refresh data core lists on every page load
    loaders.site_master_dict, loaders.valid_sites, loaders.oz_list, loaders.comments_list = loaders.get_live_data()
    return render_template(
        "Index.html",
        oz_list=loaders.oz_list
    )

# =========================
# Process (POST)
# =========================
@app.route("/process", methods=["POST"])
def process():
    file = request.files.get("NSN Update")
    selected_oz = request.form.get("oz")  # اختيار OZ من radio button

    if not file:
        return "Please upload a file named 'NSN Update'", 400

    if not selected_oz:
        return "Please select OZ", 400

    filename = file.filename
    path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(path)

    # Handle ZIP/RAR files
    if filename.lower().endswith((".zip", ".rar")):
        try:
            extract_dir = os.path.join(UPLOAD_FOLDER, "extracted_" + Path(filename).stem)
            if os.path.exists(extract_dir):
                shutil.rmtree(extract_dir)
            os.makedirs(extract_dir, exist_ok=True)
            
            if filename.lower().endswith(".zip"):
                # ZIP is handled by pure python zipfile (safe everywhere)
                with zipfile.ZipFile(path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
            else: 
                # RAR is handled by system 'tar' (works on Windows 10/11)
                result = subprocess.run(["tar", "-xf", path, "-C", extract_dir], capture_output=True, text=True)
                if result.returncode != 0:
                    # If system tar fails (common on Linux/Deployed), suggest ZIP
                    return "RAR extraction is not supported on this server/system. Please upload a .zip or .xlsx file instead.", 400
                
            # Find the first .xlsx file
            xlsx_files = list(Path(extract_dir).rglob("*.xlsx"))
            if not xlsx_files:
                return "No Excel file (.xlsx) found in the archive", 400
            
            path = str(xlsx_files[0]) # Use the first found Excel file
        except Exception as e:
            return f"Error extracting archive: {e}", 400

    # Build summary (بنفس المنطق)
    df, dashboard, dashboard_summary, tables_down_env, critical_env_table, tables_env_only, \
    tech_labels, tech_counts, down_type_counts, env_labels, env_values, excel_path = \
        build_summary(path, selected_oz)

    return render_template(
        "result.html",
        tables_down_env=tables_down_env,
        tables_env_only=tables_env_only,
        dashboard=dashboard,
        dashboard_summary=dashboard_summary,
        tech_labels=tech_labels,
        tech_counts=tech_counts,
        down_type_counts=down_type_counts,
        env_labels=env_labels,
        env_values=env_values,
        critical_env_table=critical_env_table,
        excel_path=excel_path
    )

# =========================
# Download
# =========================
@app.route("/download")
def download():
    file_path = request.args.get("file")
    if not file_path or not os.path.exists(file_path):
        return "File not found", 404
    return send_file(file_path, as_attachment=True)

# =========================
# Run
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

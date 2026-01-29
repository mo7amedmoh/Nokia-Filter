from flask import Flask, render_template, request, send_file
import os
from services.summary import build_summary
from services.loaders import oz_list

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# =========================
# Index (GET فقط)
# =========================
@app.route("/", methods=["GET"])
def index():
    return render_template(
        "Index.html",
        oz_list=oz_list
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

    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    # Build summary (بنفس المنطق)
    df, dashboard, dashboard_summary, tables_down_env, tables_env_only,critical_env_table, \
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

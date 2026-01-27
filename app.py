from flask import Flask, render_template, request, send_file
import os
from services.summary import build_summary

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files.get("NSN Update")
        if not file:
            return "Please upload a file named 'NSN Update'", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        # Build summary
        df, dashboard, dashboard_summary, tables_down_env, tables_env_only, tech_labels, tech_counts, down_type_counts, env_labels, env_values, excel_path = build_summary(path)

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
            excel_path=excel_path
        )

    return render_template("Index.html")

@app.route("/download")
def download():
    file_path = request.args.get("file")
    if not file_path or not os.path.exists(file_path):
        return "File not found", 404
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)


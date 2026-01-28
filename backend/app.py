import os
import uuid
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

# import your existing logic
from generate_testcases import (
    extract_srs_text,
    get_testcases_from_claude,
    extract_component,
    parse_markdown_table,
    fill_excel_template,
)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
CORS(app)  # allow GitHub Pages to call backend

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER


@app.route("/generate", methods=["POST"])
def generate():
    try:
        if "srs_file" not in request.files or "template_file" not in request.files:
            return jsonify({"error": "Missing required files"}), 400

        srs_file = request.files["srs_file"]
        template_file = request.files["template_file"]

        run_id = str(uuid.uuid4())

        srs_path = os.path.join(UPLOAD_FOLDER, f"{run_id}_SRS.docx")
        template_path = os.path.join(UPLOAD_FOLDER, f"{run_id}_template.xlsx")
        output_path = os.path.join(OUTPUT_FOLDER, f"{run_id}_Generated_TestCases.xlsx")

        srs_file.save(srs_path)
        template_file.save(template_path)

        # ---- Run your existing logic ----
        srs_text = extract_srs_text(srs_path)
        md_output = get_testcases_from_claude(srs_text)
        component = extract_component(md_output)
        test_cases = parse_markdown_table(md_output)

        fill_excel_template(
            test_cases,
            template_path,
            output_path,
            component
        )

        # simple short summary (first few lines)
        srs_summary = "\n".join(srs_text.splitlines()[:8])

        return jsonify({
            "component": component,
            "srs_summary": srs_summary,
            "download_url": f"/download/{os.path.basename(output_path)}"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download/<filename>", methods=["GET"])
def download(filename):
    return send_from_directory(
        app.config["OUTPUT_FOLDER"],
        filename,
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)


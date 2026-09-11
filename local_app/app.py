"""
Local WebPT Scheduler History tool. Run with:
    python app.py
Then open http://localhost:5000 in your browser.

Needs the same env vars as scripts/snowflake_writer.py (SNOWFLAKE_*),
plus GITHUB_TOKEN (for the scrape trigger), SHEET_ID and
GOOGLE_SERVICE_ACCOUNT_FILE (to read the Emails tab), and SMTP_USER /
SMTP_PASSWORD (to actually send email). See README in this folder.
"""

import os
import sys

from flask import Flask, request, jsonify, render_template
import requests

sys.path.insert(0, os.path.dirname(__file__))
import report_engine  # noqa: E402

# Load a local .env file if python-dotenv is installed and one exists —
# convenience only, never required (env vars set in the terminal work fine
# without this).
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

app = Flask(__name__)

GITHUB_USERNAME = "ibrahimramadan-COB"
GITHUB_REPO = "Scheduler_History"
GITHUB_WORKFLOW_FILE = "scrape.yml"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/filter-options")
def filter_options():
    try:
        return jsonify(report_engine.get_filter_options())
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/trigger-scrape", methods=["POST"])
def trigger_scrape():
    data = request.get_json()
    token = os.environ.get("GITHUB_TOKEN")
    if not token:
        return jsonify({"message": "❌ GITHUB_TOKEN env var is not set."})

    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{GITHUB_REPO}/actions/workflows/{GITHUB_WORKFLOW_FILE}/dispatches"
    inputs = {"start_date": data["startDate"], "end_date": data["endDate"]}
    if data.get("clinics"):
        inputs["clinics"] = "|".join(data["clinics"])

    resp = requests.post(url, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }, json={"ref": "main", "inputs": inputs})

    if 200 <= resp.status_code < 300:
        return jsonify({"message": f"✅ GitHub scrape started for {data['startDate']} → {data['endDate']}. "
                                    f"Check the Actions tab, or come back later and run a report once it's done."})
    return jsonify({"message": f"❌ GitHub Error: {resp.text}"})


@app.route("/api/run-report", methods=["POST"])
def run_report():
    config = request.get_json()
    try:
        message = report_engine.run_scheduler_report(config)
        return jsonify({"message": message})
    except Exception as e:
        return jsonify({"message": f"❌ Error: {e}"})


if __name__ == "__main__":
    app.run(debug=True, port=5000)

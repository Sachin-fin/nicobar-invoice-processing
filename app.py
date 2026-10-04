import os
import sys
import subprocess
import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS

from config import (
    EXPENSES_DIR,
    REGISTER_PATH,
    VENDOR_CODES_PATH,
    DASHBOARD_HTML_PATH,
)
from scanner import InvoiceScanner
from excel_manager import ExcelManager
from gmail_intake import GmailIntake

app = Flask(__name__, static_folder="static", template_folder="templates")
CORS(app)

scanner = InvoiceScanner()
excel_mgr = ExcelManager()
gmail_intake = GmailIntake()

# Uploads directory
UPLOAD_DIR = EXPENSES_DIR / "_web_uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/gmail/status")
def gmail_status():
    """Return current Gmail sync state and configuration status."""
    try:
        state = gmail_intake.get_state()
        has_pwd = bool(os.getenv("GMAIL_APP_PASSWORD"))
        return jsonify({
            "status": "success",
            "mailbox": state.get("mailbox", "invoice@nicobar.com"),
            "has_password": has_pwd,
            "last_scan_ist": state.get("last_scan_ist", "Never"),
            "pending_count": len(state.get("pending_messages", [])),
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/gmail/config", methods=["POST"])
def gmail_config():
    """Save Gmail credentials to .env file."""
    try:
        data = request.json or {}
        mailbox = data.get("mailbox", "invoice@nicobar.com").strip()
        app_password = data.get("app_password", "").strip().replace(" ", "")

        env_file = Path(__file__).parent / ".env"
        lines = []
        if env_file.exists():
            with open(env_file, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if not line.startswith(("GMAIL_USER=", "GMAIL_APP_PASSWORD="))]

        lines.append(f"GMAIL_USER={mailbox}")
        if app_password:
            lines.append(f"GMAIL_APP_PASSWORD={app_password}")

        with open(env_file, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines) + "\n")

        os.environ["GMAIL_USER"] = mailbox
        if app_password:
            os.environ["GMAIL_APP_PASSWORD"] = app_password

        # Refresh intake
        gmail_intake.mailbox = mailbox
        if app_password:
            gmail_intake.app_password = app_password

        return jsonify({"status": "success", "message": "Gmail settings saved successfully!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/gmail/sync", methods=["POST"])
def gmail_sync():
    """Trigger Gmail intake, download invoices, and auto-process to Excel register."""
    try:
        dry_run = request.json.get("dry_run", False) if request.is_json else False
        res = gmail_intake.scan_and_auto_process(dry_run=dry_run)
        return jsonify({"status": "success", "result": res})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/status")
def get_status():
    """Return register summary, KPIs, and recent rows."""
    try:
        state = excel_mgr.get_register_state()
        unprocessed, _ = scanner.find_unprocessed_pdfs()
        return jsonify({
            "status": "success",
            "kpis": state.get("kpis", {}),
            "unprocessed_count": len(unprocessed),
            "unprocessed_files": [p.name for p in unprocessed[:10]],
            "recent_rows": state.get("rows", [])[-15:],
            "total_rows": state.get("total_rows", 0),
            "max_sno": state.get("max_sno", 0),
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/invoices")
def get_invoices():
    """Return all rows in the register with optional search filter."""
    try:
        q = request.args.get("q", "").strip().lower()
        state = excel_mgr.get_register_state()
        rows = state.get("rows", [])
        if q:
            rows = [
                r for r in rows
                if q in r.get("invoice_no", "").lower()
                or q in r.get("supplier", "").lower()
                or q in r.get("description", "").lower()
                or q in r.get("invoice_date", "").lower()
                or q in str(r.get("sno", ""))
            ]
        return jsonify({"status": "success", "count": len(rows), "rows": rows})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/scan", methods=["POST"])
def scan_expenses():
    """Trigger folder scan and process new invoices."""
    try:
        dry_run = request.json.get("dry_run", False) if request.is_json else False
        res = scanner.scan_and_process(dry_run=dry_run)
        return jsonify({"status": "success", "result": res})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/upload", methods=["POST"])
def upload_pdf():
    """Extract and optionally add an uploaded invoice PDF."""
    try:
        if "file" not in request.files:
            return jsonify({"status": "error", "message": "No file provided"}), 400
        
        file = request.files["file"]
        if file.filename == "":
            return jsonify({"status": "error", "message": "Empty filename"}), 400

        target_path = UPLOAD_DIR / file.filename
        file.save(target_path)

        auto_commit = request.form.get("auto_commit", "false").lower() == "true"
        extract_res = scanner.process_single_pdf(target_path, dry_run=not auto_commit)

        return jsonify({
            "status": "success",
            "extracted": extract_res,
            "filename": file.filename,
            "saved_path": str(target_path),
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/commit_extracted", methods=["POST"])
def commit_extracted():
    """Commit pre-extracted/reviewed invoice data to Excel register."""
    try:
        data = request.json
        inv_data = data.get("invoice")
        if not inv_data:
            return jsonify({"status": "error", "message": "No invoice data"}), 400

        commit_res = excel_mgr.add_invoices([inv_data])
        return jsonify({"status": "success", "commit": commit_res})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/open_file", methods=["POST"])
def open_file():
    """Open Excel register, Vendor codes, or Dashboard in default Windows apps."""
    try:
        target = request.json.get("target")
        path_map = {
            "register": REGISTER_PATH,
            "vendor_codes": VENDOR_CODES_PATH,
            "dashboard": DASHBOARD_HTML_PATH,
            "expenses": EXPENSES_DIR,
        }
        target_path = path_map.get(target)
        if target_path and target_path.exists():
            if hasattr(os, "startfile"):
                os.startfile(str(target_path))
                return jsonify({"status": "success", "message": f"Opened {target_path.name}"})
            return jsonify({"status": "error", "message": "Desktop file opening only supported on host machine"}), 400
        return jsonify({"status": "error", "message": "Target file not found"}), 404
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    print(f"Starting Nicobar Invoice Manager on http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)

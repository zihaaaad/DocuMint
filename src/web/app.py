import os
import sys
import threading
import json
from flask import Flask, render_template, request, jsonify, send_from_directory

# Add src and root to path
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.dirname(current_dir)
root_dir = os.path.dirname(src_dir)
sys.path.insert(0, src_dir)
sys.path.insert(0, root_dir)

from documint.core import process_emails, validate_placeholders
from documint.database import DatabaseManager

app = Flask(__name__)
app.secret_key = "documint_secure_key"

db_path = os.path.join(root_dir, 'history.db')
db = DatabaseManager(db_path)

# Global Job State
job_status = {
    "is_running": False,
    "logs": [],
    "total": 0,
    "completed": 0,
    "success_count": 0,
    "fail_count": 0
}
current_cancel_event = None
lock = threading.Lock()

def log_callback(message: str):
    """Callback to capture live execution logs."""
    with lock:
        job_status["logs"].append(message)
        if "Compiled PDF" in message or "Sent email" in message or "Delivery Failed" in message or "Skipped" in message:
            job_status["completed"] += 1
        if "✅" in message or "Success" in message:
            job_status["success_count"] += 1
        elif "❌" in message or "⚠️" in message:
            job_status["fail_count"] += 1

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/docs')
def docs():
    return render_template('docs.html')

@app.route('/download/<filename>')
def download_file(filename):
    examples_dir = os.path.abspath(os.path.join(root_dir, 'examples'))
    return send_from_directory(examples_dir, filename, as_attachment=True)

# --- Profile Management ---
PROFILES_DIR = os.path.join(root_dir, 'profiles')
if not os.path.exists(PROFILES_DIR):
    os.makedirs(PROFILES_DIR, exist_ok=True)

@app.route('/api/profiles', methods=['GET'])
def list_profiles():
    files = [f for f in os.listdir(PROFILES_DIR) if f.endswith('.json')]
    return jsonify({"profiles": files})

@app.route('/api/profiles', methods=['POST'])
def save_profile():
    data = request.json or {}
    name = data.get('name')
    config = data.get('config')
    if not name or not config:
        return jsonify({"error": "Missing profile name or config"}), 400
    
    safe_name = "".join([c for c in name if c.isalnum() or c in (' ', '_', '-')]).strip()
    file_path = os.path.join(PROFILES_DIR, f"{safe_name}.json")
    
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)
        
    return jsonify({"status": "success", "message": f"Profile '{safe_name}' saved."})

@app.route('/api/profiles/<filename>', methods=['GET'])
def load_profile(filename):
    safe_filename = os.path.basename(filename)
    file_path = os.path.join(PROFILES_DIR, safe_filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "Profile not found"}), 404
        
    with open(file_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
        
    return jsonify(config)

@app.route('/api/profiles/<filename>', methods=['DELETE'])
def delete_profile(filename):
    safe_filename = os.path.basename(filename)
    file_path = os.path.join(PROFILES_DIR, safe_filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "Profile not found"}), 404
    
    try:
        os.remove(file_path)
        return jsonify({"status": "success", "message": f"Profile '{safe_filename}' deleted."})
    except Exception as e:
        return jsonify({"error": f"Failed to delete profile: {e}"}), 500
# --------------------------

@app.route('/api/validate', methods=['POST'])
def validate():
    data = request.json or {}
    data_path = data.get('data_path', '').strip()
    template_path = data.get('template_path', '').strip()
    
    if not os.path.exists(data_path) or not os.path.exists(template_path):
        return jsonify({"valid": False, "error": "Selected data or template file was not found on system."})

    is_valid, missing = validate_placeholders(data_path, template_path)
    return jsonify({"valid": is_valid, "missing": missing})

@app.route('/api/run', methods=['POST'])
def run_job():
    global current_cancel_event
    with lock:
        if job_status["is_running"]:
            return jsonify({"status": "error", "message": "A batch job is already actively executing."}), 409

        data = request.json or {}
        job_status["is_running"] = True
        job_status["logs"] = ["🚀 Pipeline Initializing..."]
        job_status["total"] = 0
        job_status["completed"] = 0
        job_status["success_count"] = 0
        job_status["fail_count"] = 0
        current_cancel_event = threading.Event()

    cancel_event = current_cancel_event

    def background_task():
        try:
            # COM initialization for Windows threads if needed
            if os.name == 'nt':
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                except Exception:
                    pass

            dry_run = bool(data.get('dry_run', False))
            test_email = data.get('test_email')
            if test_email:
                test_email = str(test_email).strip() or None

            email_config = data.get('email_config', {})
            
            result = process_emails(
                data_file=data.get('data_path', ''),
                template_file=data.get('template_path', ''),
                pdf_folder=data.get('pdf_folder', './output/pdf'),
                logs_folder=data.get('logs_folder', './output/logs'),
                log_callback=log_callback,
                email_subject=data.get('subject', 'Official Document'),
                email_body=data.get('body', ''),
                pdf_filename_format=data.get('filename_format', '{Name}'),
                retries=int(data.get('retries', 1)),
                delay=int(data.get('delay', 0)),
                email_config=email_config,
                dry_run=dry_run,
                test_email=test_email,
                cancel_event=cancel_event,
                db_manager=db
            )

            # Record execution in SQLite
            mode_label = "Dry Run" if dry_run else (
                "SMTP (Parallel)" if email_config.get("provider") == "smtp" else "Outlook (Serial)"
            )
            if test_email:
                mode_label += " [Test Mode]"

            db.log_job(
                total=result["total"],
                success=result["success"],
                fail=result["failed"],
                mode=mode_label,
                records=result.get("records", [])
            )
            log_callback("🏁 Execution Finished. Audit records committed to history.db.")

        except Exception as e:
            log_callback(f"❌ Critical Pipeline Failure: {str(e)}")
        finally:
            with lock:
                job_status["is_running"] = False
            if os.name == 'nt':
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    thread = threading.Thread(target=background_task, daemon=True)
    thread.start()
    
    return jsonify({"status": "success", "message": "Batch pipeline dispatched in background."})

@app.route('/api/stop', methods=['POST'])
def stop_job():
    global current_cancel_event
    with lock:
        if not job_status["is_running"] or not current_cancel_event:
            return jsonify({"status": "info", "message": "No job is currently running."})
        current_cancel_event.set()
        job_status["logs"].append("🛑 Cancellation requested by user...")
    return jsonify({"status": "success", "message": "Cancellation signal sent."})

@app.route('/api/status')
def status():
    with lock:
        return jsonify(job_status)

@app.route('/api/stats')
def get_stats():
    return jsonify(db.get_stats())

@app.route('/api/jobs/<int:job_id>')
def get_job_details(job_id):
    details = db.get_job_details(job_id)
    if not details:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(details)

if __name__ == '__main__':
    print("🌍 DocuMint Web Server Running at http://127.0.0.1:5000")
    app.run(debug=True, port=5000)

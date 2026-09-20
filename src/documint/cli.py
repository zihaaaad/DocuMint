import os
import sys
import argparse
from typing import Optional

from documint.core import process_emails, validate_placeholders, UniversalDocumentConverter
from documint.database import DatabaseManager

# Ensure Windows console uses UTF-8 without crashing on emojis
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

def main(args: Optional[list] = None) -> None:
    parser = argparse.ArgumentParser(
        prog="documint",
        description="Universal Document Automation, PDF Compiler & Multi-Threaded Dispatch Platform in Python."
    )
    
    subparsers = parser.add_subparsers(dest="command", help="DocuMint commands")

    # 1. 'run' command
    run_parser = subparsers.add_parser("run", help="Execute batch document compilation and dispatch")
    run_parser.add_argument("--data", "-d", required=True, help="Path to input dataset (.xlsx, .xls, .csv)")
    run_parser.add_argument("--template", "-t", required=True, help="Path to Word template file (.docx)")
    run_parser.add_argument("--pdf-folder", "-o", default="./output/pdf", help="Output directory for generated PDFs")
    run_parser.add_argument("--logs-folder", "-l", default="./output/logs", help="Directory for execution logs")
    run_parser.add_argument("--subject", "-s", default="Official Document Delivery", help="Email subject template")
    run_parser.add_argument("--body", "-b", default="<p>Dear &lt;Name&gt;,<br>Please find your official document attached.</p>", help="Email body template (HTML)")
    run_parser.add_argument("--filename-format", "-f", default="{Name}", help="Filename pattern for compiled PDFs (e.g. '{Name}_{ID}')")
    run_parser.add_argument("--provider", choices=["smtp", "outlook"], default="smtp", help="Email dispatch provider")
    run_parser.add_argument("--smtp-host", default="smtp.gmail.com", help="SMTP server host")
    run_parser.add_argument("--smtp-port", type=int, default=587, help="SMTP server port (465 for SSL, 587 for STARTTLS)")
    run_parser.add_argument("--smtp-user", default="", help="SMTP username / email address")
    run_parser.add_argument("--smtp-password", default="", help="SMTP password or app password")
    run_parser.add_argument("--sender-name", default=None, help="Sender display name")
    run_parser.add_argument("--use-tls", action="store_true", default=None, help="Force TLS encryption")
    run_parser.add_argument("--dry-run", action="store_true", help="Compile documents only without dispatching emails")
    run_parser.add_argument("--test-email", default=None, help="Send a single test record to this recipient email")
    run_parser.add_argument("--retries", type=int, default=1, help="Retry attempts for failed emails")
    run_parser.add_argument("--delay", type=int, default=0, help="Delay between email dispatches in seconds")
    run_parser.add_argument("--db", default="./history.db", help="Path to SQLite history database")

    # 2. 'validate' command
    val_parser = subparsers.add_parser("validate", help="Validate template placeholders against spreadsheet columns")
    val_parser.add_argument("--data", "-d", required=True, help="Path to dataset file (.xlsx, .xls, .csv)")
    val_parser.add_argument("--template", "-t", required=True, help="Path to Word template file (.docx)")

    # 3. 'studio' command
    studio_parser = subparsers.add_parser("studio", help="Launch the local Flask Web Studio dashboard")
    studio_parser.add_argument("--host", default="127.0.0.1", help="Host address to bind the web studio to")
    studio_parser.add_argument("--port", type=int, default=5000, help="Port to run the web studio server on")

    # 4. 'stats' command
    stats_parser = subparsers.add_parser("stats", help="View aggregate execution statistics from history database")
    stats_parser.add_argument("--db", default="./history.db", help="Path to SQLite history database")

    # 5. 'gui' command
    subparsers.add_parser("gui", help="Launch the legacy Tkinter desktop interface")

    parsed = parser.parse_args(args)

    if not parsed.command:
        parser.print_help()
        sys.exit(0)

    if parsed.command == "validate":
        print(f"🔍 Validating '{parsed.template}' against '{parsed.data}'...")
        is_valid, missing = validate_placeholders(parsed.data, parsed.template)
        if is_valid:
            print("✅ Pre-Flight Validation PASSED! All placeholders match data columns.")
            sys.exit(0)
        else:
            print(f"❌ Pre-Flight Validation FAILED! Missing columns: {missing}")
            sys.exit(1)

    elif parsed.command == "run":
        print("🚀 DocuMint Batch Pipeline Initializing...")
        
        # Pre-flight check
        is_valid, missing = validate_placeholders(parsed.data, parsed.template)
        if not is_valid:
            print(f"❌ Cannot proceed. Schema mismatch! Missing columns: {missing}")
            sys.exit(1)

        email_config = {
            "provider": parsed.provider,
            "smtp_host": parsed.smtp_host,
            "smtp_port": parsed.smtp_port,
            "smtp_user": parsed.smtp_user,
            "smtp_password": parsed.smtp_password,
            "sender_name": parsed.sender_name,
            "use_tls": parsed.use_tls
        }

        db_manager = DatabaseManager(parsed.db)

        def console_log(msg: str):
            print(msg)

        result = process_emails(
            data_file=parsed.data,
            template_file=parsed.template,
            pdf_folder=parsed.pdf_folder,
            logs_folder=parsed.logs_folder,
            log_callback=console_log,
            email_subject=parsed.subject,
            email_body=parsed.body,
            pdf_filename_format=parsed.filename_format,
            retries=parsed.retries,
            delay=parsed.delay,
            dry_run=parsed.dry_run,
            test_email=parsed.test_email,
            email_config=email_config,
            db_manager=db_manager
        )

        mode_str = "Dry Run" if parsed.dry_run else f"{parsed.provider.upper()}"
        db_manager.log_job(
            total=result["total"],
            success=result["success"],
            fail=result["failed"],
            mode=mode_str,
            records=result["records"]
        )

        print("\n" + "="*50)
        print(f"🏁 DocuMint Run Completed: Total={result['total']}, Success={result['success']}, Failed={result['failed']}")
        print("="*50)

    elif parsed.command == "studio":
        # Launch Flask Web Studio
        from web.app import app
        print(f"🌍 Launching DocuMint Automation Studio at http://{parsed.host}:{parsed.port}")
        app.run(host=parsed.host, port=parsed.port, debug=False)

    elif parsed.command == "stats":
        db_manager = DatabaseManager(parsed.db)
        stats = db_manager.get_stats()
        summary = stats["summary"]
        print("\n📊 --- DocuMint Execution Analytics ---")
        print(f"Total Batches Run : {summary.get('total_jobs', 0)}")
        print(f"Total Documents   : {summary.get('total_emails', 0)}")
        print(f"Successful Sends  : {summary.get('total_success', 0)}")
        print(f"Failed Sends      : {summary.get('total_failed', 0)}")
        print(f"Success Rate      : {summary.get('success_rate', 100.0)}%")
        print("----------------------------------------\n")

    elif parsed.command == "gui":
        from documint.gui import main as gui_main
        gui_main()

if __name__ == "__main__":
    main()

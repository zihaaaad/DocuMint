import sys
import os

# Add src and parent directory to sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, current_dir)
sys.path.insert(0, parent_dir)

from documint.cli import main as cli_main

if __name__ == "__main__":
    if len(sys.argv) > 1:
        cli_main(sys.argv[1:])
    else:
        # Default launcher: starts Web Studio
        from web.app import app
        print("🌍 DocuMint Automation Studio starting at http://127.0.0.1:5000")
        app.run(host="127.0.0.1", port=5000, debug=True)

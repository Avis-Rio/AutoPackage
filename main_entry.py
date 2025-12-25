import os
import sys
import webbrowser
import uvicorn
import socket
from pathlib import Path

# Add project root to sys.path so we can import modules
# In PyInstaller, we are running from a temp dir, but structure should be preserved if we bundle correctly.
# However, if we bundle as onefile, resources are in sys._MEIPASS

if getattr(sys, 'frozen', False):
    # Running as compiled exe
    base_path = sys._MEIPASS
    # Adjust sys.path to include the location of our modules
    # Assuming we bundle everything flat or in folders
    sys.path.insert(0, os.path.join(base_path, 'AutoPackage'))
    sys.path.insert(0, os.path.join(base_path, 'web_server'))
else:
    # Running from source
    base_path = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.join(base_path, 'AutoPackage'))
    sys.path.insert(0, os.path.join(base_path, 'web_server'))

# Import the FastAPI app
# Note: we import 'app' from 'web_server.web_app'
# Make sure web_server package is importable
try:
    from web_server.web_app import app
except ImportError:
    # Fallback: try importing directly if we are inside web_server or flat structure
    try:
        from web_app import app
    except ImportError as e:
        print(f"Error importing app: {e}")
        print(f"sys.path: {sys.path}")
        input("Press Enter to exit...")
        sys.exit(1)

def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

def main():
    port = 8000
    # Try to find a free port if 8000 is taken, or just stick to 8000
    # Let's try 8000, if fail, try others?
    # For simplicity, let's stick to 8000 or allow env var
    
    host = "127.0.0.1"
    
    print(f"Starting AutoPackage Server at http://{host}:{port}")
    
    # Open browser automatically
    webbrowser.open(f"http://{host}:{port}")
    
    # Run server
    # workers=1 is standard for local app
    uvicorn.run(app, host=host, port=port, log_level="info")

if __name__ == "__main__":
    main()

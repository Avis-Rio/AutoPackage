import PyInstaller.__main__
import os
import shutil
from pathlib import Path

# Project paths
BASE_DIR = Path(__file__).parent.absolute()
WEB_SERVER_DIR = BASE_DIR / "web_server"
STATIC_DIR = WEB_SERVER_DIR / "static"
AUTOPACKAGE_DIR = BASE_DIR / "AutoPackage"
TEMPLATES_DIR = BASE_DIR / "templates"

# Output name
APP_NAME = "AutoPackage_V2"

def build():
    # Ensure static dir exists
    if not STATIC_DIR.exists():
        print(f"Error: Static directory not found at {STATIC_DIR}")
        return

    # Prepare datas (source, destination_in_exe)
    datas = [
        # Include web_server/static -> web_server/static
        (str(STATIC_DIR), "web_server/static"),
        
        # Include AutoPackage source files?
        # Actually, PyInstaller analyzes imports. But data files inside AutoPackage (fonts, template.xlsx) need to be added.
        # Let's add the whole AutoPackage folder as data just in case for dynamic loading, 
        # OR better: add specific data files and rely on PyInstaller for python code.
        # Given we mess with sys.path in web_app.py, adding it as source root is important.
        
        # Add fonts
        (str(AUTOPACKAGE_DIR / "fonts"), "AutoPackage/fonts"),
        # Add template.xlsx in AutoPackage
        (str(AUTOPACKAGE_DIR / "template.xlsx"), "AutoPackage"),
    ]

    # Check for root templates dir
    if TEMPLATES_DIR.exists():
        datas.append((str(TEMPLATES_DIR), "templates"))

    # Define hidden imports
    hidden_imports = [
        "uvicorn.logging",
        "uvicorn.loops",
        "uvicorn.loops.auto",
        "uvicorn.protocols",
        "uvicorn.protocols.http",
        "uvicorn.protocols.http.auto",
        "uvicorn.lifespan",
        "uvicorn.lifespan.on",
        "multipart",
        "python-multipart",
        "openpyxl",
        "pandas",
        "win32com",
        "win32com.client",
        # AutoPackage modules might be missed if imported dynamically
        "excel_reader",
        "data_transformer",
        "template_writer",
        "delivery_note_generator",
        "assortment_generator",
        "store_detail_writer",
        "box_label_generator",
        "config"
    ]

    # PyInstaller arguments
    args = [
        "main_entry.py",  # Script to bundle
        f"--name={APP_NAME}",
        "--noconfirm",
        "--clean",
        "--onefile",      # Single exe
        # "--windowed",   # Hide console? No, keep console for server logs initially for debugging. 
                          # User can close it to stop server. Or we can make it windowed later.
                          # For a web server app, console is useful to see it's running.
        
        # Add paths to search for imports
        f"--paths={str(BASE_DIR)}",
        f"--paths={str(WEB_SERVER_DIR)}",
        f"--paths={str(AUTOPACKAGE_DIR)}",
        
        # Exclude modules to save size (optional)
        "--exclude-module=tkinter",
        "--exclude-module=matplotlib",
        "--exclude-module=scipy",
        # "--exclude-module=numpy", # pandas needs numpy
    ]

    # Add datas
    for src, dst in datas:
        # On Windows, separator is ;
        args.append(f"--add-data={src}{os.pathsep}{dst}")

    # Add hidden imports
    for hidden in hidden_imports:
        args.append(f"--hidden-import={hidden}")

    print("Starting build...")
    PyInstaller.__main__.run(args)
    print("Build finished.")

if __name__ == "__main__":
    build()

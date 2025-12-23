@echo off
echo ===================================================
echo AutoPackage V2 Web Service - Initializing...
echo ===================================================

cd /d "%~dp0"

echo [1/3] Checking dependencies...
pip install -r requirements.txt

echo [2/3] Starting Backend Server...
echo Please wait, browser will open automatically...

python web_server/web_app.py

pause

@echo off
echo Starting CSV to XLSX Converter...

:: Check if virtual environment exists
if exist "venv\Scripts\python.exe" (
    venv\Scripts\python.exe main.py
) else (
    :: Try with system Python
    python main.py
)

if errorlevel 1 (
    echo.
    echo Error running application. Make sure dependencies are installed:
    echo   pip install -r requirements.txt
    pause
)

@echo off
echo ========================================
echo   CSV to XLSX Converter - Build Script
echo ========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

:: Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

:: Activate virtual environment
call venv\Scripts\activate.bat

:: Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

:: Build executable
echo.
echo Building executable...
pyinstaller --onefile --windowed --name="CSV-to-XLSX-Converter" --add-data="README.md;." main.py

:: Check if build was successful
if exist "dist\CSV-to-XLSX-Converter.exe" (
    echo.
    echo ========================================
    echo   BUILD SUCCESSFUL!
    echo   Executable: dist\CSV-to-XLSX-Converter.exe
    echo ========================================
) else (
    echo.
    echo ERROR: Build failed!
)

:: Deactivate virtual environment
deactivate

pause

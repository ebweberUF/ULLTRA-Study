@echo off
REM ULLTRA Dashboard Launcher
REM This script runs the ULLTRA Study Dashboard

echo.
echo ================================================
echo  ULLTRA Study Dashboard
echo  Photobiomodulation for TMD Pain Management
echo ================================================
echo.

REM Change to the script directory
cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python 3.8 or later from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo Starting ULLTRA Dashboard...
echo.
echo The dashboard will:
echo  - Start a local web server
echo  - Open your browser automatically  
echo  - Show a control window
echo  - Connect to REDCap for real data
echo.
echo Press Ctrl+C to stop the server
echo.

REM Run the dashboard application
python app.py

REM If we get here, the application has closed
echo.
echo Dashboard stopped.
pause
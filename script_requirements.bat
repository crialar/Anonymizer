@echo off
echo Installing dependencies...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Download from: https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt

echo.
echo Installation complete.
echo.
pause

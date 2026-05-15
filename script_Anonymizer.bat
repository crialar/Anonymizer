@echo off
setlocal enabledelayedexpansion
color 00
cls

for /f %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"

set "CYAN=%ESC%[96m"
set "WHITE=%ESC%[97m"
set "GRAY=%ESC%[90m"
set "YELLOW=%ESC%[93m"
set "GREEN=%ESC%[92m"
set "RED=%ESC%[91m"
set "RESET=%ESC%[0m"

echo.
echo  %CYAN%+------------------------------------------+%RESET%
echo  %CYAN%^|                                          ^|%RESET%
echo  %CYAN%^|%RESET%           %WHITE%A N O N Y M I Z E R%RESET%            %CYAN%^|%RESET%
echo  %CYAN%^|%RESET%                   %GRAY%v6.2%RESET%                   %CYAN%^|%RESET%
echo  %CYAN%^|                                          ^|%RESET%
echo  %CYAN%+------------------------------------------+%RESET%
echo.
echo  %GRAY%  Clinical Document Anonymization Tool%RESET%
echo.
echo  %CYAN%-------------------------------------------%RESET%
echo.
echo  %GRAY%  Checking requirements...%RESET%
echo.

set MISSING=0

python --version >nul 2>&1
if errorlevel 1 (
    echo  %RED%  [x] Python        - NOT FOUND%RESET%
    echo.
    echo  %RED%  ERROR: Python is not installed.%RESET%
    echo  %GRAY%  Download from: https://www.python.org/downloads/%RESET%
    echo.
    pause
    exit /b 1
)
echo  %GREEN%  [+] Python        - OK%RESET%

python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo  %RED%  [x] Streamlit     - NOT FOUND%RESET%
    set MISSING=1
) else (
    echo  %GREEN%  [+] Streamlit     - OK%RESET%
)

python -c "import spacy" >nul 2>&1
if errorlevel 1 (
    echo  %RED%  [x] spaCy         - NOT FOUND%RESET%
    set MISSING=1
) else (
    echo  %GREEN%  [+] spaCy         - OK%RESET%
)

python -c "import presidio_analyzer" >nul 2>&1
if errorlevel 1 (
    echo  %RED%  [x] Presidio      - NOT FOUND%RESET%
    set MISSING=1
) else (
    echo  %GREEN%  [+] Presidio      - OK%RESET%
)

python -c "import lxml" >nul 2>&1
if errorlevel 1 (
    echo  %RED%  [x] lxml          - NOT FOUND%RESET%
    set MISSING=1
) else (
    echo  %GREEN%  [+] lxml          - OK%RESET%
)

if !MISSING!==1 (
    echo.
    echo  %RED%  Some dependencies are missing.%RESET%
    echo  %GRAY%  Run Dependencies.bat to install them.%RESET%
    echo.
    pause
    exit /b 1
)

echo.
echo  %CYAN%-------------------------------------------%RESET%
echo.
echo  %YELLOW%  Starting application...%RESET%
echo.
echo  %GRAY%  Browser  : %WHITE%Google Chrome%RESET%
echo  %GRAY%  Port     : %WHITE%5000%RESET%
echo.
echo  %GRAY%  To close, press %WHITE%Ctrl+C%GRAY% in this window.%RESET%
echo.
echo  %CYAN%-------------------------------------------%RESET%
echo.

start "" http://localhost:5000
python -m streamlit run app.py --server.port 5000

pause

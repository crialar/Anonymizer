@echo off
echo ============================================
echo              Anonymizer v6.1
echo ============================================
echo.
echo Starting program...
echo.
echo It is recommended to use an updated Chrome browser and Python 3.12.10.
echo To close, press Ctrl+C in this window.
echo.

start "" http://localhost:5000
streamlit run app.py --server.port 5000

pause

bat
@echo off
cd /d "%~dp0"

set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

if not exist ".venv\Scripts\python.exe" (
    echo Python virtual environment not found.
    echo Run setup first or check your Python installation.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

echo Starting the local time manager...
python -m streamlit run app.py

pause

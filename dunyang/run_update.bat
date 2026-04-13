@echo off
cd /d "%~dp0"

echo [1/3] Checking or creating Python virtual environment (.venv)...
if not exist ".venv\Scripts\activate.bat" (
    echo [1/3] Initializing virtual environment. This may take a few seconds...
    python -m venv .venv
)

echo [2/3] Activating virtual environment and checking requirements...
call .venv\Scripts\activate.bat

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo ----------------------------------------
echo [3/3] Running table conversion script (update_tables.py)...
echo ----------------------------------------
python update_tables.py

echo ----------------------------------------
echo Done! Press any key to exit...
pause >nul

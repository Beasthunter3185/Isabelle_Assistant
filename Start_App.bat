@echo off
title Isabelle Assistant Launcher
cd /d "%~dp0"

echo ===================================================
echo           Isabelle Assistant Launcher
echo ===================================================

:: 1. Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not added to PATH.
    echo Please install Python 3.10+ and make sure to check "Add Python to PATH".
    pause
    exit /b
)

:: 2. Check if Virtual Environment exists; create it if missing
if not exist "venv" (
    echo [SETUP] First-run detected! Creating virtual environment in 'venv'...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b
    )
)

:: 3. Activate Virtual Environment and check dependencies
call venv\Scripts\activate.bat

if not exist "venv\.setup_complete" (
    echo [SETUP] Installing required Python libraries from ChatServer\requirements.txt...
    pip install -r ChatServer\requirements.txt
    if %errorlevel% equ 0 (
        echo ready > venv\.setup_complete
        echo [SETUP] Dependencies installed successfully!
    ) else (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b
    )
)

:: 4. Start the ChatServer using setup_and_run.py
echo [SERVER] Starting Isabelle Chat Server...
start /min "Isabelle Server" venv\Scripts\python.exe ChatServer\setup_and_run.py

:: 5. Wait a moment for FastAPI/Uvicorn to bind to port 8000
timeout /t 3 /nobreak >nul

:: 6. Launch the Godot Application
echo [APP] Launching Desktop Application...
start "" App\IsabelleUI12.exe

exit
@echo off
title Philippine Municipal Officials Directory & Multi-Format Exporter
cd /d "%~dp0.."

:: 1. Check Python installation
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ============================================================
    echo [ERROR] Python was not found on your system!
    echo Please install Python 3.8 or later from https://www.python.org
    echo and ensure "Add Python to PATH" is checked during installation.
    echo ============================================================
    echo.
    pause
    exit /b 1
)

:: 2. Check for required dependencies (openpyxl, docxtpl)
python -c "import openpyxl, docxtpl, docxcompose" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Installing required packages from requirements.txt...
    pip install -r requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Failed to install dependencies. Please run:
        echo   pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
)

:: 3. Launch Application
python app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred while running the application.
    pause
)

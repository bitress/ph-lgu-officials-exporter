@echo off
title Philippine Municipal Officials Directory and CSV Exporter
cd /d "%~dp0"
python app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred while running the application.
    pause
)

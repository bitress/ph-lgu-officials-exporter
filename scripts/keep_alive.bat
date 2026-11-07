@echo off
:: ============================================================
:: keep_alive.bat - Watchdog wrapper for the PH LGU web server
:: Runs the server and restarts it automatically if it crashes.
:: This is invoked by Windows Task Scheduler on logon.
:: ============================================================
setlocal enabledelayedexpansion
cd /d "%~dp0.."

:: Create logs directory if missing
if not exist "data\logs" mkdir "data\logs"

set LOGFILE=data\logs\web_server.log
set APP=app.py
set PORT=8000

:LOOP
echo [%DATE% %TIME%] Starting PH LGU web server on port %PORT%... >> "%LOGFILE%"
python "%APP%" --web %PORT% >> "%LOGFILE%" 2>&1
echo [%DATE% %TIME%] Server stopped (exit code: %ERRORLEVEL%). Restarting in 5 seconds... >> "%LOGFILE%"
timeout /nobreak /t 5 >nul
goto LOOP

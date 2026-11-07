@echo off
:: manage_server.bat - Quick control for the PH LGU web server task
:: Double-click, or run from command line: manage_server.bat [start|stop|restart|status|logs|open]
cd /d "%~dp0.."
set TASK=PHLGUOfficials-WebServer
set PORT=8000
set HAS_ARG=0

if /i "%1"=="start"   set HAS_ARG=1 & call :START & goto DONE
if /i "%1"=="stop"    set HAS_ARG=1 & call :STOP & goto DONE
if /i "%1"=="restart" set HAS_ARG=1 & call :RESTART & goto DONE
if /i "%1"=="status"  set HAS_ARG=1 & call :STATUS & goto DONE
if /i "%1"=="logs"    set HAS_ARG=1 & call :LOGS & goto DONE
if /i "%1"=="open"    set HAS_ARG=1 & call :OPEN & goto DONE

:MENU
echo.
echo  Philippine LGU Officials Web Server - Control Panel
echo  =====================================================
echo.
echo  [1] Start server
echo  [2] Stop server
echo  [3] Restart server
echo  [4] Show status
echo  [5] View logs
echo  [6] Open in browser  (http://localhost:%PORT%)
echo  [7] Install as auto-start (run as Admin)
echo  [Q] Quit
echo.
set /p choice=Enter choice: 

if "%choice%"=="1" call :START & goto DONE
if "%choice%"=="2" call :STOP & goto DONE
if "%choice%"=="3" call :RESTART & goto DONE
if "%choice%"=="4" call :STATUS & goto DONE
if "%choice%"=="5" call :LOGS & goto DONE
if "%choice%"=="6" call :OPEN & goto DONE
if "%choice%"=="7" call :INSTALL & goto DONE
if /i "%choice%"=="Q" exit /b 0
echo Invalid choice.
pause
exit /b 1

:START
echo Starting web server...
schtasks /Run /TN "%TASK%" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Task not installed. Starting directly...
    start "" /B pythonw app.py --web %PORT%
    echo Server started on http://localhost:%PORT%
) else (
    echo Server started via Task Scheduler.
)
exit /b 0

:STOP
echo Stopping web server...
schtasks /End /TN "%TASK%" >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq *LGU*" >nul 2>&1
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":%PORT% "') do taskkill /F /PID %%a >nul 2>&1
echo Server stopped.
exit /b 0

:RESTART
call :STOP
ping -n 3 127.0.0.1 >nul
call :START
exit /b 0

:STATUS
echo.
echo Task Scheduler status:
schtasks /Query /TN "%TASK%" /FO LIST 2>nul | findstr /i "status last"
echo.
echo Port %PORT% listener:
netstat -ano | findstr ":%PORT% " | findstr "LISTENING"
if %ERRORLEVEL% NEQ 0 echo (no process listening on port %PORT%)
exit /b 0

:LOGS
if exist "data\logs\web_server.log" (
    type "data\logs\web_server.log" | more
) else (
    echo No log file found at data\logs\web_server.log
)
exit /b 0

:OPEN
start "" "http://localhost:%PORT%"
exit /b 0

:INSTALL
echo Running installer as Administrator...
powershell -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0install_service.ps1""' -Verb RunAs"
exit /b 0

:DONE
echo.
if "%HAS_ARG%"=="0" pause
exit /b 0

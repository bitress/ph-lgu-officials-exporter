' start_web_server.vbs
' Launches the PH LGU Officials web server silently (no console window).
' Registered in HKCU\Run to auto-start on every Windows logon.

Dim fso, scriptDir, pythonwExe, appScript, logFile, logDir, sh, cmd

Set fso       = CreateObject("Scripting.FileSystemObject")
scriptDir     = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))  ' project root (this file lives in scripts\)
appScript     = scriptDir & "\app.py"
logDir        = scriptDir & "\data\logs"
logFile       = logDir    & "\web_server.log"

' Ensure logs folder exists (silently ignore if already exists)
If Not fso.FolderExists(scriptDir & "\data") Then fso.CreateFolder(scriptDir & "\data")
If Not fso.FolderExists(logDir) Then fso.CreateFolder(logDir)

' Use pythonw.exe (no console). Path is patched by install_service.ps1 at install time.
pythonwExe = "C:\Users\Admin\AppData\Local\Python\pythoncore-3.14-64\pythonw.exe"
If Not fso.FileExists(pythonwExe) Then pythonwExe = "pythonw"

' Run the web server hidden, no console
cmd = pythonwExe & " """ & appScript & """ --web 8000 >> """ & logFile & """ 2>&1"

Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = scriptDir
sh.Run "cmd /c " & cmd, 0, False   ' 0 = hidden, False = async

Set sh  = Nothing
Set fso = Nothing

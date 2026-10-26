@echo off
setlocal
set "SCRIPT=%~dp0BatchRename.ps1"
if not exist "%SCRIPT%" (
    echo [BatchRename] Cannot find BatchRename.ps1 next to this launcher.
    echo Keep this .bat and BatchRename.ps1 in the same folder.
    echo.
    pause
    exit /b 1
)
start "" powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT%"
exit /b 0

@echo off
echo Starting URIT Data Export...
echo.
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\ExportFromLiveDB.ps1"
echo.
if errorlevel 1 (
    echo Error occurred. Check export_log.txt
) else (
    echo Done.
)
pause

@echo off
setlocal
cd /d "%~dp0"

echo Starting WorkbookLens on this computer only.
echo Your browser will open after the local service is ready.
echo A dedicated WorkbookLens console window will open.
echo Keep that window open. Press Ctrl+C there to stop WorkbookLens.
echo.

start "WorkbookLens" /D "%~dp0" "%~dp0WorkbookLens.exe" serve --open-browser --fallback-port
if errorlevel 1 (
    echo.
    echo WorkbookLens could not be started.
    pause
    exit /b 1
)

exit /b 0

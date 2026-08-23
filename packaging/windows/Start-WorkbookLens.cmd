@echo off
setlocal
cd /d "%~dp0"
start "" /D "%~dp0" "%~dp0WorkbookLens.exe"
if errorlevel 1 (
    exit /b 1
)
exit /b 0

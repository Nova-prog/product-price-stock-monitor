@echo off
setlocal
cd /d "%~dp0"

if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" "%~dp0monitor.py"
    exit /b %errorlevel%
)

python "%~dp0monitor.py"
exit /b %errorlevel%

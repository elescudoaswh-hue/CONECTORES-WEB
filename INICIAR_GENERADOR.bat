@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    py -3.12 -m venv .venv
    if errorlevel 1 (
        echo Instala Python 3.12 de 64 bits y el lanzador py para Windows.
        pause
        exit /b 1
    )
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    pause
    exit /b 1
)
".venv\Scripts\python.exe" -m windows_service
if errorlevel 1 pause

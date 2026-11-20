@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0runtime\python.exe" (
    echo Extrae todos los archivos del ZIP de Conectores antes de instalar.
    pause
    exit /b 1
)
"%~dp0runtime\python.exe" "%~dp0windows_local.py" --install
set "CONECTORES_RESULT=%ERRORLEVEL%"
pause
exit /b %CONECTORES_RESULT%

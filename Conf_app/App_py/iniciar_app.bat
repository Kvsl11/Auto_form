@echo off
REM Iniciador sem console para Hxg_auto

set "BASE_DIR=%~dp0"
set "PYTHONW=%BASE_DIR%..\app_py\Python313\Python313\pythonw.exe"
set "SCRIPT=%BASE_DIR%..\app_py\main.py"

if exist "%PYTHONW%" (
    start "" "%PYTHONW%" "%SCRIPT%"
) else (
    echo [ERRO] pythonw.exe não encontrado em: "%PYTHONW%"
    pause
    exit /b 1
)
exit /b 0

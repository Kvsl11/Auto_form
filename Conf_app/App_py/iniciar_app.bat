@echo off
REM 🚀 Iniciador silencioso do Hxg_auto (interface Tkinter)

REM Define diretório base (onde o BAT está salvo)
cd /d "%~dp0"

REM Caminhos absolutos
set "BASE_DIR=%~dp0"
set "PYTHONW=%BASE_DIR%Python313\pythonw.exe"
set "UPDATER=%BASE_DIR%updater.py"

REM ==========================
REM   VERIFICAÇÕES BÁSICAS
REM ==========================
if not exist "%PYTHONW%" (
    echo ❌ ERRO: Python nao encontrado em "%PYTHONW%"
    pause
    exit /b 1
)

if not exist "%UPDATER%" (
    echo ❌ ERRO: updater.py nao encontrado em "%UPDATER%"
    pause
    exit /b 1
)

REM ==========================
REM   EXECUÇÃO SILENCIOSA
REM ==========================
echo 🔄 Iniciando o atualizador (background)...

REM Usa START para rodar sem travar o CMD e sem console
start "" "%PYTHONW%" "%UPDATER%"

REM Fecha imediatamente o CMD
exit /b 0

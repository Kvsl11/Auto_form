@echo off
title Gerar requirements.txt + Verificar Ambiente Python
setlocal enabledelayedexpansion

REM ============================================
REM   SCRIPT ÚNICO: GERAR E REPARAR AMBIENTE
REM   1. Gera requirements.txt automaticamente
REM   2. Verifica pip/SSL e reinstala pacotes
REM ============================================

REM === Caminho base ===
set "BASE_DIR=%~dp0"

REM === Caminhos principais ===
set "MAIN_FILE=%BASE_DIR%..\app_py\main.py"
set "REQ_FILE=%BASE_DIR%..\app_py\requirements.txt"
set "PYTHON_PATH=%BASE_DIR%..\app_py\Python313\Python313\python.exe"
set "LOG_FILE=%BASE_DIR%pip_repair_log.txt"
set "SSL_URL=https://github.com/python/cpython/raw/main/PCbuild/amd64"

echo ============================================
echo   GERADOR DE REQUIREMENTS E REPARADOR PYTHON
echo ============================================
echo.

REM ===============================
REM   PARTE 1 - GERAR REQUIREMENTS
REM ===============================

if not exist "%MAIN_FILE%" (
    if exist "%BASE_DIR%..\App_py\main.py" (
        set "MAIN_FILE=%BASE_DIR%..\App_py\main.py"
        set "REQ_FILE=%BASE_DIR%..\App_py\requirements.txt"
    ) else (
        echo [ERRO] Nao foi encontrado o arquivo main.py em:
        echo "%MAIN_FILE%"
        echo Estrutura esperada:
        echo   app_py\main.py
        echo   ou App_py\main.py
        pause
        exit /b 1
    )
)

echo [1/7] Gerando requirements.txt...
del "%REQ_FILE%" >nul 2>&1
echo.>"%REQ_FILE%"

for /f "tokens=2 delims= " %%i in ('findstr /r "^import ^from" "%MAIN_FILE%"') do (
    set "mod=%%i"
    for /f "delims=. tokens=1" %%a in ("!mod!") do set "mod=%%a"
    echo !mod!>>"%REQ_FILE%"
)

set "TMP_FILE=%REQ_FILE%.tmp"
type nul > "%TMP_FILE%"

set "CORE_MODULES=asyncio base64 calendar collections concurrent contextlib copy csv ctypes datetime difflib email enum fnmatch functools glob hashlib heapq http io itertools json linecache locale logging math mimetypes numbers operator os pathlib pickle pkgutil platform queue random re shutil signal socket sqlite3 ssl statistics string struct subprocess sys tempfile textwrap threading time tkinter types typing unicodedata urllib uuid warnings weakref xml zipfile webbrowser argparse inspect dataclasses"

for /f "usebackq delims=" %%M in ("%REQ_FILE%") do (
    set "mod=%%M"
    if not "!mod!"=="" (
        echo !CORE_MODULES! | findstr /i "\<!mod!\>" >nul
        if errorlevel 1 (
            findstr /ix "!mod!" "%TMP_FILE%" >nul || echo !mod!>>"%TMP_FILE%"
        )
    )
)

move /y "%TMP_FILE%" "%REQ_FILE%" >nul
echo setuptools>>"%REQ_FILE%"
echo wheel>>"%REQ_FILE%"

echo [OK] requirements.txt gerado com sucesso.
echo Local: %REQ_FILE%
echo.

REM ===============================
REM   PARTE 2 - VERIFICAR AMBIENTE
REM ===============================

echo [2/7] Verificando ambiente Python...
if not exist "%PYTHON_PATH%" (
    echo [ERRO] Python nao encontrado em:
    echo "%PYTHON_PATH%"
    pause
    exit /b 1
)

echo [3/7] Verificando pip e SSL...
"%PYTHON_PATH%" -m pip -V >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Pip ausente. Tentando reparar...
    "%PYTHON_PATH%" -m ensurepip --upgrade >nul 2>&1
    "%PYTHON_PATH%" -m pip install --upgrade pip setuptools wheel >nul 2>&1
)

"%PYTHON_PATH%" -m pip install requests --dry-run >nul 2>&1
if %errorlevel% neq 0 (
    echo [ALERTA] Problema SSL detectado. Tentando reparar...
    for %%D in (libssl-3.dll libcrypto-3.dll) do (
        if not exist "%BASE_DIR%%%D" (
            echo Baixando %%D ...
            powershell -Command "Invoke-WebRequest -Uri '%SSL_URL%/%%D' -OutFile '%BASE_DIR%%%D'" >nul 2>&1
        )
        copy /Y "%BASE_DIR%%%D" "%BASE_DIR%" >nul 2>&1
        copy /Y "%BASE_DIR%%%D" "%BASE_DIR%..\app_py\Python313" >nul 2>&1
    )
    echo [INFO] DLLs SSL adicionadas (libssl-3.dll / libcrypto-3.dll)
)

echo [4/7] Corrigindo requirements.txt...
set "TMP_FILE=%BASE_DIR%req_fixed.txt"
copy "%REQ_FILE%" "%TMP_FILE%" >nul 2>&1
powershell -Command "(Get-Content '%TMP_FILE%') -replace '\bPIL\b','Pillow' -replace '\bfitz\b','PyMuPDF' | Set-Content '%TMP_FILE%' -Encoding UTF8"
echo [INFO] Corrigido: PIL → Pillow / fitz → PyMuPDF

echo.
echo Pacotes detectados:
type "%TMP_FILE%"
echo.
pause

echo [5/7] Removendo pacotes existentes...
"%PYTHON_PATH%" -m pip freeze > "%BASE_DIR%all.txt" 2>nul
for /f "usebackq delims=" %%P in ("%BASE_DIR%all.txt") do (
    echo Removendo %%P ...
    "%PYTHON_PATH%" -m pip uninstall -y %%P >> "%LOG_FILE%" 2>&1
)
del "%BASE_DIR%all.txt" >nul 2>&1

echo [6/7] Atualizando pip, setuptools e wheel...
"%PYTHON_PATH%" -m pip install --upgrade pip setuptools wheel >> "%LOG_FILE%" 2>&1

echo.
echo Deseja instalar os pacotes do requirements.txt agora? (Y/N)
set /p USER_CHOICE="> "
if /i "%USER_CHOICE%"=="N" (
    echo [INFO] Instalacao cancelada.
    echo Log salvo em: %LOG_FILE%
    pause
    exit /b 0
)
if /i not "%USER_CHOICE%"=="Y" (
    echo [ERRO] Opcao invalida.
    pause
    exit /b 1
)

echo [7/7] Instalando pacotes...
"%PYTHON_PATH%" -m pip install -r "%TMP_FILE%" >> "%LOG_FILE%" 2>&1

if %errorlevel% neq 0 (
    echo [AVISO] Erros durante a instalacao. Veja o log:
    echo %LOG_FILE%
) else (
    echo [SUCESSO] Todos os pacotes foram instalados corretamente!
)

echo.
echo ===============================================
echo Processo concluido!
echo Log salvo em: %LOG_FILE%
echo ===============================================
pause
exit /b 0

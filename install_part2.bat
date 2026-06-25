@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "INSTALL_DIR=C:\fam-app"

echo.
echo  ================================================================
echo   FAM App - Instalador (Parte 2 de 2)
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.
echo  Diretorio de instalacao: %INSTALL_DIR%
echo.

if not exist "%INSTALL_DIR%\app.py" (
    echo [ERRO] app.py nao encontrado em %INSTALL_DIR%.
    echo        Execute INSTALAR.bat primeiro.
    pause & exit /b 1
)

set PYTHON_CMD=
py --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=py
if "!PYTHON_CMD!"=="" (
    python --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=python
)
if "!PYTHON_CMD!"=="" (
    echo [ERRO] Python nao encontrado. Instale o Python e marque "Add to PATH".
    pause & exit /b 1
)
echo [OK] Python: !PYTHON_CMD!
!PYTHON_CMD! --version

set PYTHONIOENCODING=utf-8

echo.
echo Instalando dependencias...
!PYTHON_CMD! -m pip install --upgrade pip --quiet
!PYTHON_CMD! -m pip install flask openpyxl pdfplumber requests --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause & exit /b 1
)
echo [OK] Dependencias instaladas

if not exist "%INSTALL_DIR%\data"    mkdir "%INSTALL_DIR%\data"
if not exist "%INSTALL_DIR%\uploads" mkdir "%INSTALL_DIR%\uploads"
if not exist "%INSTALL_DIR%\outputs" mkdir "%INSTALL_DIR%\outputs"
if not exist "%INSTALL_DIR%\logs"    mkdir "%INSTALL_DIR%\logs"
echo [OK] Pastas criadas

if not exist "%INSTALL_DIR%\config.json" (
    echo {"port": 5002} > "%INSTALL_DIR%\config.json"
    echo [OK] config.json criado
)

echo.
echo Criando atalho na area de trabalho...
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'FAM App.lnk')); $s.TargetPath = 'wscript.exe'; $s.Arguments = '""C:\fam-app\launch.vbs""'; $s.WorkingDirectory = 'C:\fam-app'; $s.WindowStyle = 7; if (Test-Path 'C:\fam-app\fam.ico') { $s.IconLocation = 'C:\fam-app\fam.ico,0' }; $s.Description = 'FAM App'; $s.Save()"
echo [OK] Atalho criado

echo.
echo  ================================================================
echo   Instalacao concluida!
echo   Use o atalho "FAM App" na area de trabalho para iniciar.
echo  ================================================================
echo.
pause

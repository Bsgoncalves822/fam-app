@echo off
chcp 65001 >nul
echo.
echo  ================================================================
echo   FAM App - Instalador (Parte 2 de 2)
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.

:: Derive install dir dynamically from this script's location
set "INSTALL_DIR=%~dp0"
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

:: Verify essential files
if not exist "%INSTALL_DIR%\app.py" (
    echo [ERRO] app.py nao encontrado em %INSTALL_DIR%.
    echo        Execute bootstrap.bat primeiro.
    pause & exit /b 1
)
if not exist "%INSTALL_DIR%\templates\index.html" (
    echo [ERRO] templates\index.html nao encontrado.
    pause & exit /b 1
)
if not exist "%INSTALL_DIR%\launch.vbs" (
    echo [ERRO] launch.vbs nao encontrado.
    pause & exit /b 1
)

:: Find Python — check py launcher, then python, then python3
set PYTHON_CMD=
py --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=py
if "%PYTHON_CMD%"=="" (
    python --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=python
)
if "%PYTHON_CMD%"=="" (
    python3 --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=python3
)
if "%PYTHON_CMD%"=="" (
    echo [ERRO] Python nao encontrado. Instale o Python e marque "Add to PATH".
    pause & exit /b 1
)
echo [OK] Python encontrado via: %PYTHON_CMD%
%PYTHON_CMD% --version

:: Set PYTHONIOENCODING to avoid encoding issues on PT-BR Windows
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

:: Install dependencies
echo.
echo Instalando dependencias Python...
%PYTHON_CMD% -m pip install --upgrade pip --quiet
%PYTHON_CMD% -m pip install flask openpyxl pdfplumber requests --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause & exit /b 1
)
echo [OK] Dependencias instaladas

:: Create runtime folders
if not exist "%INSTALL_DIR%\data"    mkdir "%INSTALL_DIR%\data"
if not exist "%INSTALL_DIR%\uploads" mkdir "%INSTALL_DIR%\uploads"
if not exist "%INSTALL_DIR%\outputs" mkdir "%INSTALL_DIR%\outputs"
if not exist "%INSTALL_DIR%\logs"    mkdir "%INSTALL_DIR%\logs"
echo [OK] Pastas criadas

:: Write config.json with defaults if not present
if not exist "%INSTALL_DIR%\config.json" (
    echo { > "%INSTALL_DIR%\config.json"
    echo     "port": 5002 >> "%INSTALL_DIR%\config.json"
    echo } >> "%INSTALL_DIR%\config.json"
    echo [OK] config.json criado
)

:: Patch launch.vbs to use correct Python command
echo.
echo Configurando launcher para usar: %PYTHON_CMD%
powershell -NoProfile -Command ^
    "(Get-Content '%INSTALL_DIR%\launch.vbs') -replace 'python app\.py', '%PYTHON_CMD% app.py' | Set-Content '%INSTALL_DIR%\launch.vbs' -Encoding UTF8"

:: Create desktop shortcut
echo.
echo Criando atalho na area de trabalho...
powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'FAM App.lnk')); $s.TargetPath = 'wscript.exe'; $s.Arguments = '\""%INSTALL_DIR%\launch.vbs\""'; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.WindowStyle = 7; if (Test-Path '%INSTALL_DIR%\fam.ico') { $s.IconLocation = '%INSTALL_DIR%\fam.ico' }; $s.Description = 'FAM App - Parser de Comprovantes'; $s.Save()"
echo [OK] Atalho "FAM App" criado na area de trabalho

echo.
echo  ================================================================
echo   Instalacao concluida!
echo   Clique em "FAM App" na area de trabalho para iniciar.
echo  ================================================================
echo.
pause

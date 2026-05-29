@echo off
chcp 65001 >nul
echo.
echo  ================================================================
echo   FAM App - Instalador
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.

set INSTALL_DIR=C:\fam-app

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Instale o Python 3.10+ e tente novamente.
    echo        Download: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python encontrado

:: Install dependencies
echo.
echo Instalando dependencias...
pip install flask pandas openpyxl pdfplumber requests --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas

:: Create folders
if not exist "%INSTALL_DIR%\data" mkdir "%INSTALL_DIR%\data"
if not exist "%INSTALL_DIR%\uploads" mkdir "%INSTALL_DIR%\uploads"
if not exist "%INSTALL_DIR%\outputs" mkdir "%INSTALL_DIR%\outputs"
echo [OK] Pastas criadas

:: Generate reference data if DuplicateHandler exists
if exist "C:\Users\%USERNAME%\Downloads\FAM\Duplicate handling\DuplicateHandler.xlsx" (
    echo.
    echo Gerando dados de referencia...
    python "%INSTALL_DIR%\generate_data.py"
    echo [OK] Dados de referencia gerados
) else (
    echo [AVISO] DuplicateHandler.xlsx nao encontrado - importe os CSVs manualmente na interface.
)

:: Create desktop shortcut
echo.
echo Criando atalho na area de trabalho...
powershell -command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'FAM App.lnk')); $s.TargetPath = 'wscript.exe'; $s.Arguments = 'C:\fam-app\launch.vbs'; $s.WorkingDirectory = 'C:\fam-app'; $s.WindowStyle = 7; $s.Description = 'FAM Reconciliacao Bancaria'; $s.Save()"
echo [OK] Atalho criado na area de trabalho

echo.
echo  ================================================================
echo   Instalacao concluida!
echo   Use o atalho "FAM App" na area de trabalho para iniciar.
echo  ================================================================
echo.
pause

@echo off
chcp 65001 >nul
echo.
echo  ================================================================
echo   FAM App - Instalador (Parte 2 de 2)
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.

set INSTALL_DIR=C:\fam-app

:: Verify this is a git checkout (required for auto-update)
if not exist "%INSTALL_DIR%\.git" (
    echo [AVISO] %INSTALL_DIR% nao e um repositorio git.
    echo         A atualizacao automatica nao funcionara.
    echo         Use bootstrap.bat para uma instalacao completa via git.
    echo.
)

:: Verify Python now available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python ainda nao encontrado no PATH.
    echo        Reinstale o Python e marque "Add Python to PATH".
    pause
    exit /b 1
)
echo [OK] Python encontrado
python --version

:: Install dependencies
echo.
echo Instalando dependencias Python...
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

:: Generate reference data
echo.
echo Gerando dados de referencia...
python "%INSTALL_DIR%\generate_data.py"

:: Create desktop shortcut
echo.
echo Criando atalho na area de trabalho...
powershell -command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'FAM App.lnk')); $s.TargetPath = 'wscript.exe'; $s.Arguments = 'C:\fam-app\launch.vbs'; $s.WorkingDirectory = 'C:\fam-app'; $s.WindowStyle = 7; $s.Description = 'FAM Reconciliacao Bancaria'; $s.Save()"
echo [OK] Atalho "FAM App" criado na area de trabalho

echo.
echo  ================================================================
echo   Instalacao concluida!
echo   Use o atalho "FAM App" na area de trabalho para iniciar.
echo  ================================================================
echo.
pause

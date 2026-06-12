@echo off
chcp 65001 >nul
echo.
echo  ================================================================
echo   FAM App - Instalador (Parte 2 de 2)
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.

:: Derive install dir from this script's own location (no hardcoded path)
set "INSTALL_DIR=%~dp0"
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

:: Hard check: app code must be present before continuing
if not exist "%INSTALL_DIR%\app.py" (
    echo [ERRO] %INSTALL_DIR%\app.py nao encontrado.
    echo        A instalacao nao pode continuar. Execute bootstrap.bat
    echo        para obter o codigo do FAM App corretamente.
    pause
    exit /b 1
)
if not exist "%INSTALL_DIR%\templates\index.html" (
    echo [ERRO] %INSTALL_DIR%\templates\index.html nao encontrado.
    echo        A instalacao nao pode continuar. Execute bootstrap.bat
    echo        para obter o codigo do FAM App corretamente.
    pause
    exit /b 1
)
if not exist "%INSTALL_DIR%\launch.vbs" (
    echo [ERRO] %INSTALL_DIR%\launch.vbs nao encontrado.
    echo        A instalacao nao pode continuar. Execute bootstrap.bat
    echo        para obter o codigo do FAM App corretamente.
    pause
    exit /b 1
)

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
if not exist "%INSTALL_DIR%\launch.vbs" (
    echo [ERRO] launch.vbs nao encontrado - atalho nao sera criado.
    pause
    exit /b 1
)
echo Criando atalho na area de trabalho...
if exist "%INSTALL_DIR%\fam.ico" (
    powershell -command "$q=[char]34; $ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'FAM App.lnk')); $s.TargetPath = 'wscript.exe'; $s.Arguments = $q + '%INSTALL_DIR%\launch.vbs' + $q; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.WindowStyle = 7; $s.IconLocation = '%INSTALL_DIR%\fam.ico'; $s.Description = 'FAM Reconciliacao Bancaria'; $s.Save()"
) else (
    echo [AVISO] fam.ico nao encontrado - atalho usara icone padrao.
    powershell -command "$q=[char]34; $ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'FAM App.lnk')); $s.TargetPath = 'wscript.exe'; $s.Arguments = $q + '%INSTALL_DIR%\launch.vbs' + $q; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.WindowStyle = 7; $s.Description = 'FAM Reconciliacao Bancaria'; $s.Save()"
)
echo [OK] Atalho "FAM App" criado na area de trabalho

echo.
echo  ================================================================
echo   Instalacao concluida!
echo   Use o atalho "FAM App" na area de trabalho para iniciar.
echo  ================================================================
echo.
pause

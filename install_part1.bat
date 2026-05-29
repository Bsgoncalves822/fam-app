@echo off
chcp 65001 >nul
echo.
echo  ================================================================
echo   FAM App - Instalador (Parte 1 de 2)
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.
echo  Esta parte vai instalar o Python no seu computador.
echo  Uma janela de instalacao vai abrir - clique em:
echo.
echo    [x] Add Python to PATH
echo    [Install Now]
echo.
echo  Apos concluir, a Parte 2 abrira automaticamente.
echo.
pause

:: Check if Python already installed
python --version >nul 2>&1
if not errorlevel 1 (
    echo [OK] Python ja esta instalado. Pulando instalacao...
    goto run_part2
)

:: Download Python installer
echo Baixando Python 3.13...
powershell -command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.13.0/python-3.13.0-amd64.exe' -OutFile '%TEMP%\python_installer.exe'"
if errorlevel 1 (
    echo [ERRO] Falha ao baixar Python. Verifique sua conexao com a internet.
    pause
    exit /b 1
)

:: Run installer visibly so user can check "Add to PATH"
echo.
echo Abrindo instalador do Python...
start /wait "%TEMP%\python_installer.exe"

:run_part2
:: Open Part 2 in a new terminal with updated PATH
echo.
echo Abrindo Parte 2...
start cmd /k "cd /d C:\fam-app && install_part2.bat"
exit

@echo off
chcp 65001 >nul
echo.
echo  ================================================================
echo   FAM App - Instalador Inicial (Bootstrap)
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.
echo  Este script prepara a maquina do zero:
echo    1. Verifica/instala o Git
echo    2. Clona o repositorio para C:\fam-app
echo    3. Continua com a instalacao normal (Python, dependencias, etc.)
echo.
pause

set INSTALL_DIR=C:\fam-app
set REPO_URL=https://github.com/Bsgoncalves822/fam-app.git

:: --- Check Git ---
git --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo Git nao encontrado. Baixando instalador do Git...
    powershell -command "Invoke-WebRequest -Uri 'https://github.com/git-for-windows/git/releases/latest/download/Git-64-bit.exe' -OutFile '%TEMP%\git_installer.exe'"
    if errorlevel 1 (
        echo [ERRO] Falha ao baixar o Git. Verifique sua conexao com a internet.
        pause
        exit /b 1
    )
    echo Abrindo instalador do Git...
    start /wait "%TEMP%\git_installer.exe" /VERYSILENT /NORESTART
    echo [OK] Git instalado.
) else (
    echo [OK] Git ja esta instalado.
)

:: --- Clone or update repo ---
if exist "%INSTALL_DIR%\.git" (
    echo.
    echo [OK] Repositorio ja existe em %INSTALL_DIR%. Atualizando...
    cd /d "%INSTALL_DIR%"
    git pull
) else (
    if exist "%INSTALL_DIR%" (
        echo.
        echo [AVISO] %INSTALL_DIR% ja existe mas nao e um repositorio git.
        echo         Renomeie ou remova essa pasta antes de continuar, ou
        echo         este instalador pode sobrescrever arquivos existentes.
        pause
    )
    echo.
    echo Clonando repositorio para %INSTALL_DIR%...
    git clone "%REPO_URL%" "%INSTALL_DIR%"
    if errorlevel 1 (
        echo [ERRO] Falha ao clonar o repositorio.
        echo        Verifique sua conexao e as credenciais de acesso ao repositorio privado.
        pause
        exit /b 1
    )
    echo [OK] Repositorio clonado.
)

:: --- Continue with normal install ---
echo.
echo Continuando instalacao...
cd /d "%INSTALL_DIR%"
call install_part1.bat

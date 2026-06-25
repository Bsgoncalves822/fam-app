@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
echo.
echo  ================================================================
echo   FAM App - Instalador Inicial (Bootstrap)
echo   ORPROCON / Contadores Digitais
echo  ================================================================
echo.
echo  Este script prepara a maquina do zero:
echo    1. Verifica/instala o Git
echo    2. Obtem o codigo do FAM App (via git ou copia local)
echo    3. Continua com a instalacao normal (Python, dependencias, etc.)
echo.
pause

set INSTALL_DIR=C:\fam-app
set REPO_URL=https://github.com/Bsgoncalves822/fam-app.git
set GIT_INSTALLER_URL=https://github.com/git-for-windows/git/releases/download/v2.51.0.windows.1/Git-2.51.0-64-bit.exe
set SCRIPT_DIR=%~dp0

:: ================================================================
:: STEP 1: Garantir que o Git esta disponivel
:: ================================================================
git --version >nul 2>&1
if not errorlevel 1 (
    echo [OK] Git ja esta instalado.
    goto get_code
)

echo.
echo Git nao encontrado. Tentando instalar...
powershell -NoProfile -command "try { Invoke-WebRequest -Uri '%GIT_INSTALLER_URL%' -OutFile '%TEMP%\git_installer.exe' -UseBasicParsing } catch { exit 1 }"
if errorlevel 1 (
    echo [AVISO] Nao foi possivel baixar o Git automaticamente.
    goto get_code_offline
)

echo Instalando Git silenciosamente, aguarde...
start /wait "%TEMP%\git_installer.exe" /VERYSILENT /NORESTART /NOCANCEL /SP-

:: Refresh PATH for current session by re-checking common install dir
set "PATH=%PATH%;C:\Program Files\Git\cmd"
git --version >nul 2>&1
if errorlevel 1 (
    echo [AVISO] Git foi instalado mas nao esta disponivel nesta sessao.
    echo         Sera necessario reabrir este instalador apos reiniciar o terminal.
    goto get_code_offline
)
echo [OK] Git instalado com sucesso.


:: ================================================================
:: STEP 2: Obter o codigo (clone via git, ou copia offline)
:: ================================================================
:get_code
if exist "%INSTALL_DIR%\.git" (
    echo.
    echo [OK] Repositorio ja existe em %INSTALL_DIR%. Atualizando...
    cd /d "%INSTALL_DIR%"
    git pull
    goto code_ready
)

if exist "%INSTALL_DIR%" (
    echo.
    echo [AVISO] %INSTALL_DIR% ja existe mas nao e um repositorio git.
    echo         Esta pasta sera mantida; arquivos podem ser sobrescritos.
)

echo.
echo Clonando repositorio para %INSTALL_DIR%...
git clone "%REPO_URL%" "%INSTALL_DIR%"
if not errorlevel 1 (
    echo [OK] Repositorio clonado com sucesso.
    goto code_ready
)

echo [AVISO] Falha ao clonar via git.


:get_code_offline
echo.
echo Tentando usar copia local incluida neste pacote...
if not exist "%SCRIPT_DIR%fam-app-snapshot\app.py" (
    echo.
    echo [ERRO] Nao foi possivel obter o codigo do FAM App:
    echo          - Git nao pode ser instalado/usado
    echo          - Copia local (fam-app-snapshot) nao encontrada
    echo.
    echo        Verifique a conexao com a internet e tente novamente,
    echo        ou use um pacote de instalacao que inclua fam-app-snapshot.
    pause
    exit /b 1
)

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
echo Copiando arquivos para %INSTALL_DIR%...
xcopy "%SCRIPT_DIR%fam-app-snapshot\*" "%INSTALL_DIR%\" /E /I /Y /Q >nul
if errorlevel 1 (
    echo [ERRO] Falha ao copiar arquivos para %INSTALL_DIR%.
    pause
    exit /b 1
)
echo [OK] Arquivos copiados da copia local.
echo [AVISO] Esta instalacao NAO e um repositorio git.
echo         A atualizacao automatica nao funcionara enquanto o Git
echo         nao for instalado e este instalador for executado de novo.


:: ================================================================
:: STEP 3: Verificar que o codigo essencial esta presente
:: ================================================================
:code_ready
if not exist "%INSTALL_DIR%\app.py" (
    echo.
    echo [ERRO] %INSTALL_DIR%\app.py nao foi encontrado apos a obtencao do
    echo        codigo. A instalacao nao pode continuar.
    pause
    exit /b 1
)
if not exist "%INSTALL_DIR%\templates\index.html" (
    echo.
    echo [ERRO] %INSTALL_DIR%\templates\index.html nao foi encontrado.
    echo        A instalacao nao pode continuar.
    pause
    exit /b 1
)
echo [OK] Codigo do FAM App presente em %INSTALL_DIR%.


:: ================================================================
:: STEP 4: Continuar com a instalacao normal
:: ================================================================
echo.
echo Continuando instalacao...
cd /d "%INSTALL_DIR%"
cd /d "C:\fam-app" && call install_part2.bat

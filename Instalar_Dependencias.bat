@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ===============================================
echo  Instalando dependencias do Classificador de UPG
echo ===============================================
echo.

set "PY_VERSION=3.12.7"
set "PYTHON_CMD=python"

where python >nul 2>nul
if errorlevel 1 (
    echo Python nao foi encontrado nesta maquina.
    echo Vou baixar e instalar o Python !PY_VERSION! automaticamente
    echo ^(instalacao so para o seu usuario, nao precisa ser administrador^).
    echo.

    set "PY_INSTALLER=%TEMP%\python-instalador-upg.exe"
    set "PY_URL=https://www.python.org/ftp/python/!PY_VERSION!/python-!PY_VERSION!-amd64.exe"

    echo Baixando Python !PY_VERSION! de python.org ...
    curl -fSL -o "!PY_INSTALLER!" "!PY_URL!"
    if errorlevel 1 (
        echo Tentando um jeito alternativo de baixar ^(PowerShell^)...
        powershell -NoProfile -Command "try { Invoke-WebRequest -Uri '!PY_URL!' -OutFile '!PY_INSTALLER!' } catch { exit 1 }"
    )
    if not exist "!PY_INSTALLER!" (
        echo.
        echo ===============================================
        echo  ERRO: nao foi possivel baixar o instalador do Python.
        echo  Verifique sua conexao com a internet, ou instale
        echo  manualmente em https://www.python.org/downloads/
        echo  ^(marque "Add python.exe to PATH" na instalacao^)
        echo  e rode este arquivo de novo depois.
        echo ===============================================
        pause
        exit /b 1
    )

    echo Instalando Python ^(sem janelas, so para o seu usuario^)...
    "!PY_INSTALLER!" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
    if errorlevel 1 (
        echo.
        echo ===============================================
        echo  ERRO: a instalacao do Python falhou.
        echo  Instale manualmente em https://www.python.org/downloads/
        echo  ^(marque "Add python.exe to PATH"^) e rode este
        echo  arquivo de novo.
        echo ===============================================
        pause
        exit /b 1
    )
    del "!PY_INSTALLER!" >nul 2>nul

    rem O instalador atualiza o PATH do Windows, mas esta janela de CMD
    rem ja aberta nao enxerga essa mudanca sozinha (so novas janelas veem
    rem o PATH atualizado). Por isso procuramos aqui o python.exe recem
    rem instalado na pasta padrao do usuario, para continuar sem precisar
    rem fechar e abrir este arquivo de novo.
    set "PYTHON_CMD="
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
        if exist "%%D\python.exe" set "PYTHON_CMD=%%D\python.exe"
    )

    if not defined PYTHON_CMD (
        echo.
        echo Python foi instalado com sucesso, mas esta janela nao
        echo consegue usa-lo ainda. Feche esta janela e clique de novo
        echo em "Instalar_Dependencias.bat" para continuar.
        pause
        exit /b 0
    )

    echo Python instalado com sucesso!
    echo.
)

"!PYTHON_CMD!" -m pip install --upgrade pip
"!PYTHON_CMD!" -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ===============================================
    echo  Ocorreu um erro durante a instalacao.
    echo.
    echo  Causas comuns:
    echo  - Sem internet no momento, ou PyPI bloqueado pelo proxy da empresa.
    echo  - Seguranca de rede corporativa interceptando o HTTPS ^(ex:
    echo    Netskope, Zscaler^) e o Python nao reconhecendo o certificado
    echo    dela. Aparece como erro mencionando "TLS CA certificate" ou
    echo    um caminho dentro de C:\WINDOWS\IMECache\. Nesse caso, fale
    echo    com a TI: eh uma configuracao do certificado nessa maquina,
    echo    nao um problema deste programa.
    echo  - Pouco espaco em disco.
    echo.
    echo  Copie a mensagem acima e peca ajuda.
    echo ===============================================
    pause
    exit /b 1
)

echo.
echo ===============================================
echo  Instalacao concluida com sucesso!
echo  Agora voce ja pode abrir "Abrir_Classificador_UPG.bat"
echo ===============================================
pause

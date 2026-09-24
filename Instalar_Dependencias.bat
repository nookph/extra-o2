@echo off
setlocal
cd /d "%~dp0"

echo ===============================================
echo  Instalando dependencias do Classificador de UPG
echo ===============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERRO: Python nao foi encontrado nesta maquina.
    echo.
    echo Instale o Python em https://www.python.org/downloads/
    echo IMPORTANTE: na tela de instalacao, marque a caixa
    echo "Add python.exe to PATH" antes de clicar em Install.
    echo.
    echo Depois de instalar, rode este arquivo de novo.
    echo.
    pause
    exit /b 1
)

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ===============================================
    echo  Ocorreu um erro durante a instalacao.
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

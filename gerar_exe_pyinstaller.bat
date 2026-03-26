@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "ENV_NAME=colaboradores-python"
set "PY_CMD="
set "PI_CMD="

echo ================================================
echo Build do projeto Colaboradores com PyInstaller
echo Ambiente esperado: %ENV_NAME%
echo ================================================
echo.

where conda >nul 2>nul
if %errorlevel%==0 (
    echo Conda encontrado. Tentando usar o ambiente %ENV_NAME%...
    conda run -n %ENV_NAME% python -c "import fitz, pytesseract, PIL, openpyxl, PyInstaller" >nul 2>nul
    if %errorlevel%==0 (
        set "PY_CMD=conda run -n %ENV_NAME% python"
        set "PI_CMD=conda run -n %ENV_NAME% python -m PyInstaller"
        echo Ambiente %ENV_NAME% validado com sucesso.
    ) else (
        echo.
        echo ERRO: o ambiente %ENV_NAME% nao foi encontrado ou esta incompleto.
        echo Crie ou atualize o ambiente com:
        echo   conda env create -f environment.yml
        echo ou
        echo   conda env update -f environment.yml --prune
        echo.
        pause
        exit /b 1
    )
) else (
    echo Conda nao foi encontrado no PATH. Tentando usar o Python atual...
    python -c "import fitz, pytesseract, PIL, openpyxl, PyInstaller" >nul 2>nul
    if %errorlevel%==0 (
        set "PY_CMD=python"
        set "PI_CMD=python -m PyInstaller"
        echo Python atual validado com sucesso.
    ) else (
        echo.
        echo ERRO: nao foi possivel validar as dependencias com o Python atual.
        echo Recomendacao: use o ambiente Conda definido em environment.yml.
        echo.
        pause
        exit /b 1
    )
)

echo.
echo Limpando build anterior...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist build\spec rmdir /s /q build\spec

echo.
echo Gerando executavel...
if exist "tesseract" (
    call %PI_CMD% --noconfirm --clean --windowed --name Colaboradores ^
    --specpath build\spec ^
    --exclude-module PyQt5 ^
    --exclude-module PySide6 ^
    --exclude-module PySide2 ^
    --exclude-module PyQt6 ^
    --exclude-module matplotlib ^
    --exclude-module IPython ^
    --exclude-module pandas ^
    --add-data "tesseract;tesseract" ^
    colaboradores_local.py
) else (
    echo Pasta tesseract nao encontrada ao lado do projeto.
    echo O build sera gerado sem empacotar o OCR portatil.
    call %PI_CMD% --noconfirm --clean --windowed --name Colaboradores ^
    --specpath build\spec ^
    --exclude-module PyQt5 ^
    --exclude-module PySide6 ^
    --exclude-module PySide2 ^
    --exclude-module PyQt6 ^
    --exclude-module matplotlib ^
    --exclude-module IPython ^
    --exclude-module pandas ^
    colaboradores_local.py
)

if errorlevel 1 (
    echo.
    echo FALHA ao gerar o executavel.
    pause
    exit /b 1
)

if exist "dist\Colaboradores" (
    if exist "iniciar_colaboradores.bat" copy /y "iniciar_colaboradores.bat" "dist\Colaboradores\iniciar_colaboradores.bat" >nul
    echo.
    echo Sucesso. Saida gerada em: dist\Colaboradores
) else (
    echo.
    echo O PyInstaller terminou sem criar a pasta esperada em dist\Colaboradores.
    pause
    exit /b 1
)

echo.
echo Build concluido.
pause

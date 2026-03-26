@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "ENV_NAME=colaboradores-python"
set "USE_CONDA=0"

echo ================================================
echo Build do projeto Colaboradores com PyInstaller
echo Ambiente esperado: %ENV_NAME%
echo ================================================
echo.

where conda >nul 2>nul
if errorlevel 1 goto USE_CURRENT_PYTHON

echo Conda encontrado. Tentando usar o ambiente %ENV_NAME%...
call conda run -n %ENV_NAME% python -c "import fitz, pytesseract, PIL, openpyxl, PyInstaller, cryptography" >nul 2>nul
if errorlevel 1 goto CONDA_ENV_ERROR

set "USE_CONDA=1"
echo Ambiente %ENV_NAME% validado com sucesso.
goto BUILD

:USE_CURRENT_PYTHON
echo Conda nao foi encontrado no PATH. Tentando usar o Python atual...
python -c "import fitz, pytesseract, PIL, openpyxl, PyInstaller, cryptography" >nul 2>nul
if errorlevel 1 goto PYTHON_ENV_ERROR

echo Python atual validado com sucesso.
goto BUILD

:CONDA_ENV_ERROR
echo.
echo ERRO: o ambiente %ENV_NAME% nao foi encontrado ou esta incompleto.
echo Crie ou atualize o ambiente com:
echo   conda env create -f environment.yml
echo ou
echo   conda env update -f environment.yml --prune
echo.
pause
exit /b 1

:PYTHON_ENV_ERROR
echo.
echo ERRO: nao foi possivel validar as dependencias com o Python atual.
echo Recomendacao: use o ambiente Conda definido em environment.yml.
echo.
pause
exit /b 1

:BUILD
echo.
echo Limpando build anterior...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo Gerando executavel...

if exist "tesseract" (
    if "%USE_CONDA%"=="1" (
        call conda run -n %ENV_NAME% python -m PyInstaller --noconfirm --clean --windowed --name Colaboradores ^
        --exclude-module PyQt5 ^
        --exclude-module PySide6 ^
        --exclude-module PySide2 ^
        --exclude-module PyQt6 ^
        --exclude-module matplotlib ^
        --exclude-module IPython ^
        --exclude-module pandas ^
        --add-data "tesseract;tesseract" ^
        --add-data "public_key.pem;." ^
        colaboradores_local.py
    ) else (
        call python -m PyInstaller --noconfirm --clean --windowed --name Colaboradores ^
        --specpath build\spec ^
        --exclude-module PyQt5 ^
        --exclude-module PySide6 ^
        --exclude-module PySide2 ^
        --exclude-module PyQt6 ^
        --exclude-module matplotlib ^
        --exclude-module IPython ^
        --exclude-module pandas ^
        --add-data "tesseract;tesseract" ^
        --add-data "public_key.pem;." ^
        colaboradores_local.py
    )
) else (
    echo Pasta tesseract nao encontrada ao lado do projeto.
    echo O build sera gerado sem empacotar o OCR portatil.
    if "%USE_CONDA%"=="1" (
        call conda run -n %ENV_NAME% python -m PyInstaller --noconfirm --clean --windowed --name Colaboradores ^
        --exclude-module PyQt5 ^
        --exclude-module PySide6 ^
        --exclude-module PySide2 ^
        --exclude-module PyQt6 ^
        --exclude-module matplotlib ^
        --exclude-module IPython ^
        --exclude-module pandas ^
        --add-data "public_key.pem;." ^
        colaboradores_local.py
    ) else (
        call python -m PyInstaller --noconfirm --clean --windowed --name Colaboradores ^
        --specpath build\spec ^
        --exclude-module PyQt5 ^
        --exclude-module PySide6 ^
        --exclude-module PySide2 ^
        --exclude-module PyQt6 ^
        --exclude-module matplotlib ^
        --exclude-module IPython ^
        --exclude-module pandas ^
        --add-data "public_key.pem;." ^
        colaboradores_local.py
    )
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
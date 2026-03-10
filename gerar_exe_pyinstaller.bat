@echo off
setlocal
cd /d "%~dp0"

echo Limpando build anterior...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Colaboradores.spec del /f /q Colaboradores.spec

echo Instalando dependencias de empacotamento...
python -m pip install --upgrade pip
python -m pip install pyinstaller pymupdf pytesseract pillow openpyxl

echo Gerando executavel...
if exist "tesseract" (
    pyinstaller --noconfirm --clean --windowed --name Colaboradores ^
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
    pyinstaller --noconfirm --clean --windowed --name Colaboradores ^
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
    echo O PyInstaller terminou sem criar a pasta esperada em dist\Colaboradores
)

pause

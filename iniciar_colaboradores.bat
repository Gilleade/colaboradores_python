@echo off
setlocal
cd /d "%~dp0"

if exist "Colaboradores.exe" (
    start "" "%~dp0Colaboradores.exe"
    exit /b 0
)

if exist "python\python.exe" (
    start "" "%~dp0python\python.exe" "%~dp0colaboradores_local.py"
    exit /b 0
)

echo Nao foi encontrado Colaboradores.exe nem python portatil em .\python\python.exe
pause

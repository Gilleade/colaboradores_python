@echo off
setlocal
cd /d "%~dp0"

if exist "Colaboradores.exe" (
    start "" "%~dp0Colaboradores.exe"
    exit /b 0
)

if exist "python\python.exe" if exist "colaboradores_local.py" (
    start "" "%~dp0python\python.exe" "%~dp0colaboradores_local.py"
    exit /b 0
)

echo Nao foi encontrado Colaboradores.exe nesta pasta.
echo.
echo Se esta for uma versao portatil gerada pelo build, execute o arquivo Colaboradores.exe.
echo O fallback com python\python.exe so funciona em uma estrutura complementar montada manualmente.
pause

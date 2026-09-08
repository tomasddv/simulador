@echo off
setlocal
pushd "%~dp0"
title Simulador MULTI-SKU ERP - Puerto 8051

echo ======================================================
echo   SIMULADOR COMERCIAL - VERSION MULTI-SKU
echo ======================================================
echo.

set "PY_CMD="
where py >nul 2>&1
if not errorlevel 1 set "PY_CMD=py"
if not defined PY_CMD (
    where python >nul 2>&1
    if not errorlevel 1 set "PY_CMD=python"
)

if not defined PY_CMD (
    echo ERROR: No se encontro Python instalado.
    echo Instala Python 3 y marca la opcion "Add Python to PATH".
    echo.
    pause
    popd
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual...
    %PY_CMD% -m venv ".venv"
    if errorlevel 1 goto :venverror
)

set "VENV_PY=%CD%\.venv\Scripts\python.exe"

echo Verificando dependencias...
"%VENV_PY%" -c "import dash, xlsxwriter" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias por primera vez...
    "%VENV_PY%" -m pip install -r "requirements.txt"
    if errorlevel 1 goto :piperror
)

echo.
echo Servidor: http://127.0.0.1:8051
echo Esta ventana debe permanecer abierta mientras uses el Dash.
echo.
echo Esperando que el servidor responda antes de abrir Chrome...

start "" /min powershell -NoProfile -ExecutionPolicy Bypass -Command "$u='http://127.0.0.1:8051'; for($i=0;$i -lt 60;$i++){try{$r=Invoke-WebRequest -UseBasicParsing -Uri $u -TimeoutSec 1; if($r.StatusCode -ge 200){Start-Process $u; exit}}catch{}; Start-Sleep -Seconds 1}"

"%VENV_PY%" "app.py"

echo.
echo ======================================================
echo El servidor se detuvo.
echo Si arriba aparece un error, sacale una captura y mandamela.
echo ======================================================
pause
popd
exit /b 0

:venverror
echo.
echo ERROR: No se pudo crear el entorno virtual de Python.
pause
popd
exit /b 1

:piperror
echo.
echo ERROR: No se pudieron instalar las dependencias.
echo Revisa la conexion a Internet o el acceso de pip en tu PC.
pause
popd
exit /b 1

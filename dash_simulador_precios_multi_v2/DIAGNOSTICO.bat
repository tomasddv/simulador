@echo off
setlocal
pushd "%~dp0"
echo ===== DIAGNOSTICO MULTI-SKU =====
echo Carpeta: %CD%
echo.
where py
where python
echo.
if exist ".venv\Scripts\python.exe" (
  echo Entorno virtual encontrado.
  ".venv\Scripts\python.exe" --version
  ".venv\Scripts\python.exe" -c "import dash; print('Dash:', dash.__version__); import xlsxwriter; print('XlsxWriter OK'); from app import CATALOG, EXPIRY; print('Catalogo SKU:', len(CATALOG)); print('Vencimientos:', len(EXPIRY))"
) else (
  echo No existe .venv todavia.
)
echo.
pause
popd

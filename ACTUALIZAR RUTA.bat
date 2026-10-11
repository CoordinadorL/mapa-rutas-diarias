@echo off
REM Doble clic: abre la ventana para generar la ruta del dia desde el
REM archivo base (hoja SABANA DE VENTA) y publicarla en el mapa.
REM Ver generar_rutas.py.
where pythonw >nul 2>nul
if errorlevel 1 (
  python "%~dp0generar_rutas.py"
  if errorlevel 1 pause
  exit /b
)
start "" pythonw "%~dp0generar_rutas.py"

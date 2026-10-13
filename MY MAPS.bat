@echo off
REM Doble clic: abre la ventana de My Maps (subir la ruta del dia y ver
REM los links de los ultimos 30 dias). Ver ventana_mymaps.py.
where pythonw >nul 2>nul
if errorlevel 1 (
  python "%~dp0ventana_mymaps.py"
  if errorlevel 1 pause
  exit /b
)
start "" pythonw "%~dp0ventana_mymaps.py"

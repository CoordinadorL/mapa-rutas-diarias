@echo off
REM Doble clic: deja corriendo (escondido) el servicio que usa la app web
REM OPAUSTRO para "Generar mapas" y "My Maps". Ver servicio_web.py.
REM No hace falta cerrarlo; si ya estaba abierto, no pasa nada.
where pythonw >nul 2>nul
if errorlevel 1 (
  start "" /min python "%~dp0servicio_web.py"
  exit /b
)
start "" pythonw "%~dp0servicio_web.py"

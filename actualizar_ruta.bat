@echo off
REM Arrastra la carpeta del dia (ej: "15MARTES 15") sobre este archivo,
REM o ejecutalo y escribe la ruta cuando la pida.
setlocal

set CARPETA=%~1
if "%CARPETA%"=="" (
  set /p CARPETA="Ruta de la carpeta del dia (ej: C:\RUTAS DIARIAS\15MARTES 15): "
)

python "%~dp0convertir.py" "%CARPETA%"

echo.
echo Listo: data\latest.json actualizado (no se sube a git, ver README.md
echo para como entregarlo sin exponer datos de clientes).
echo.
pause

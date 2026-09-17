@echo off
REM Copia data\latest.json (generado por actualizar_ruta.bat) al repo
REM PRIVADO mapa-rutas-datos y lo sube. Este es el unico paso que sube
REM datos reales de clientes a algun lado, por eso es privado.
setlocal

set ORIGEN=%~dp0data\latest.json
set DESTINO=C:\MAPA-RUTAS-DATOS

if not exist "%ORIGEN%" (
  echo No existe %ORIGEN%
  echo Corre primero actualizar_ruta.bat con la carpeta del dia.
  pause
  exit /b 1
)

copy /y "%ORIGEN%" "%DESTINO%\latest.json"

cd /d "%DESTINO%"
git add latest.json
git commit -m "Ruta del dia"
git push

echo.
echo Listo. Los choferes ya ven la ruta actualizada en el mapa.
echo.
pause
